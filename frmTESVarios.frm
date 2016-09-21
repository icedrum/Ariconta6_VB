VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESVarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "1"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   14385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6240
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameModiRemeTal 
      Height          =   3015
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6765
      Begin VB.CommandButton cmdModRemTal 
         Caption         =   "Modificar"
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
         Left            =   4080
         TabIndex        =   13
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
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
         Index           =   25
         Left            =   5280
         TabIndex        =   14
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtDescCta 
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
         Index           =   8
         Left            =   2040
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox txtCta 
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
         Index           =   8
         Left            =   480
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   1800
         Width           =   1485
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
         Index           =   27
         Left            =   480
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label3 
         Caption         =   "Banco"
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
         Height          =   285
         Index           =   1
         Left            =   510
         TabIndex        =   27
         Top             =   1530
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
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
         Height          =   285
         Index           =   18
         Left            =   510
         TabIndex        =   26
         Top             =   780
         Width           =   840
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Modificar remesa"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   5295
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   8
         Left            =   1710
         Top             =   1530
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   27
         Left            =   1680
         Top             =   810
         Width           =   240
      End
   End
   Begin VB.Frame FrameReclamaEmail 
      Height          =   6975
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton cmdEliminarReclama 
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Eliminar"
         Top             =   6360
         Width           =   375
      End
      Begin VB.CommandButton cmdReclamas 
         Caption         =   "Continuar"
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
         Left            =   7560
         TabIndex        =   24
         Top             =   6360
         Width           =   1215
      End
      Begin VB.OptionButton optReclama 
         Caption         =   "Correctos"
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
         Left            =   8760
         TabIndex        =   22
         Top             =   450
         Width           =   1365
      End
      Begin VB.OptionButton optReclama 
         Caption         =   "Sin email"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   7230
         TabIndex        =   21
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
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
         Index           =   31
         Left            =   9000
         TabIndex        =   18
         Top             =   6360
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   5295
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   9340
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuenta"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Email"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   600
         Top             =   6360
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   4
         Left            =   240
         Top             =   6360
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   44
         Left            =   6300
         TabIndex        =   23
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label15 
         Caption         =   "Email cuentas reclamacion"
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
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Frame FrameAgregarCuentas 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton cmdInsertaCta 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         TabIndex        =   4
         Top             =   1080
         Width           =   315
      End
      Begin VB.TextBox txtDCtaNormal 
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
         Index           =   5
         Left            =   1560
         TabIndex        =   5
         Text            =   "Text9"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtCtaNormal 
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
         Index           =   5
         Left            =   120
         TabIndex        =   3
         Text            =   "Text9"
         Top             =   1080
         Width           =   1365
      End
      Begin VB.CommandButton cmdAceptarCtas 
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
         Left            =   3360
         TabIndex        =   6
         Top             =   5400
         Width           =   1095
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
         Index           =   21
         Left            =   4680
         TabIndex        =   8
         Top             =   5400
         Width           =   1095
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1560
         Width           =   5655
      End
      Begin VB.Image imgCtaNorma 
         Height          =   240
         Index           =   6
         Left            =   1230
         ToolTipText     =   "Cuentas agrupadas"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   600
         TabIndex        =   9
         Top             =   5400
         Width           =   1470
      End
      Begin VB.Image imgEliminarCta 
         Height          =   240
         Left            =   240
         Top             =   5400
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Image imgCtaNorma 
         Height          =   240
         Index           =   5
         Left            =   870
         ToolTipText     =   "Cuentas individuales"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "AGREGAR CUENTAS"
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
         Index           =   15
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmTESVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As Byte
    '
    '0 .- Pedir datos para ordenar cobros
    
    '3.- Reclamaciones
    '4.- Remesas
    
    
    '5.- Pregunta numero TALON pagare
    
    'Cambio situacion remesa
    '----------------------------
    '6.-  De A a B.   Generar banco
            
    '8.- Contabilizar remesa
        
    '9.- Devolucion remesa
        
    '10.- Mostrar vencimientos impagdos

    '11.- ACERCA DE
        
    '12  - Eliminar vtos
    
    '13.- Deuda total consolidada
    '14.-   "         ""      pero desde hasta
        
        
    '15.- Realizar transferencias
        
    '16.- Devolucion remesa desde fichero del banco
    '--------------------------------
    
    
    '17.- Eliminar informacion HCO remesas
    
    '18.- Selección de gastos para el listado de tesoreria
    
    '19.- Contabilizar gastos
    
    '20.- Seleccion de empresas disponibles, para el usuario
    
    
    '21- Listado pagos (cobros donde se indican las cuentas que quiero que apar
    
    
    'Mas sobre remesas.
    '22.- Cancelacion cliente
    '23.- Confirmacion remesa
    
    
    
    '24.- Impresion de todos los tipos de recibos
    
    '25.- Cambiar banco y/o fecha vto para la remesa de talon pagare
    
    '28 .- Devolucion remesa desde un vto
    
    
    '29 .- Recaudacion ejecutiva
    
    
    '31 .- Reclamaciones por email.
            'Tendra los que tienen email
    
    
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
Private WithEvents frmB As frmBasico
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1


Dim Rs As ADODB.Recordset
Dim Sql As String
Dim I As Integer
Dim IT As ListItem  'Comun
Dim PrimeraVez As Boolean
Dim Cancelado As Boolean
Dim CuentasCC As String





Private Sub chkAgrupadevol_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptarCtas_Click()
    If List1.ListCount = 0 Then
        MsgBox "Introduzca cuentas", vbExclamation
        Exit Sub
    End If
    
    'Cargo en CadenaDesdeOtroForm las cuentas empipadas
    CuentasCC = ""
    For I = 0 To List1.ListCount - 1
        CuentasCC = CuentasCC & Mid(List1.List(I), 1, vEmpresa.DigitosUltimoNivel) & "|"
    Next I

    CadenaDesdeOtroForm = CuentasCC
    Unload Me

End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 21 Or Index = 25 Or Index = 31 Then CadenaDesdeOtroForm = "" 'ME garantizo =""
    If Index = 31 Then
        If MsgBox("¿Cancelar el proceso?", vbQuestion + vbYesNo) = vbYes Then SubTipo = 0
    End If
    Unload Me
End Sub









Private Sub cmdEliminarReclama_Click()
    Sql = ""
    For I = 1 To Me.ListView6.ListItems.Count
        If Me.ListView6.ListItems(I).Checked Then Sql = Sql & "X"
    Next
    
    If Sql = "" Then Exit Sub
    Sql = "Desea quitar de la reclamacion las cuentas seleccionadas(" & Len(Sql) & ") ?"
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        Sql = "DELETE FROM  tmpentrefechas WHERE codUsu = " & vUsu.Codigo & " AND fechaadq = '"
        For I = Me.ListView6.ListItems.Count To 1 Step -1
            If ListView6.ListItems(I).Checked Then
                CuentasCC = Sql & ListView6.ListItems(I).Text & "'"
                Conn.Execute CuentasCC
                ListView6.ListItems.Remove I
            End If
        Next I
    End If
End Sub



Private Sub cmdInsertaCta_Click()
    
    txtCtaNormal(5).Text = Trim(txtCtaNormal(5).Text)
    If txtCtaNormal(5).Text = "" Then Exit Sub
    
    If InStr(1, CuentasCC, txtCtaNormal(5).Text & "|") > 0 Then
        MsgBox "Ya la ha añadido", vbExclamation
    Else
        CuentasCC = CuentasCC & txtCtaNormal(5).Text & "|"
        Sql = txtCtaNormal(5).Text & "      " & txtDCtaNormal(5).Text
        List1.AddItem Sql
        txtCtaNormal(5).Text = ""
        txtDCtaNormal(5).Text = ""
    End If
    PonerFoco Me.txtCtaNormal(5)
    
End Sub


Private Sub cmdModRemTal_Click()
    If Text1(27).Text = "" And Me.txtCta(8).Text = "" Then Exit Sub
    Sql = ""
    If Text1(27).Text <> "" Then Sql = Sql & vbCrLf & "Fecha: " & Text1(27).Text
    If txtCta(8).Text <> "" Then Sql = Sql & vbCrLf & "Cuenta: " & txtCta(8).Text & " " & txtDescCta(8).Text
    Sql = "Desea actualizar a los valores indicados?"
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    CadenaDesdeOtroForm = Text1(27).Text & "|" & Me.txtCta(8).Text & "|"
    Unload Me
End Sub




Private Sub cmdReclamas_Click()
    
    'Borraremos los que tienen mail erroneo
    Set Rs = New ADODB.Recordset
    Sql = "SELECT fechaadq FROM  tmpentrefechas,cuentas WHERE fechaadq=codmacta  "
    Sql = Sql & " AND codUsu = " & vUsu.Codigo & " AND "
    Sql = Sql & " coalesce(maidatos,'')='' GROUP BY fechaadq  "
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not Rs.EOF
        Sql = Sql & ", '" & Rs!fechaadq & "'"
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    If Sql <> "" Then
        Sql = "DELETE FROM  tmpentrefechas WHERE codUsu = " & vUsu.Codigo & " AND  fechaadq IN (" & Mid(Sql, 2) & ")"
        Conn.Execute Sql
    End If
        
        
    Sql = DevuelveDesdeBD("count(*)", "tmpentrefechas", "codusu", CStr(vUsu.Codigo))
    If Val(Sql) = 0 Then
        MsgBox "Ninguna reclamacion a enviar", vbExclamation
    Else
        CadenaDesdeOtroForm = "OK"
    End If
    SubTipo = 0
    
    Unload Me
End Sub



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
        Select Case Opcion
            Case 21
              
                CargalistaCuentas
                PonerFoco txtCtaNormal(5)
                
            Case 25
                PonerFoco Text1(27)
                
            Case 31
                
                ReclamacionGargarList
                If ListView6.ListItems.Count = 0 Then optReclama(1).Value = True
        End Select
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
    CargaImagenesAyudas Me.imgCtaNorma, 1, "Seleccionar cuenta"
    CargaImagenesAyudas imgCuentas, 1, "Cuenta contable banco"
    CargaImagenesAyudas Image1, 2


'    CargaImagenesAyudas ImageAyuda, 3

    Me.imgEliminarCta.Picture = frmPpal.ImaListBotoneras32.ListImages(5).Picture
    
    
    
    FrameAgregarCuentas.Visible = False
    FrameModiRemeTal.Visible = False
    FrameReclamaEmail.Visible = False
    
    Select Case Opcion
        
        
        
    Case 21
        Caption = "Seleccionar cuentas"
        FrameAgregarCuentas.Visible = True
        W = Me.FrameAgregarCuentas.Width
        H = Me.FrameAgregarCuentas.Height + 200
        
    Case 25
        Caption = "Remesas"
        FrameModiRemeTal.Visible = True
        W = Me.FrameModiRemeTal.Width
        H = Me.FrameModiRemeTal.Height + 100
    Case 31
        
        Caption = "Reclamacion"
        FrameReclamaEmail.Visible = True
        W = Me.FrameReclamaEmail.Width
        H = Me.FrameReclamaEmail.Height + 100
        SubTipo = 1 'Para que cuando le de al ASPA del forma NO cierre
        
    End Select
    
    
    Me.Height = H + 360
    Me.Width = W + 90
    
    H = Opcion
    If Opcion = 7 Then H = 6
    If Opcion = 14 Then H = 13
    If Opcion = 16 Or Opcion = 28 Then H = 9
    If Opcion = 22 Or Opcion = 23 Then H = 8
    Me.cmdCancelar(H).Cancel = True
    
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Sql = RecuperaValor(CadenaDevuelta, 1)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    If Opcion = 31 Then
        If SubTipo = 1 Then
            Cancel = 1
            Exit Sub
        End If
    End If

    If Opcion = 4 Then
        'REMESAS BANCARIAS
        If vParamT.RemesasPorEntidad Then
            If txtCta(3).Text <> txtCta(3).Tag Then LeerGuardarBancoDefectoEntidad False
        End If
        
    End If

    Set Rs = Nothing
    Set miRsAux = Nothing
        
    
    NumeroDocumento = "" 'Para reestrablecerlo siempre
End Sub


Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    I = CInt(imgCuentas(0).Tag)
    Me.txtCta(I).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescCta(I).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(Image1(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    Sql = RecuperaValor(CadenaSeleccion, 1)
End Sub


Private Sub Image1_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text1(Index).Text <> "" Then frmC.Fecha = CDate(Text1(Index).Text)
    Image1(0).Tag = Index
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

Private Sub imgCheck_Click(Index As Integer)

    If Index < 2 Then

    ElseIf Index < 4 Then
    Else
        'Reclamaciones
        If Me.optReclama(1).Value Then
            'Solo en correctos, los incorrectos se iran tooodos
            For I = 1 To Me.ListView6.ListItems.Count
                Me.ListView6.ListItems(I).Checked = Index = 5
            Next
        End If
    End If
End Sub



Private Sub imgCtaNorma_Click(Index As Integer)

        If Index <> 6 Then

               Set frmCCtas = New frmColCtas
               Sql = ""
               frmCCtas.DatosADevolverBusqueda = "0"
               frmCCtas.Show vbModal
               
               Set frmCCtas = Nothing
               If Sql <> "" Then
                   txtCtaNormal(Index).Text = Sql
                   txtCtaNormal_LostFocus Index
               End If
            
        Else
'--
'            'Para las cuentas agrupadas
'            Sql = ""
'            LanzaBuscaGrid 3
'            If Sql <> "" Then
'                If MsgBox("Va a insetar las cuentas del grupo de tesoreria: " & Sql & vbCrLf & "¿Continuar?", vbQuestion + vbYesNo) = vbYes Then
'                    Screen.MousePointer = vbHourglass
'                    Set miRsAux = New ADODB.Recordset
'                    CargaGrupo
'                    Set miRsAux = Nothing
'                    Screen.MousePointer = vbDefault
'                End If
'            End If
        End If
            
            
End Sub

Private Sub imgCuentas_Click(Index As Integer)

    imgCuentas(0).Tag = Index
    Set frmBa = New frmBanco
    frmBa.DatosADevolverBusqueda = "OK"
    frmBa.Show vbModal
    Set frmBa = Nothing
End Sub



Private Sub imgEliminarCta_Click()
    If List1.SelCount = 0 Then Exit Sub
    
    Sql = "Desea quitar la(s) cuenta(s): " & vbCrLf
    For I = 0 To List1.ListCount - 1
        If List1.Selected(I) Then Sql = Sql & List1.List(I) & vbCrLf
    Next I
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        For I = List1.ListCount - 1 To 0 Step -1
            If List1.Selected(I) Then
                Sql = Trim(Mid(List1.List(I), 1, vEmpresa.DigitosUltimoNivel + 2))
                NumRegElim = InStr(1, CuentasCC, Sql)
                If NumRegElim > 0 Then CuentasCC = Mid(CuentasCC, 1, NumRegElim - 1) & Mid(CuentasCC, NumRegElim + vEmpresa.DigitosUltimoNivel + 1) 'para que quite el pipe final
                List1.RemoveItem I
            End If
        Next I
    
    End If
    NumRegElim = 0
End Sub




Private Sub optReclama_Click(Index As Integer)
    ReclamacionGargarList
    cmdEliminarReclama.Visible = Index = 1
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



Private Sub CargaList()
    


        Sql = DevuelveDesdeBD("descformapago", "stipoformapago", "tipoformapago", CStr(SubTipo), "N")
                
        
End Sub



Private Sub txtCta_GotFocus(Index As Integer)
    ObtenerFoco txtCta(Index)
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtCta_LostFocus(Index As Integer)
Dim DevfrmCCtas As String

        txtCta(Index).Text = Trim(txtCta(Index).Text)
        DevfrmCCtas = txtCta(Index).Text
        I = 0
        If DevfrmCCtas <> "" Then
            If CuentaCorrectaUltimoNivel(DevfrmCCtas, Sql) Then
                DevfrmCCtas = DevuelveDesdeBD("codmacta", "bancos", "codmacta", DevfrmCCtas, "T")
                If DevfrmCCtas = "" Then
                    Sql = ""
                    MsgBox "La cuenta contable no esta asociada a ninguna cuenta bancaria", vbExclamation
                End If
            Else
                MsgBox Sql, vbExclamation
                DevfrmCCtas = ""
                Sql = ""
            End If
            I = 1
        Else
            Sql = ""
        End If
        
        
        txtCta(Index).Text = DevfrmCCtas
        txtDescCta(Index).Text = Sql
        If DevfrmCCtas = "" And I = 1 Then

            PonerFoco txtCta(Index)
        End If

        
End Sub



Private Function CopiarArchivo() As Boolean
On Error GoTo ECopiarArchivo

    CopiarArchivo = False
    'cd1.CancelError = True
    cd1.FileName = ""
    cd1.ShowSave
    If cd1.FileName <> "" Then
    
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo " & cd1.FileName & " ya existe" & vbCrLf & vbCrLf & "¿Sobreescribir?", vbQuestion + vbYesNo) = vbNo Then Exit Function
            Kill cd1.FileName
        End If
        'Hacemos la copia
        FileCopy Sql, cd1.FileName
        CopiarArchivo = True
    End If
    
    
    Exit Function
ECopiarArchivo:
    MuestraError Err.Number, "Copiar Archivo"
End Function







Private Sub txtCtaNormal_GotFocus(Index As Integer)
    ObtenerFoco txtCtaNormal(Index)
End Sub
    
Private Sub txtCtaNormal_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCtaNormal_LostFocus(Index As Integer)
Dim DevfrmCCtas As String
       
        DevfrmCCtas = Trim(txtCtaNormal(Index).Text)
        I = 0
        If DevfrmCCtas <> "" Then
            If CuentaCorrectaUltimoNivel(DevfrmCCtas, Sql) Then
                
            Else
                MsgBox Sql, vbExclamation
                If Index < 3 Or Index = 9 Or Index = 10 Or Index = 11 Then
                    DevfrmCCtas = ""
                    Sql = ""
                End If
            End If
            I = 1
        Else
            Sql = ""
        End If
        
        
        txtCtaNormal(Index).Text = DevfrmCCtas
        txtDCtaNormal(Index).Text = Sql
        If DevfrmCCtas = "" And I = 1 Then
            PonerFoco txtCtaNormal(Index)
        End If
        VisibleCC
    
        
        If Index = 10 Then
'            FrameDocPorveedor.Visible = False
'            If SubTipo = 2 Or SubTipo = 3 Then
'                FrameDocPorveedor.Visible = Sql <> ""
'                If Sql = "" Then
'                    txtTexto(2).Text = ""
'                    txtTexto(3).Text = ""
'                End If
'            End If
        
        End If
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


Private Sub VisibleCC()
Dim B As Boolean

    B = False
    If vParam.autocoste Then
        If txtCtaNormal(0).Text <> "" Then
                Sql = "|" & Mid(txtCtaNormal(0).Text, 1, 1) & "|"
                If InStr(1, CuentasCC, Sql) > 0 Then B = True
        End If
    End If
    Label1(14).Visible = B
End Sub


Private Sub CargalistaCuentas()
    List1.Clear
    If CadenaDesdeOtroForm <> "" Then
        Do
            I = InStr(1, CadenaDesdeOtroForm, "|")
            If I > 0 Then
                Sql = Mid(CadenaDesdeOtroForm, 1, I - 1)
                CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, I + 1)
                CuentaCorrectaUltimoNivel Sql, CuentasCC
                Sql = Sql & "      " & CuentasCC
                List1.AddItem Sql
            End If
        Loop Until I = 0
        CadenaDesdeOtroForm = ""
        
        'Genero Cuentas CC  (por no declarar mas variables vamos)
        CuentasCC = ""
        For I = 0 To List1.ListCount - 1
            Sql = Mid(List1.List(I), 1, vEmpresa.DigitosUltimoNivel)
            CuentasCC = CuentasCC & Sql & "|"
        Next I
    Else
        CuentasCC = ""
    End If
    
End Sub

Private Sub CargaGrupo()

    On Error GoTo ECargaGrupo
    
    Sql = "Select codmacta,nommacta FROM cuentas where grupotesoreria ='" & DevNombreSQL(Sql) & "'"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not miRsAux.EOF
        Sql = miRsAux!codmacta
        If InStr(1, CuentasCC, Sql & "|") > 0 Then
            I = 1
        Else
            CuentasCC = CuentasCC & Sql & "|"
            Sql = Sql & "      " & miRsAux!Nommacta
            List1.AddItem Sql
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If I > 0 Then MsgBox "Algunas cuentas YA habian sido insertadas", vbExclamation
    Exit Sub
ECargaGrupo:
    MuestraError Err.Number, "CargaGrupo"
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







Private Sub ReclamacionGargarList()
    ListView6.ListItems.Clear
    
    Sql = "SELECT fechaadq,maidatos,razosoci,nommacta FROM  tmpentrefechas,cuentas WHERE fechaadq=codmacta  "
    Sql = Sql & " AND codUsu = " & vUsu.Codigo & " AND "
    If Me.optReclama(0).Value Then
        'Sin email
        Sql = Sql & " coalesce(maidatos,'')='' "
        ListView6.Checkboxes = False
    Else
        Sql = Sql & " maidatos<>'' "
        ListView6.Checkboxes = True
    End If
    Sql = Sql & " GROUP BY fechaadq  "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Set IT = ListView6.ListItems.Add
        IT.Text = Rs!fechaadq
        IT.SubItems(1) = Rs!Nommacta
        IT.SubItems(2) = DBLet(Rs!maidatos, "T")
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

End Sub





Private Sub LeerGuardarBancoDefectoEntidad(Leer As Boolean)
On Error GoTo eLeerGuardarBancoDefectoEntidad

    I = -1
    Sql = App.Path & "\BancRemEn.xdf"
    If Leer Then
        txtCta(3).Text = ""
        If Dir(Sql, vbArchive) <> "" Then
            I = FreeFile
            Open Sql For Input As #I
            If Not EOF(I) Then
                Line Input #I, Sql
                txtCta(3).Text = Sql
                txtCta(3).Tag = Sql
            End If
        End If
    
    Else
        'Guardar
        If Me.txtCta(3).Text = "" Then
            If Dir(Sql, vbArchive) <> "" Then Kill Sql
        Else
            I = FreeFile
            Open Sql For Output As #I
            Print #I, txtCta(3).Text
            
        End If
        
        
    End If
    
    If I >= 0 Then Close #I
    Exit Sub
eLeerGuardarBancoDefectoEntidad:
    Err.Clear
End Sub
