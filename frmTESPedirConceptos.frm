VERSION 5.00
Begin VB.Form frmTESPedirConceptos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos contabilización"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCLiente 
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6375
      Begin VB.CheckBox chkAgrupaContraPuente 
         Caption         =   "Agrupa importes cuenta puente"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   2760
         Width           =   3375
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   5
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   4
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtConcepto 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   1
         Text            =   "Text10"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtDConcpeto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2640
         TabIndex        =   12
         Text            =   "Text9"
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txtDescDiario 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2640
         TabIndex        =   10
         Text            =   "Text9"
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtDiario 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtDConcpeto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   7
         Text            =   "Text9"
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox txtConcepto 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   2
         Text            =   "Text10"
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Debe"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   13
         Top             =   1680
         Width           =   975
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   0
         Left            =   1560
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Diario"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Image imgDiario 
         Height          =   240
         Index           =   0
         Left            =   1560
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Eliminar efectos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   5655
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   1
         Left            =   1560
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Haber"
         Height          =   195
         Index           =   9
         Left            =   480
         TabIndex        =   8
         Top             =   2160
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmTESPedirConceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
Public Intercambio As String  'Para diversos usos


Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmDi As frmTiposDiario
Attribute frmDi.VB_VarHelpID = -1


Private PrimeraVez As Boolean
Dim SQL As String

Private Sub chkAgrupaContraPuente_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub cmdAceptar_Click(Index As Integer)
    If txtDiario(0).Text = "" Or Me.txtConcepto(0).Text = "" Or Me.txtConcepto(1).Text = "" Then
        MsgBox "Campos obligados", vbExclamation
        Exit Sub
    End If
    
    CadenaDesdeOtroForm = txtDiario(0).Text & "|" & Me.txtConcepto(0).Text & "|" & Me.txtConcepto(1).Text & "|"
    'Si agrupa o no
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Me.chkAgrupaContraPuente.Value & "|"
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 0 Then
            txtDiario(0).Text = RecuperaValor(Intercambio, 1)
            txtConcepto(0).Text = RecuperaValor(Intercambio, 2)
            txtConcepto(1).Text = RecuperaValor(Intercambio, 3)
            txtDiario_LostFocus 0
            txtConcepto_LostFocus 0
            txtConcepto_LostFocus 1
            Me.cmdAceptar(0).SetFocus
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    CargaImagenesAyudas imgConcepto, 1, "Concepto"
    CargaImagenesAyudas Me.imgDiario, 1, "Diario"
    
    Limpiar Me
    PrimeraVez = True
    'pongo los valores
    
    
    
    '
    cmdCancelar(Opcion).Cancel = True
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    SQL = RecuperaValor(CadenaDevuelta, 1)
End Sub

Private Sub imgConcepto_Click(Index As Integer)
      
    LanzaBuscaGrid 1
    If SQL <> "" Then
        txtConcepto(Index).Text = SQL
        txtConcepto_LostFocus Index
    End If
End Sub

Private Sub imgDiario_Click(Index As Integer)
    LanzaBuscaGrid 0
    If SQL <> "" Then
        txtDiario(Index).Text = SQL
        txtDiario_LostFocus Index
    End If
End Sub

Private Sub txtConcepto_GotFocus(Index As Integer)
    ConseguirFoco txtConcepto(Index), 0
End Sub

Private Sub txtConcepto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtConcepto_LostFocus(Index As Integer)
Dim I As Byte
    'Lost focus
    txtConcepto(Index).Text = Trim(txtConcepto(Index).Text)
    SQL = ""
    I = 0
    If txtConcepto(Index).Text <> "" Then
        If Not IsNumeric(txtConcepto(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            I = 1
        Else
            
            SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtConcepto(Index).Text, "N")
            If SQL = "" Then
                MsgBox "Concepto no existe", vbExclamation
                I = 1
            End If
        End If
    End If
    Me.txtDConcpeto(Index).Text = SQL
    If I = 1 Then
        txtConcepto(Index).Text = ""
        PonFoco txtConcepto(Index)
    End If
End Sub


Private Sub LanzaBuscaGrid(Opcion As Integer)

'No tocar variable SQL
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String



    SQL = ""
    Screen.MousePointer = vbHourglass
    
    'Ejemplo
    'Cod Diag.|idDiag|N|10·
    Select Case Opcion
    Case 0
        'Tipos diario
        Set frmDi = New frmTiposDiario
        frmDi.DatosADevolverBusqueda = "0|"
        frmDi.Show vbModal
        Set frmDi = Nothing
    Case 1
        'CONCEPTO
        Set frmCon = New frmConceptos
        frmCon.DatosADevolverBusqueda = "0|"
        frmCon.vWhere = " conceptos.codconce < 900 "
        frmCon.Show vbModal
        Set frmCon = Nothing
        
    End Select
       

        
    Screen.MousePointer = vbDefault
End Sub


Private Sub txtDiario_GotFocus(Index As Integer)
    ConseguirFoco txtDiario(Index), 3
End Sub

Private Sub txtDiario_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDiario_LostFocus(Index As Integer)
    
    SQL = ""
    txtDiario(Index).Text = Trim(txtDiario(Index).Text)
    If txtDiario(Index).Text <> "" Then
        
        If Not IsNumeric(txtDiario(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            txtDiario(Index).Text = ""
            PonFoco txtDiario(Index)
        Else
            txtDiario(Index).Text = Val(txtDiario(Index).Text)
            SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtDiario(Index).Text, "N")
            
            If SQL = "" Then
                MsgBox "No existe el diario: " & Me.txtDiario(Index).Text, vbExclamation
                Me.txtDiario(Index).Text = ""
                PonFoco txtDiario(Index)
            End If
        End If
    End If
    Me.txtDescDiario(Index).Text = SQL
     
End Sub
