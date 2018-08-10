VERSION 5.00
Begin VB.Form frmAseguradosAccion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdProrroga 
      Caption         =   "Prorroga"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdAviso 
      Caption         =   "Aviso"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdSiniestro 
      Caption         =   "Siniestro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "Aviso falta de pago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   7095
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   0
      Left            =   2880
      Top             =   1800
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   7335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   7215
   End
End
Attribute VB_Name = "frmAseguradosAccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0.- Aviso
    '1.- Prorroga

Public SQLVto As String 'where para el vto


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Dim Aux As String

Private Sub cmdAviso_Click()
    If Text3(0).Text = "" Then Exit Sub
    Aux = "Aviso de falta de pago el: " & Text3(0).Text & vbCrLf & vbCrLf & "¿Continuar?"
    If MsgBoxA(Aux, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Aux = "UPDATE cobros SET feccomunica = '" & Format(Text3(0).Text, FormatoFecha) & "' WHERE " & SQLVto
    If Ejecuta(Aux) Then
         CadenaDesdeOtroForm = "SI"
        Unload Me
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub



Private Sub cmdProrroga_Click()
    If Text3(0).Text = "" Then Exit Sub
    Aux = "Prorrogar  el vencimiento en aseguradoras. Fecha: " & Text3(0).Text & vbCrLf & vbCrLf & "¿Continuar?"
    If MsgBoxA(Aux, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Aux = "UPDATE cobros SET fecprorroga = '" & Format(Text3(0).Text, FormatoFecha) & "' WHERE " & SQLVto
    If Ejecuta(Aux) Then
         CadenaDesdeOtroForm = "SI"
        Unload Me
    End If
End Sub

Private Sub cmdSiniestro_Click()
    If Text3(0).Text = "" Then Exit Sub
    Aux = "Seguro que desea establecer la fecha de siniestro?"
    If MsgBox(Aux, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Aux = "UPDATE cobros SET fecsiniestro = '" & Format(Text3(0).Text, FormatoFecha) & "' WHERE " & SQLVto
    If Ejecuta(Aux) Then
        CadenaDesdeOtroForm = "SI"
        Unload Me
    End If
    
End Sub

Private Sub Form_Activate()
    If Caption = "" Then
         Caption = "Operaciones aseguradas"
         'Ponerfoco Text3(0)
    End If
End Sub

Private Sub Form_Load()
   Caption = ""
    Me.Icon = frmppal.Icon
    CargaImagenesAyudas Image2, 2
    Me.Text3(0).Text = Format(Now, "dd/mm/yyyy")
    If Me.Opcion = 0 Then
      
      Me.lblTitulo.ForeColor = &H8000&
    Else
      Me.lblTitulo.ForeColor = &H80&
      
    End If
    cmdProrroga.visible = Me.Opcion = 0
    cmdAviso.visible = Me.Opcion = 0
    cmdSiniestro.visible = Me.Opcion = 1
    
End Sub


Private Sub frmC_Selec(vFecha As Date)
    Aux = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image2_Click(Index As Integer)
    Aux = ""
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text3(Index).Text <> "" Then frmC.Fecha = CDate(Text3(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    If Aux <> "" Then
        Text3(Index).Text = Aux
        Aux = ""
    End If
    
    
    
    
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    ConseguirFoco Text3(Index), 3
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index))
    If Text3(Index) = "" Then Exit Sub
    If Not EsFechaOK(Text3(Index)) Then
        MsgBox "Fecha incorrecta: " & Text3(Index), vbExclamation
        Text3(Index).Text = ""
        Text3(Index).SetFocus
    End If
End Sub


