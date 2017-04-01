VERSION 5.00
Begin VB.Form frmCuentasSEPA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir mandato sepa"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
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
      Left            =   3720
      TabIndex        =   6
      Top             =   2400
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Cancel          =   -1  'True
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
      Left            =   5040
      TabIndex        =   5
      Top             =   2400
      Width           =   1035
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   5895
   End
   Begin VB.ComboBox cmbReferencia 
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
      ItemData        =   "frmCuentasSEPA.frx":0000
      Left            =   240
      List            =   "frmCuentasSEPA.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   3315
   End
   Begin VB.Label Label7 
      Caption         =   "Banco domiciliacion"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2385
   End
   Begin VB.Label Label7 
      Caption         =   "Referencia domiciliacion"
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
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2865
   End
End
Attribute VB_Name = "frmCuentasSEPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ReferenciaSEP As String
Public FechaSEP As String
Public ctaBancoxDefecto As String
Public Cuenta As String


Private Sub cmdAceptar_Click()
 
Dim NifEmpresa As String
    
    
    If Combo1.ListIndex < 0 Then Exit Sub
    If Not PonerParamRPT("0201-02", Msg) Then Exit Sub
    cadNomRPT = Msg
     If ReferenciaSEP = "" Then
    
        If cmbReferencia.ListIndex = 1 Then
            Msg = "nifdatos"
        Else
            Msg = "iban"
        End If
        
        Msg = "UPDATE cuentas set SEPA_Refere=" & Msg & ",  SEPA_FecFirma =" & DBSet(Now, "F")
        Msg = Msg & " WHERE codmacta =" & DBSet(Cuenta, "T")
        Conn.Execute Msg
        espera 0.25
    End If
    
    
    
    NifEmpresa = DevuelveDesdeBD("nifempre", "empresa2", "1", "1")
    
    Msg = Trim(NifEmpresa) + "ES00"   'Identificacion acreedor
    Msg = CadenaTextoMod97(Msg)
    'Si no es dos digitos es un mensaje de error
    If Len(Msg) <> 2 Then
        MsgBox "Error obteniendi identificador", vbExclamation
        Exit Sub
    End If
    MsgErr = Mid(Combo1.Text, 1, 10)
    MsgErr = Trim(Replace(MsgErr, "-", " "))
    MsgErr = DevuelveDesdeBD("sufijoem", "bancos", "codmacta", MsgErr)
    Msg = "ES" & Msg & MsgErr & NifEmpresa
    MsgErr = Msg
    
    cadParam = cadParam & "|ReferPropia=""" & Msg & """|"
    numParam = numParam + 2
    
    Msg = FechaSEP
    If Msg = "" Then Msg = Format(Now, "dd/mm/yyyy")
    cadParam = cadParam & "FechaSEP=""" & Msg & """|"
    numParam = numParam + 1
    
    
    cadFormula = "{cuentas.codmacta}='" & Cuenta & "'"
    
    ImprimeGeneral
    
    
    

    
    Screen.MousePointer = vbDefault
    If ReferenciaSEP = "" Then CadenaDesdeOtroForm = "OK"
    Unload Me
End Sub

Private Sub cmdRegresar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
        
    Me.Text1.Visible = ReferenciaSEP <> ""
    'Me.Combo1.Visible = ReferenciaSEP = ""
    'Label7(0).Visible = ReferenciaSEP = ""
    
    Text1.Text = ReferenciaSEP
    Set miRsAux = New ADODB.Recordset
    Msg = "select bancos.codmacta,descripcion,nommacta from bancos,cuentas where bancos.codmacta=cuentas.codmacta and bancos.iban<>''"
    miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = -1
    Me.Combo1.Clear
    J = -1
    While Not miRsAux.EOF
        i = i + 1
        Msg = DBLet(miRsAux!Descripcion, "T")
        If Msg = "" Then Msg = miRsAux!Nommacta
        Combo1.AddItem miRsAux!codmacta & " - " & Msg
        If ctaBancoxDefecto = miRsAux!codmacta Then J = i
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If J >= 0 Then
        Combo1.ListIndex = J
    Else
        If i = 1 Then Combo1.ListIndex = 0
    End If
    
End Sub
