VERSION 5.00
Begin VB.Form frmEMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviar E-MAIL"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmEMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCopia 
      Caption         =   "Copia remitente"
      Height          =   255
      Left            =   960
      TabIndex        =   20
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   19
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   2940
      TabIndex        =   18
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5715
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   660
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   960
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   2055
         Index           =   3
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmEMail.frx":0442
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Para"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   300
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Asunto"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   1140
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   600
         Picture         =   "frmEMail.frx":0448
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   1560
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   60
      Width           =   5715
      Begin VB.TextBox Text3 
         Height          =   1695
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "frmEMail.frx":0E4A
         Top             =   1800
         Width           =   5355
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Otro"
         Height          =   255
         Index           =   2
         Left            =   2460
         TabIndex        =   15
         Top             =   1140
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Error"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Top             =   1140
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sugerencia"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   1140
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Mensaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   1500
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Asunto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enviar e-Mail Ariadna Software"
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
         Height          =   360
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   4305
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   300
      Picture         =   "frmEMail.frx":0E50
      Top             =   3780
      Width           =   480
   End
End
Attribute VB_Name = "frmEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As Byte
    '0 - Envio del PDF
    '1 - Envio Mail desde menu soporte
    
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private DatosDelMailEnUsuario As String
Dim Cad As String

Private Sub Enviar()
    Dim success
    Dim mailman As ChilkatMailMan
    Dim Valores2 As String
    
    On Error GoTo GotException
    Set mailman = New ChilkatMailMan
    
    'Esta cadena es constante de la lincencia comprada a CHILKAT
    mailman.UnlockComponent "1AriadnaMAIL_BOVuuRWYpC9f"
    mailman.LogMailSentFilename = ""    'App.path & "\mailSent.log"
    
    
    'Servidor smtp
    If Not vParam.EnviarPorOutlook Then
        Valores2 = DatosDelMailEnUsuario  'Empipado: diremail,smtphost,smtpuser, pass
        If Valores2 = "" Then
            MsgBox "Falta configurar en paremtros la opcion de envio mail(servidor, usuario, clave)"
            Exit Sub
        End If
    End If
    mailman.SmtpHost = RecuperaValor(Valores2, 2) ' vParam.SmtpHOST
    mailman.SmtpUsername = RecuperaValor(Valores2, 3) 'vParam.SmtpUser
    mailman.SmtpPassword = RecuperaValor(Valores2, 4) 'vParam.SmtpPass
    mailman.SmtpAuthMethod = "LOGIN"

'
    ' Create the email, add content, address it, and sent it.
    Dim email As ChilkatEmail
    Set email = New ChilkatEmail
    
    
    
    
    'Si es de SOPORTE
    If Opcion = 1 Then
         'Obtenemos la pagina web de los parametros
        Cad = DevuelveDesdeBD("mailsoporte", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
        If Cad = "" Then
            MsgBox "Falta configurar en parametros el mail de soporte", vbExclamation
            Exit Sub
        End If
    
        If Cad = "" Then GoTo GotException
        email.AddTo "Soporte Contabilidad", Cad
        
        Cad = "Soporte Ariconta. "
        If Option1(0).Value Then Cad = Cad & Option1(0).Caption
        If Option1(1).Value Then Cad = Cad & Option1(1).Caption
        If Option1(2).Value Then Cad = Cad & "Otro: " & Text2.Text
        email.Subject = Cad
        
        'Ahora en text1(3).text generaremos nuestro mensaje
        Cad = "Fecha: " & Format(Now, "dd/mm/yyyy") & vbCrLf
        Cad = Cad & "Hora: " & Format(Now, "hh:mm") & vbCrLf
        Cad = Cad & "CONTA:  " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
        Cad = Cad & "Usuario: " & vUsu.Nombre & vbCrLf
        Cad = Cad & "Nivel USU: " & vUsu.Nivel & vbCrLf
        Cad = Cad & "Empresa: " & vEmpresa.nomempre & vbCrLf
        Cad = Cad & "&nbsp;<hr>"
        Cad = Cad & Text3.Text & vbCrLf & vbCrLf
        Text1(3).Text = Cad
    Else
        'Envio de mensajes normal
        email.AddTo Text1(0).Text, Text1(1).Text
        email.Subject = Text1(2).Text
        If chkCopia.Value = 1 Then email.AddBcc "Ariconta: " & vEmpresa.nomempre, RecuperaValor(Valores2, 1)
    End If
    
    'El resto lo hacemos comun
    'La imagen
    'imageContentID = email.AddRelatedContent(App.path & "\minilogo.bmp")
    
    
    Cad = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
    Cad = Cad & "<HTML><HEAD><TITLE>Mensaje</TITLE></HEAD>"
    Cad = Cad & "<TABLE BORDER=""0"" CELLSPACING=1 CELLPADDING=0 WIDTH=576>"
    'Cuerpo del mensaje
    Cad = Cad & "<TR><TD VALIGN=""TOP""><P>"
    FijarTextoMensaje
    Cad = Cad & "</P></TD></TR>"
    Cad = Cad & "<TR><TD VALIGN=""TOP""><P><BR><BR><BR><BR><hr></P>"
    'La imagen
    'Cad = Cad & "<P ALIGN=""CENTER""><IMG SRC=" & Chr(34) & "cid:" & imageContentID & Chr(34) & "></P>"
    'Cad = Cad & "<P ALIGN=""CENTER""><FONT SIZE=2>Mensaje creado desde el programa ARICONTA de"
    'Cad = Cad & "<A HREF=""http://www.ariadnasoftware.com/"">Ariadna&nbsp;"
    'Cad = Cad & "Software S.L.</A></P><P ALIGN=""CENTER""></P>"
    Cad = Cad & "<P></P><P></P><P></P><P>Este correo electrónico y sus documentos adjuntos estan dirigidos EXCLUSIVAMENTE a "
    Cad = Cad & " los destinatarios especificados. La información contenida puesde ser CONFIDENCIAL"
    Cad = Cad & " y/o estar LEGALMENTE PROTEGIDA.</P>"
    Cad = Cad & "<P>Si usted recibe este mensaje por ERROR, por favor comuníqueselo inmediatamente al"
    Cad = Cad & " remitente y elimínelo ya que usted NO ESTA AUTORIZADO al uso, revelación, distribución"
    Cad = Cad & " impresión o copia de toda o alguna parte de la información contenida, Gracias "
    Cad = Cad & ".</FONT></P><P><HR ALIGN=""LEFT"" SIZE=1></TD>"
    Cad = Cad & "</TR></TABLE></BODY></HTML>"
    
    email.SetHtmlBody (Cad)
    


    'Texto alternativo
    Cad = "Mensaje creado desde el programa " & App.EXEName & " de Ariadna Software S.L." & vbCrLf
    Cad = Cad & "Este correo electronico y sus documentos adjuntos estan dirigidos EXCLUSIVAMENTE a " & vbCrLf
    Cad = Cad & " los destinatarios especificados. La informacion contenida puesde ser CONFIDENCIAL" & vbCrLf
    Cad = Cad & " y/o estar LEGALMENTE PROTEGIDA." & vbCrLf & vbCrLf
    Cad = Cad & "Si usted recibe este mensaje por ERROR, por favor comuniqueselo inmediatamente al" & vbCrLf
    Cad = Cad & " remitente y ELIMINELO ya que usted NO ESTA AUTORIZADO al uso, revelacion, distribucion" & vbCrLf
    Cad = Cad & " impresion o copia de toda o alguna parte de la informacion contenida, Gracias " & vbCrLf

    
    'Por si no acepta HTML
    Cad = UCase(Cad)
    email.AddPlainTextAlternativeBody Text1(3).Text & vbCrLf & vbCrLf & vbCrLf & Cad

    email.From = RecuperaValor(Valores2, 1) 'vParam.diremail
    If vUsu.Codigo = 0 Then
        email.FromName = "Administrador Contabilidad"
    Else
        email.FromName = vUsu.Nombre
    End If
    
    If Opcion = 0 Then
        'ADjunatmos el PDF
        email.AddFileAttachment App.path & "\docum.pdf"
    End If
        
    
    
    
    
    'sI ENVIA POR OUTLOOK O NO
    If vParam.EnviarPorOutlook Then
        'Si envia por outlook
         mailman.SendViaOutlook email
         success = 1
        
    Else
        success = mailman.SendEmail(email)
    End If
    If (success = 1) Then
        If vParam.EnviarPorOutlook Then
            Cad = "El mensaje se ha llevado al outlook"
        Else
            Cad = "Mensaje enviado correctamente."
        End If
        MsgBox Cad, vbInformation
        Command2(0).SetFocus
    Else
        Cad = "Han ocurrido errores durante el envio.Compruebe el archivo log.xml para mas informacion"
        mailman.SaveXmlLog App.path & "\log.xml"
        MsgBox Cad, vbExclamation
    End If
    

    
    
    
    
    
    
    
GotException:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set email = Nothing
    Set mailman = Nothing
    
End Sub



Private Sub Command1_Click()
    If Not DatosOk Then Exit Sub
    Screen.MousePointer = vbHourglass
    Image2.Visible = True
    Me.Refresh
    Enviar
    Image2.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Form_Load()
    
    Me.Icon = frmPpal.Icon

    Image2.Visible = False
    Limpiar Me
    Frame1(0).Visible = (Opcion = 0)
    Frame1(1).Visible = (Opcion = 1)
    If Opcion = 1 Then HabilitarText
'    Text1(0).Text = "Pa er Davi"
'    Text1(1).Text = vParam.diremail
'    Text1(2).Text = "Asuntillo"
'    Text1(3).Text = "Datos"
    
    If vParam.EnviarPorOutlook Then
        Cad = "||||"
    Else
        PonDisponibilidadEmail
    End If
    Me.Command1.Enabled = (Cad <> "")
End Sub




Private Sub PonDisponibilidadEmail()
    
        
    
    Cad = DevuelveDesdeBD("dirfich", "Usuarios.usuarios", "codusu", (vUsu.Codigo Mod 100), "N")
    If Cad = "" Then
        'Primero compruebo si los datos los tengo en el usuario
        Cad = "select diremail,smtphost,smtpuser,smtppass from parametros"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        If Not miRsAux.EOF Then
            If Not IsNull(miRsAux!SmtpHost) Then
                For NumRegElim = 0 To miRsAux.Fields.Count - 1
                    Cad = Cad & DBLet(miRsAux.Fields(NumRegElim), "T") & "|"
                Next NumRegElim
            End If
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        
    End If
    DatosDelMailEnUsuario = Cad
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Opcion = 0
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)

    Screen.MousePointer = vbHourglass
    Text1(0).Tag = RecuperaValor(CadenaSeleccion, 1)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 2)
    'Si regresa con datos tengo k devolveer desde la bd el campo e-mail
    Cad = DevuelveDesdeBD("maidatos", "cuentas", "codmacta", Text1(0).Tag)
    Text1(1).Text = Cad
    Screen.MousePointer = vbDefault
End Sub

Private Sub Image1_Click()
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1"
    frmC.ConfigurarBalances = 5  'NUEVO opcion
    frmC.Show vbModal
    Set frmC = Nothing
    If Text1(0).Text <> "" Then Text1(2).SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)
    HabilitarText
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Function DatosOk() As Boolean
Dim I As Integer

    DatosOk = False
    If Opcion = 0 Then
                'Pocas cosas a comprobar
                For I = 0 To 2
                    Text1(I).Text = Trim(Text1(I).Text)
                    If Text1(I).Text = "" Then
                        MsgBox "El campo: " & Label1(I).Caption & " no puede estar vacio.", vbExclamation
                        Exit Function
                    End If
                Next I
                
                'EL del mail tiene k tener la arroba @
                I = InStr(1, Text1(1).Text, "@")
                If I = 0 Then
                    MsgBox "Direccion e-mail erronea", vbExclamation
                    Exit Function
                End If
    Else
        Text2.Text = Trim(Text2.Text)
        'SOPORTE
        If Trim(Text3.Text) = "" Then
            MsgBox "El mensaje no puede ir en blanco", vbExclamation
            Exit Function
        End If
        If Option1(2).Value Then
            If Text2.Text = "" Then
                MsgBox "El campo 'OTRO asunto' no puede ir en blanco", vbExclamation
                Exit Function
            End If
        End If
    End If
      
    'Llegados aqui OK
    DatosOk = True
        
End Function


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then Exit Sub 'Si estamos en el de datos nos salimos
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

'El procedimiento servira para ir buscando los vbcrlf y cambiarlos por </p><p>
Private Sub FijarTextoMensaje()
Dim I As Integer
Dim J As Integer

    J = 1
    Do
        I = InStr(J, Text1(3).Text, vbCrLf)
        If I > 0 Then
              Cad = Cad & Mid(Text1(3).Text, J, I - J) & "</P><P>"
        Else
            Cad = Cad & Mid(Text1(3).Text, J)
        End If
        J = I + 2
    Loop Until I = 0
End Sub

Private Sub HabilitarText()
    If Option1(2).Value Then
        Text2.Enabled = True
        Text2.BackColor = vbWhite
    Else
        Text2.Enabled = False
        Text2.BackColor = &H80000018
    End If
End Sub



Private Function RecuperarDatosEMAILAriadna() As Boolean
Dim NF As Integer

    RecuperarDatosEMAILAriadna = False
    NF = FreeFile
    Open App.path & "\soporte.dat" For Input As #NF
    Line Input #NF, Cad
    Close #NF
    If Cad <> "" Then RecuperarDatosEMAILAriadna = True
    
End Function


