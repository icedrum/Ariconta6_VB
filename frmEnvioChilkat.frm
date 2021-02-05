VERSION 5.00
Begin VB.Form frmEnvioChilkat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envio masivo"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "frmEnvioChilkat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private userEmail As String
Private passEmail As String
Private srvEmail As String
Private fromEmail As String

Dim PrimVez As Boolean

Private Sub Command1_Click()
    Enviar2 Nothing
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    
    PrimVez = True
    userEmail = "AKIASZJPZUOBJECOTZGK"
    passEmail = "BA5u/MSMFw3EpwR0mi3da7JUX5rAC5vhAWN2pdHRApOo"
    srvEmail = "email-smtp.eu-west-1.amazonaws.com"
    fromEmail = "david@myariadna.com"
End Sub






Private Sub Enviar2(ListaArchivos As Collection)
    Dim success
    Dim mailman As ChilkatMailMan
    Dim email As ChilkatEmail
    Dim Valores As String
    Dim J As Integer
    Dim Cad As String
    
    
    On Error GoTo GotException
    Set mailman = New ChilkatMailMan
    
    'Esta cadena es constante de la lincencia comprada a CHILKAT
    mailman.UnlockComponent "1AriadnaMAIL_BOVuuRWYpC9f"

    
    
   
    mailman.SmtpHost = srvEmail
    mailman.SmtpUsername = userEmail
    mailman.SmtpPassword = passEmail
    
    'David 2 Mayo 2007
    mailman.SmtpAuthMethod = "LOGIN"
    
    
    
    ' Create the email, add content, address it, and sent it.
    
    Set email = New ChilkatEmail
    
        
        
        
        
        email.Subject = "ASunto " & Now
        
        ' ----
        email.AddTo "Destino", "icedrum@hotmail.com"
        email.AddBcc "Envio", "david@myariadna.com"
        
        
    
    
    
    Cad = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
    Cad = Cad & "<HTML><HEAD><TITLE>Mensaje</TITLE></HEAD>"
    Cad = Cad & "<TABLE BORDER=""0"" CELLSPACING=1 CELLPADDING=0 WIDTH=576>"
    'Cuerpo del mensaje
    Cad = Cad & "<TR><TD VALIGN=""TOP""><P>"
    
    
    
    'FijarTextoMensaje
    
    
    
    
    
    Cad = Cad & "</P></TD></TR>"
    Cad = Cad & "<TR><TD VALIGN=""TOP""><P><hr></P>"
    'La imagen
    'cad = cad & "<P ALIGN=""CENTER""><IMG SRC=" & Chr(34) & "cid:" & imageContentID & Chr(34) & "></P>"
    'cad = cad & "<P ALIGN=""CENTER""><FONT SIZE=2>Mensaje creado desde el programa ARIGES de"
    'cad = cad & "<A HREF=""http://www.ariadnasoftware.com/"">Ariadna&nbsp;"
    'cad = cad & "Software S.L.</A></P><P ALIGN=""CENTER""></P>"
    Cad = Cad & "<FONT SIZE=2>"
    Cad = Cad & "<P><P><P><P align=""justify"">Este correo electrónico y sus documentos adjuntos estan dirigidos EXCLUSIVAMENTE a "
    Cad = Cad & " los destinatarios especificados. La información contenida puesde ser CONFIDENCIAL"
    Cad = Cad & " y/o estar LEGALMENTE PROTEGIDA.</P>"
    Cad = Cad & "<P align=""justify"">Si usted recibe este mensaje por ERROR, por favor comuníqueselo inmediatamente al"
    
    Cad = Cad & " remitente y ELIMINELO ya que usted NO ESTA AUTORIZADO al uso, revelación, distribución"
    Cad = Cad & " impresión o copia de toda o alguna parte de la información contenida, Gracias "
    Cad = Cad & ".</FONT></P><P><HR ALIGN=""LEFT"" SIZE=1></TD>"
    Cad = Cad & "</TR></TABLE></BODY></HTML>"
    
    email.SetHtmlBody (Cad)
    
    'Texto alternativo
    Cad = ""
    Cad = Cad & "Este correo electronico y sus documentos adjuntos estan dirigidos EXCLUSIVAMENTE a " & vbCrLf
    Cad = Cad & " los destinatarios especificados. La informacion contenida puesde ser CONFIDENCIAL" & vbCrLf
    Cad = Cad & " y/o estar LEGALMENTE PROTEGIDA." & vbCrLf & vbCrLf
    Cad = Cad & "Si usted recibe este mensaje por ERROR, por favor comuniqueselo inmediatamente al" & vbCrLf
    Cad = Cad & " remitente y ELIMINELO ya que usted NO ESTA AUTORIZADO al uso, revelacion, distribucion" & vbCrLf
    Cad = Cad & " impresion o copia de toda o alguna parte de la informacion contenida, Gracias " & vbCrLf

    
    'Por si no acepta HTML
    Cad = UCase(Cad)
    email.AddPlainTextAlternativeBody "Asunto texto plano " & vbCrLf & vbCrLf & vbCrLf & Cad
    email.From = fromEmail
    
    
    
    email.AddFileAttachment App.Path & "\docum.pdf"
'
'            For J = 1 To ListaArchivos.Count
'                   email.AddFileAttachment ListaArchivos.Item(J)
'            Next J
'        End If
    
        
    
    email.SendEncrypted = 1
    
   
    success = mailman.SendEmail(email)
   
    If (success = 1) Then
        
        
    Else
        Cad = "Han ocurrido errores durante el envio.Compruebe el archivo log.xml para mas informacion"
        mailman.SaveXmlLog App.Path & "\log.xml"
        MsgBox Cad, vbExclamation
    End If
    
    
GotException:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set email = Nothing
    Set mailman = Nothing

End Sub

