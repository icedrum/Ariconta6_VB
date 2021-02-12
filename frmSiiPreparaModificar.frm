VERSION 5.00
Begin VB.Form frmSiiPreparaModificar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProceo 
      Cancel          =   -1  'True
      Caption         =   "Correcta"
      Height          =   435
      Index           =   2
      Left            =   120
      TabIndex        =   26
      ToolTipText     =   "dar por buena la factura"
      Top             =   5640
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   9
      Left            =   1680
      Locked          =   -1  'True
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmSiiPreparaModificar.frx":0000
      Top             =   3120
      Width           =   6135
   End
   Begin VB.CommandButton cmdProceo 
      Caption         =   "xx"
      Height          =   435
      Index           =   1
      Left            =   5280
      TabIndex        =   1
      Top             =   5640
      Width           =   1185
   End
   Begin VB.CommandButton cmdProceo 
      Caption         =   "Cancelar"
      Height          =   435
      Index           =   0
      Left            =   6600
      TabIndex        =   2
      Top             =   5640
      Width           =   1185
   End
   Begin VB.Frame FrameInicio 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      TabIndex        =   16
      Top             =   4200
      Width           =   7575
      Begin VB.TextBox Text1 
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
         Index           =   11
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "año"
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Text1 
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
         Index           =   12
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text1 
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
         Index           =   10
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inico "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PC"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   4320
         TabIndex        =   22
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.TextBox Text1 
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
      Index           =   7
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "año"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text1 
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
      Index           =   8
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "año"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
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
      Index           =   6
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
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
      Index           =   5
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1560
      Width           =   4455
   End
   Begin VB.TextBox Text1 
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
      Index           =   4
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
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
      Index           =   3
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
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
      Index           =   2
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
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
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
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
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblInd 
      AutoSize        =   -1  'True
      Caption         =   "Inico "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1560
      TabIndex        =   25
      Top             =   5760
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   2760
      Width           =   1620
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      Caption         =   "Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   675
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   6435
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   6720
      Picture         =   "frmSiiPreparaModificar.frx":0006
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   7320
      Picture         =   "frmSiiPreparaModificar.frx":0BF1
      Stretch         =   -1  'True
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SII"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4080
      TabIndex        =   14
      Top             =   2160
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   2220
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombe"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1620
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   840
   End
End
Attribute VB_Name = "frmSiiPreparaModificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Escliente As Boolean
Public AbrirProceso As Boolean
Public where As String

Private ValorDesdeAswii As String
    'al modificar leeremos el CVSV
    'al cerrar me dira el ID de la tabla modificaSII
    
    
    
Private EsAceptadaConErrores As Boolean
Dim Cadena As String

Private Sub cmdProceo_Click(Index As Integer)
Dim B As Boolean




    If Index = 2 Then
        'Factura aceptada con errores. Dar por buena tal y como está
        B = AceptarFacturaAceptadaConErrores
    ElseIf Index = 1 Then
            If AbrirProceso Then
                B = InciarProcesoModificacion(False)
            
            Else
                B = SubirFacturaDenuevo
            End If
            lblInd.Caption = ""
            
    Else
        B = True
    End If
    If Not B Then Exit Sub
    CadenaDesdeOtroForm = "OK"
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    
    
    If Escliente Then
        Cadena = "select numserie,numfactu clave, fecfactu, '' numfactu ,anofactu,nifdatos,nommacta,totfaccl totalfac,Sii_Id from factcli "
        lblTipo.Caption = "Emitida"
        
    Else
        Cadena = "select numserie,numregis clave, fecfactu, numfactu ,anofactu,nifdatos,nommacta,totfacpr totalfac, SII_ID  from factpro "
        lblTipo.Caption = "Recibida"
        lblTipo.ForeColor = &H80&
    End If
    Cadena = Cadena & " WHERE " & where
    
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cadena, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO puede ser eof
    Text1(0).Text = miRsAux!NUmSerie
    Text1(1).Text = Format(miRsAux!Clave, "0000000")
    Text1(2).Text = miRsAux!FecFactu
    Text1(3).Text = miRsAux!numfactu
    Text1(4).Text = DBLet(miRsAux!nifdatos, "T")
    Text1(5).Text = miRsAux!Nommacta
    Text1(6).Text = miRsAux!TotalFac
    Text1(7).Text = miRsAux!Anofactu
    Text1(8).Text = miRsAux!SII_ID
    miRsAux.Close
    
    Text1(3).visible = Not Escliente
    
    
    
    
    
    Me.FrameInicio.visible = Not AbrirProceso
    Me.cmdProceo(2).visible = False
    If AbrirProceso Then
        Text1(9).Text = ""
        'Voy a ver el CSV y si el estado era correcto
        ValorDesdeAswii = "" 'será el CSV
        EsAceptadaConErrores = False
        LeerDatosAswii
    
        If ValorDesdeAswii = "" Then
            cmdProceo(1).Enabled = False
            cmdProceo(2).Enabled = False
        Else
            cmdProceo(2).visible = EsAceptadaConErrores
        End If
        
        
        
        'Voy a ir a mirar la tabla de modificacion
        Me.Caption = "Permitir modificar factura SII. "
        cmdProceo(1).Caption = "Modificar"
        
        Text1(9).Locked = False
    Else
        ValorDesdeAswii = "" 'será el CSV
        LeerDatosRegistroModificarSii
        
        If ValorDesdeAswii = "" Then cmdProceo(1).Enabled = False
        Me.Caption = "Volver a subir factura. "
        cmdProceo(1).Caption = "Subir"
    
    End If
    Set miRsAux = Nothing
    
    
    For NumRegElim = 0 To Text1.Count - 1
        If NumRegElim = 9 Then
            BloqueaTXT Text1(NumRegElim), Not AbrirProceso
        Else
            BloqueaTXT Text1(NumRegElim), True
        End If
        
    Next
    lblInd.Caption = ""
End Sub


Private Sub LeerDatosAswii()
    On Error GoTo eLeerDatosAswii
   
    
    If Escliente Then
        Cadena = "Select Enviada ,Resultado ,CSV ,Mensaje from aswsii.envio_facturas_emitidas WHERE  IDEnvioFacturasEmitidas = " & Text1(8).Text
    Else
        Cadena = "Select Enviada, Resultado ,CSV ,Mensaje,REG_IDF_IDEF_NIF from aswsii.envio_facturas_recibidas  WHERE  IDEnvioFacturasRecibidas = " & Text1(8).Text
    End If
    miRsAux.Open Cadena, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If miRsAux.EOF Then Err.Raise 513, , "No se encuentra el registro en facturas subidas "
    
    If miRsAux!enviada = 0 Then Err.Raise 513, , "Registro SII no enviado"
    If miRsAux!Resultado <> "Correcto" Then
        If miRsAux!Resultado = "AceptadoConErrores" Then
            EsAceptadaConErrores = True
            
            Cadena = DBLet(miRsAux!Mensaje, "T")
            If InStr(1, Cadena, "AEAT]") = 0 Then Cadena = "[AEAT] Aceptada con errores.   " & Cadena
            Text1(9).Text = CStr(Cadena)
            Cadena = ""
        Else
            Err.Raise 513, , "No es correcto el envio anterior"
        End If
    End If
    
    
    If DBLet(miRsAux!csv, "T") = "" Then Err.Raise 513, , "CSV anterior vacio"
          
    If Not Escliente Then
        'Si el NIF guardado NO es el mismo que el anterior NO podemos.
        'El NIF en proveedores debe ser el mismo
        
        Cadena = DevuelveDesdeBD("nifdatos", "factpro", where & " AND 1", "1")
        
        If Cadena <> miRsAux!REG_IDF_IDEF_NIF Then Err.Raise 513, , "Nifs diferentes: " & Cadena & "    Sii:" & miRsAux!REG_IDF_IDEF_NIF
        
    End If
    'ok
    ValorDesdeAswii = miRsAux!csv
    miRsAux.Close
    
eLeerDatosAswii:
    If Err.Number <> 0 Then MuestraError Err.Number, "ASWII :" & Text1(8).Text & " " & vbCrLf & Err.Description

End Sub



Private Sub LeerDatosRegistroModificarSii()
    On Error GoTo eLeerDatosAswii
    
    
    Cadena = "select * from modificarsii where estado=0 and esfacturacliente=" & IIf(Me.Escliente, 1, 0)
    Cadena = Cadena & " and numserie =" & DBSet(Text1(0).Text, "T") & " AND  factura_regis  =" & Text1(1).Text & " AND anofactu =" & Text1(7).Text
    
    miRsAux.Open Cadena, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If miRsAux.EOF Then Err.Raise 513, , "No se encuentra el registro en tabla modificar fras."
    
    Text1(10).Text = miRsAux!Usuario
    Text1(11).Text = miRsAux!PC
    Text1(12).Text = CDate(miRsAux!FechaHoraProceso)
    Text1(9).Text = DBLet(miRsAux!Observaciones, "T")
    
    
    'Ya lo esta haciendo al cargar
   ' If Not EsCliente Then
   '     If miRsAux!NIF <> Text1(3).Text Then Err.Raise 513, , "No es el mismo proveedor"
   ' End If
    
    
    
    
    'ok
    ValorDesdeAswii = miRsAux!Id
    miRsAux.Close
    
eLeerDatosAswii:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description

End Sub






Private Function SubirFacturaDenuevo() As Boolean
Dim Resultado As String

    SubirFacturaDenuevo = False
    If MsgBox("¿Desea volver a subir la factura ya modificada?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
        
        
        
    'Proceso de subida.
    ' Actualizaremos en ASWIII , pondremos el estado a 9
        
    
        
    
    
    Dim B As Boolean
    
        
    Screen.MousePointer = vbHourglass
    lblInd.Caption = "Actualizando en aswii"
    lblInd.Refresh
    If Me.Escliente Then
            B = Sii_FraCLI(Text1(0).Text, CLng(Text1(1).Text), CInt(Text1(7).Text), CLng(Text1(8).Text), Cadena, True)
    Else
        
        
            B = Sii_FraPRO(Text1(0).Text, CLng(Text1(1).Text), CInt(Text1(7).Text), CLng(Text1(8).Text), Cadena, True)
            
    End If
                
        
    If B Then
        'Ha generado bien El UPDATE
        'Ahora ahremos un checkforgi false, el sql  y true denuevo
        Ejecuta "set FOREIGN_KEY_CHECKS=0;"
        B = Ejecuta(Cadena, True)
        Ejecuta "set FOREIGN_KEY_CHECKS=1;"
        If Not B Then
            MsgBox "Error "
            B = False
            
        Else
            'OK ha ido bien. La ha marcado para actualiar
            'Esperaremos un poco AL sii
            Screen.MousePointer = vbHourglass
            lblInd.Caption = "Servicio SII "
            lblInd.Refresh
            espera 0.5
            I = 0
            B = False
            Do
                I = I + 1
                lblInd.Caption = "Comunicando: " & I
                lblInd.Refresh
                espera 1 + I
                
                Resultado = "resultado"
                If Escliente Then
                    Cadena = DevuelveDesdeBD("enviada", "aswsii.envio_facturas_emitidas", "IDEnvioFacturasEmitidas ", Text1(8).Text, "N", Resultado)
                Else
                    Cadena = DevuelveDesdeBD("enviada", "aswsii.envio_facturas_recibidas", "IDEnvioFacturasRecibidas ", Text1(8).Text, "N", Resultado)
                End If
                If Val(Cadena) = 1 Then
                    I = 100
                    If Resultado = "Correcto" Then
                        'Perfecto. Ha ido todo de puta madre
                        B = True
                    Else
                        Cadena = "Error subiendo al portal AEAT-SII."
                    End If
                    I = 100
                Else
                    'No se ha enviado todavia
                   Cadena = "No se ha enviado todavia. "
                End If
                DoEvent2
            Loop Until I > 7
            
            If B Then
                'Todo ok
                ACtualizaFacturaEstado
            Else
                MsgBox Cadena & vbCrLf & vbCrLf & " Avise soporte técnico", vbExclamation
            End If
            
            SubirFacturaDenuevo = True 'que salga de aqui
            
            
        End If
    End If
    
    Screen.MousePointer = vbDefault
End Function


Private Function ACtualizaFacturaEstado()
    'Vamos a actualizar SII_Estado de la factura y vamos a cerrar en modifcarSII
    Cadena = "UPDATE  modificarsii SET estado=1 , FechaHoraCierreProceso =" & DBSet(Now, "FH")
    Cadena = Cadena & " WHERE ID = " & ValorDesdeAswii
    If Ejecuta(Cadena) Then
        'ACtualizamos en factura poniendo estatus el que 8
        Cadena = IIf(Me.Escliente, "factcli", "factpro")
        lblInd.Caption = "Actualiza tabla " & Cadena
        lblInd.Refresh
        Cadena = "UPDATE " & Cadena & " SET sii_estado=9 WHERE " & Me.where
        If Not Ejecuta(Cadena) Then
            MsgBox "Avise a soporte técnico", vbCritical
      
        End If
    Else
        
    End If
    
    
End Function

Private Function InciarProcesoModificacion(DeAceptadaConErres As Boolean) As Boolean
Dim Sql As String


    InciarProcesoModificacion = False

    'Inciar proceso consiste en:
    '       1) insertar en tabla modificar SII, comprobando que no hay ninguna factura igual
    '       2) Updatear SII_estado a 8
    
    Sql = "EsFacturaCliente =" & IIf(Me.Escliente, 1, 0) & " AND  numserie  =" & DBSet(Text1(0).Text, "T")
    Sql = Sql & " AND  factura_regis =" & Text1(1).Text & " AND  anofactu =" & Text1(8).Text & " AND     Estado = 0"
    Sql = DevuelveDesdeBD("count(*)", "modificarsii", Sql & " AND 1", "1")
    If Val(Sql) > 0 Then
        MsgBox "Error. Ya existe esta factura en proceso de modificacion", vbExclamation
        Exit Function
    End If
    
    'Si es de aceptada con errores  (DeAceptadaConErres)
    '       hara la insertcion en modificassi - pero con el status=1
    '       y SII_estado de la factura será 9  y en aswii_factura pondrá un Correcto con la mensaje=Aceptaconerrores  -->=Ok
    '       no hace pregunta. YA la ha hecho
        
    If Not DeAceptadaConErres Then
        If MsgBox("Seguro que desea poder modificar la factura presentada al SII?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    
    'modificarsii (FechaHoraProceso,usuario,PC,estado,SII_ID,CSV_anterior,EsFacturaCliente,numserie,factura_regis,fecfactu,anofactu,nif,Observaciones)
    lblInd.Caption = "Inserta modificarsii"
    lblInd.Refresh
    Sql = "INSERT INTO modificarsii (FechaHoraProceso,usuario,PC,estado,SII_ID,CSV_anterior,EsFacturaCliente,numserie,factura_regis,fecfactu,anofactu,nif,Observaciones"
    If DeAceptadaConErres Then Sql = Sql & ",FechaHoraCierreProceso"
    Sql = Sql & ") VALUES ("
    Sql = Sql & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(vUsu.PC, "T") & ","
    Sql = Sql & IIf(DeAceptadaConErres, "1", "0")
    Sql = Sql & "," & Text1(8).Text & "," & DBSet(ValorDesdeAswii, "T") & ","
    'EsFacturaCliente,numserie,factura_regis,fecfactu,
    Sql = Sql & IIf(Escliente, 1, 0) & ",'" & Text1(0).Text & "'," & Text1(1).Text & "," & DBSet(Text1(2).Text, "F") & ","
    ' anofactu,nif,Observaciones
    Sql = Sql & Text1(7).Text & "," & DBSet(Text1(4).Text, "T") & "," & DBSet(Text1(9).Text, "F")
    If DeAceptadaConErres Then Sql = Sql & "," & DBSet(Now, "FH")
    Sql = Sql & ")"
    
    
    If Ejecuta(Sql) Then
        'ACtualizamos en factura poniendo estatus el que 8
        
        
        Sql = IIf(Me.Escliente, "factcli", "factpro")
        lblInd.Caption = "Actualiza tabla " & Sql
        lblInd.Refresh
        Sql = "UPDATE " & Sql & " SET sii_estado= " & IIf(DeAceptadaConErres, "9", "8") & " WHERE " & Me.where
        If Not Ejecuta(Sql) Then
            MsgBox "Avise a soporte técnico", vbCritical
        Else
            InciarProcesoModificacion = True
        End If
    End If
    
    If InciarProcesoModificacion Then
        
        If Me.cmdProceo(2).visible Then
            'si el boton esta visible es que
            'Venga de aceptada(es decir que la da por buena) o no (que la quiere modificar)
            ' en aswiii tendremos que decir que es correcta
            Cadena = " SET Resultado= 'Correcto', mensaje =concat('[Aceptada con errores] ',coalesce(mensaje,''))"
            'Updateamos en aswii.sii
            If Escliente Then
                Cadena = "aswsii.envio_facturas_emitidas " & Cadena & " WHERE  IDEnvioFacturasEmitidas"
            Else
                Cadena = "aswsii.envio_facturas_recibidas " & Cadena & " WHERE  IDEnvioFacturasRecibidas"
            End If
            Cadena = "UPDATE " & Cadena & " = " & Text1(8).Text
            If Not Ejecuta(Cadena) Then MsgBox "Actualizando aceptadacon errores en aswiii. Avise a soporte técnico", vbCritical
        End If
    End If
        
    
End Function



Private Function AceptarFacturaAceptadaConErrores() As Boolean
    
    AceptarFacturaAceptadaConErrores = False
    
    Cadena = "Factura ACEPTADA CON ERRORES" & vbCrLf & vbCrLf
    Cadena = "Va a darla por correcta en contabilidad.  " & vbCrLf & Space(10) & "¿Continuar?"
    
    If MsgBox(Cadena, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    
        
        
    If InciarProcesoModificacion(True) Then AceptarFacturaAceptadaConErrores = True
        
End Function
