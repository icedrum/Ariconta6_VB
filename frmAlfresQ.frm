VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form frmAlfresQ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albaranes"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   15075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   12120
      TabIndex        =   5
      Top             =   8520
      Width           =   1155
   End
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
      Left            =   13440
      TabIndex        =   6
      Top             =   8520
      Width           =   1155
   End
   Begin VB.TextBox txtNomFich 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   480
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   240
      Width           =   7695
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      DragMode        =   1  'Automatic
      Height          =   8415
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   8175
      _cx             =   5080
      _cy             =   5080
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   8880
      TabIndex        =   8
      Top             =   480
      Width           =   6015
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   240
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   4680
         Width           =   2175
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
         Left            =   2760
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   3600
         Width           =   1695
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
         Index           =   2
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "frmAlfresQ.frx":0000
         Top             =   1800
         Width           =   5535
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
         Left            =   2640
         MaxLength       =   15
         TabIndex        =   1
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
         Index           =   3
         Left            =   240
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   3600
         Width           =   2175
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
         Left            =   240
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   3600
         Picture         =   "frmAlfresQ.frx":0006
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total(Base imponible)"
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
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   4440
         Width           =   2250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   6
         Left            =   2760
         TabIndex        =   16
         Top             =   3360
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Numero"
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
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
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
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Albaran"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   3000
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NIF"
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
         Left            =   2640
         TabIndex        =   12
         Top             =   720
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1275
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   1680
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   120
      Top             =   240
      Width           =   240
   End
End
Attribute VB_Name = "frmAlfresQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Carpeta As String


Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim SQL As String

Private Sub cmdAceptar_Click()


On Error GoTo eCmdAceptar_Click

    Msg = ""
    If Me.txtNomFich.Text = "" Then Msg = "- FICHERO" & vbCrLf
    If Me.txtNomFich.Tag = "" Then Msg = "- FICHERO(2)" & vbCrLf
    For i = 0 To 5
        If Text1(i).Text = "" Then Msg = Msg & "-" & RecuperaValor("Cuenta|NIF|Nombre|NºAlb|Fecha|Importe|", i + 1) & vbCrLf
    Next i
    If Msg <> "" Then
        MsgBox "Campos obligatorios" & vbCrLf & Msg, vbExclamation
        Exit Sub
    End If
    
    Msg = ""
    If Dir(Me.txtNomFich.Tag, vbArchive) = "" Then Msg = "No existe el PDF origen: " & txtNomFich.Tag
    
    If Dir(Carpeta & "\" & txtNomFich.Text, vbArchive) <> "" Then Msg = "YA existe el PDF en el destino: " & Carpeta & "\" & txtNomFich.Text
    
    If Msg <> "" Then
        MsgBox Msg, vbExclamation
        Exit Sub
    End If
        
    If MsgBox("Desea insertar el albaran?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    If ProcesoInsAlb Then
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If
    
    
    Exit Sub
eCmdAceptar_Click:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub cmdCancelar_Click()
    If Me.txtNomFich.Tag <> "" Then
        If MsgBox("Desea cancelar el proceso?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    Unload Me
    
End Sub

Private Sub Form_Load()
    'Me.AcroPDF1.DragIcon = vbNoDrop
    'Screen.MousePointer = vbHourglass
'    AcroPDF1.LoadFile "C:\Users\David\Downloads\borrame.pdf"
'    AcroPDF1.setZoom 100
    
    Me.Icon = frmppal.Icon
    Image1(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Image1(1).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Limpiar Me
    PonerTxtTo "", ""
End Sub


Private Sub PonerTxtTo(NomFichero As String, Corto As String)

    If NomFichero = "" Then
        txtNomFich.Text = "Arraste y suelte aqui el archivo pdf o haga click en la lupa para buscarlo"
        txtNomFich.FontItalic = True
        txtNomFich.ForeColor = &H808080
        txtNomFich.Tag = ""
    Else
        txtNomFich.Text = Corto
        txtNomFich.FontItalic = False
        txtNomFich.ForeColor = vbBlack
        txtNomFich.Tag = NomFichero
    End If
End Sub

Private Sub Form_Resize()
Dim H


    If Me.WindowState = vbMinimized Then Exit Sub
    
    
    H = Me.Width - Frame1.Width - 400  '400 es el minimo
    
    If H < 0 Then
        Me.Width = Frame1.Width + 400
        H = 400
    End If
    Me.Frame1.Left = Me.Width - Me.Frame1.Width - 240
    Me.AcroPDF1.Width = H
    Me.cmdCancelar.Left = Me.Width - 420 - Me.cmdCancelar.Width
    Me.cmdAceptar.Left = cmdCancelar.Left - cmdCancelar.Width - 240
    
    H = Me.Height - 400 - 8000   '8000 es el minimo
    
    If H < 0 Then Me.Height = 8000

    
    Me.AcroPDF1.Height = Me.Height - 420 - 820
    Me.cmdCancelar.top = Me.Height - 640 - Me.cmdCancelar.Height
    cmdAceptar.top = cmdCancelar.top
    
    
    
End Sub



Private Sub frmC_Selec(vFecha As Date)
    SQL = Format(vFecha, "dd/mm/yyyy")
    
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub Image1_Click(Index As Integer)
    
    Select Case Index
    Case 0
        frmppal.cd1.FileName = ""
        frmppal.cd1.Filter = "*.pdf|*.pdf"
        frmppal.cd1.InitDir = "c:\" 'PathSalida
        frmppal.cd1.FilterIndex = 1
        frmppal.cd1.ShowOpen
        Msg = ""
        If frmppal.cd1.FileName <> "" Then
            
            If UCase(Right(frmppal.cd1.FileTitle, 4)) <> ".PDF" Then
                MsgBox "Solo acepta archivos '.pdf'", vbExclamation
                Exit Sub
            End If
            Msg = CStr(frmppal.cd1.FileName)
            SQL = CStr(frmppal.cd1.FileTitle)
            
          '  CargarPDF "", ""
            Me.txtNomFich.Text = "Leyendo fichero ..."
            Me.AcroPDF1.visible = False
            Screen.MousePointer = vbHourglass
            Me.Refresh
            CargarPDF Msg, SQL
            Me.AcroPDF1.visible = True
        End If
        
    Case 1
        SQL = ""
        Set frmCtas = New frmColCtas
        frmCtas.DatosADevolverBusqueda = "0|1|2|"
        frmCtas.ConfigurarBalances = 3  'NUEVO
        frmCtas.Show vbModal
        Set frmCtas = Nothing
        
        If SQL <> "" Then
            Text1(0).Text = RecuperaValor(SQL, 1)
            Text1_LostFocus 0
        End If
        
    Case 2
        SQL = ""
        Set frmC = New frmCal
        frmC.Fecha = Now
        If Text1(4).Text <> "" Then
            If IsDate(Text1(4).Text) Then frmC.Fecha = CDate(Text1(4).Text)
        End If
        frmC.Show vbModal
        Set frmC = Nothing
        If SQL <> "" Then Text1(4).Text = SQL
    End Select
    Me.Refresh
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtNomFich_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cad As String
Dim Nombre As String
    On Error GoTo eT

    
    
    Cad = ""
    If Data.Files Is Nothing Then
        Cad = "N"
    Else
        If Data.Files.Count <> 1 Then Cad = "N"
    End If
    If Cad <> "" Then
        MsgBox "Solo se puede cargar un fichero", vbExclamation
        Exit Sub
    End If
    
    Cad = UCase(Right(Data.Files.Item(1), 4))
    If Cad <> ".PDF" Then
        MsgBox "Solo acepta archivos '.pdf'", vbExclamation
        Exit Sub
    End If
    
    Cad = Data.Files.Item(1)
    If Dir(Cad, vbArchive) = "" Then
        MsgBox "No es una ruta válidad", vbExclamation
        Exit Sub
    End If
    
    NumRegElim = InStrRev(Cad, "\")
    If NumRegElim = 0 Then
        MsgBox "Imposible encontrar path archivo", vbExclamation
        Exit Sub
    End If
    
    Nombre = Mid(Cad, NumRegElim + 1)
    
    
    CargarPDF Cad, Nombre
        
    Exit Sub
eT:
        MsgBox Err.Description, vbExclamation
        Err.Clear
End Sub



Private Sub CargarPDF(Archivo As String, Nombre As String)

On Error GoTo EC
    Screen.MousePointer = vbHourglass
    AcroPDF1.LoadFile Archivo
    AcroPDF1.setZoom 70
    PonerTxtTo Archivo, Nombre
    
    
EC:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        PonerTxtTo "", ""
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub Text1_GotFocus(Index As Integer)
  
    ConseguirFoco Text1(Index), 3
  
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim Limpi As Boolean
Dim Id As String
    If Not PerderFocoGnral(Text1(Index), 3) Then Exit Sub
    
   

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
   
    
    Select Case Index
        
        Case 0, 1
                
                Limpi = True
                If Text1(Index).Text <> "" Then
                    Id = ""
                    If Index = 0 Then
                        If Not IsNumeric(Text1(Index).Text) Then
                            MsgBox "Cuenta debe ser numerica", vbExclamation
                        Else
                            Id = RellenaCodigoCuenta(Text1(Index).Text)
                        End If
                    Else
                        Id = Text1(Index).Text
                    End If
                    If Id <> "" Then
                        Limpi = False
                        PonerNombreCuentaNIF Index = 0, Id
                        If Text1(2).Text <> "" Then PonFoco Text1(3)
                    End If
                End If
                If Limpi Then
                    Text1(0).Text = ""
                    Text1(1).Text = ""
                    Text1(2).Text = ""
                End If
        Case 5 '
        
            If Not PonerFormatoDecimal(Text1(Index), 1) Then Text1(Index).Text = ""
        
        Case 4 ' fecha de factura
            If Text1(Index).Text <> "" Then
                If Not EsFechaOK(Text1(Index)) Then
                    MsgBox "Fecha incorrecta", vbExclamation
                    Text1(Index).Text = ""
                    PonFoco Text1(Index)
                    Exit Sub
                End If
            End If
        
    End Select
End Sub

Private Sub PonerNombreCuentaNIF(Cuenta As Boolean, Id As String)
    Msg = IIf(Cuenta, "nifdatos", "codmacta")
    Text1(2).Text = DevuelveDesdeBD("nommacta", "cuentas", IIf(Cuenta, "codmacta", "nifdatos"), Id, "T", Msg)
    If Text1(2).Text = "" Then
        MsgBox "No existe ninguna cuenta vinculada al " & IIf(Cuenta, "codigo", "NIF") & " " & Id, vbExclamation
        Text1(0).Text = ""
        Text1(1).Text = ""
    Else
        Text1(IIf(Cuenta, 1, 0)).Text = Msg
    End If
    
End Sub


Private Function ProcesoInsAlb() As Boolean


    Msg = DevuelveDesdeBD("max(id)", "factproalbaranes", "1", "1")
    i = Val(Msg) + 1
    Msg = "Insert into factproalbaranes(id,codmacta,numalbar,fechaalb,BIMponible,nombre) VALUES (" & i & ","
    Msg = Msg & DBSet(Text1(0).Text, "T") & "," & DBSet(Text1(3).Text, "T") & "," & DBSet(Text1(4).Text, "F") & ","
    Msg = Msg & DBSet(Text1(5).Text, "N") & "," & DBSet(txtNomFich.Text, "T") & ")"
    
    If Not Ejecuta(Msg) Then Exit Function
    
    'copiapmops
    On Error Resume Next
    FileCopy Me.txtNomFich.Tag, Carpeta & "\" & txtNomFich.Text
        
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        Ejecuta "DELETE FROM factproalbaranes WHERE id=" & i
    
    Else
        'Ha ido bien.
        'Si tiene parametro borramos albaran
        
        
        If vParam.EliminaPdfOriginal Then Kill txtNomFich.Tag
        If Err.Number <> 0 Then
            MsgBox "Error eliminando fichero: " & txtNomFich.Tag & vbCrLf & " Programa continua", vbExclamation
            Err.Clear
        End If
        
    End If
    ProcesoInsAlb = True
End Function
