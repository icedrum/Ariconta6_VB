VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmintegraciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "I N T E G R A C I O N E S"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "frmintegraciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Height          =   435
      Left            =   720
      Picture         =   "frmintegraciones.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar archivos"
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Height          =   435
      Left            =   120
      Picture         =   "frmintegraciones.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Guardar copia y borrar"
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmintegraciones.frx":6076
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmintegraciones.frx":C8D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   435
      Left            =   3960
      TabIndex        =   0
      Top             =   5880
      Width           =   1275
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5715
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   10081
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tipo"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Datos"
         Object.Width           =   4939
      EndProperty
   End
   Begin VB.Label Label1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   5940
      Width           =   3435
   End
End
Attribute VB_Name = "frmintegraciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TablasDeErrores As String

Dim PrimeraVez As Boolean






Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Screen.MousePointer = vbHourglass
    If EliminarArchivos(True) Then
        Command2.Visible = False
        Command3.Visible = False
        Me.Refresh
        VerErrores
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        Screen.MousePointer = vbHourglass
        PrimeraVez = False
        VerErrores
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    ListView1.ListItems.Clear
    Command2.Visible = False
    Command3.Visible = False
End Sub



Private Function VerErrores()
Dim Cad As String
Dim C1 As String
Dim I As Integer
Dim ItmX As ListItem


    On Error GoTo EVerErrores

    ListView1.ListItems.Clear
    If vConfig.Integraciones <> "" Then
    
        Cad = vConfig.Integraciones
        If Right(Cad, 1) <> "\" Then Cad = Cad & "\"
    
        If Dir(Cad, vbDirectory) = "" Then
            MsgBox "Carpeta de errores no encontrada: " & vConfig.Integraciones, vbExclamation
            Exit Function
        End If
        
        Cad = vConfig.Integraciones & "\ERRORES"
        I = 0
        'Facturas clientes
        Label1.Caption = "Fact. clientes"
        Label1.Refresh
        C1 = Dir(Cad & "\FRACLI\*.Z" & Format(vEmpresa.codempre, "00") & ".LOG")
        Do While C1 <> ""   ' Inicia el bucle.
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = "CLIENTES"
            ItmX.SubItems(1) = C1
            ItmX.SmallIcon = 1
            ItmX.Tag = Cad & "\FRACLI\" & C1
            I = I + 1
            C1 = Dir
        Loop
        
        'Facturas Proveedores
        Label1.Caption = "Fact. proveedores"
        Label1.Refresh
        C1 = Dir(Cad & "\FRAPRO\*.Z" & Format(vEmpresa.codempre, "00") & ".LOG")
        Do While C1 <> ""   ' Inicia el bucle.
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = "PROVEED"
            ItmX.SubItems(1) = C1
            ItmX.SmallIcon = 1
            ItmX.Tag = Cad & "\FRAPRO\" & C1
            I = I + 1
            C1 = Dir
        Loop
        
        'Asientos al diario
        Label1.Caption = "Asientos a diario"
        Label1.Refresh
        C1 = Dir(Cad & "\ASIDIA\*.Z" & Format(vEmpresa.codempre, "00") & ".LOG")
        Do While C1 <> ""   ' Inicia el bucle.
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = "ASIENTOS"
            ItmX.SubItems(1) = C1
            ItmX.SmallIcon = 1
            ItmX.Tag = Cad & "\ASIDIA\" & C1
            I = I + 1
            C1 = Dir
        Loop
        
        'Asientos al historico
        Label1.Caption = "Asientos historico"
        Label1.Refresh
        C1 = Dir(Cad & "\ASIHCO\*.Z" & Format(vEmpresa.codempre, "00") & ".LOG")
        Do While C1 <> ""   ' Inicia el bucle.
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = "HCO"
            ItmX.SubItems(1) = C1
            ItmX.SmallIcon = 1
            ItmX.Tag = Cad & "\ASIHCO\" & C1
            I = I + 1
            C1 = Dir
        Loop
        
    End If   'De las integracions
    
    If TablasDeErrores <> "" Then
        If InStr(1, TablasDeErrores, "CABAPU|") Then
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = "ASIENTOS"
            ItmX.SubItems(1) = "Asientos con errores"
            ItmX.SmallIcon = 2
        End If
        If InStr(1, TablasDeErrores, "FRACLI|") Then
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = "CLIENTES"
            ItmX.SubItems(1) = "Facturas clientes con error"
            ItmX.SmallIcon = 2
        End If
        If InStr(1, TablasDeErrores, "FRAPRO|") Then
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = "PROVEED"
            ItmX.SubItems(1) = "Facturas proveedores con error"
            ItmX.SmallIcon = 2
        End If
    End If
    Label1.Caption = ""
    If I > 0 Then
        'Command2.Visible = True
        'Command3.Visible = True
    End If
    
    
    Exit Function
EVerErrores:
    MuestraError Err.Number, "Imposible ver los errores."
End Function


Private Sub ListView1_DblClick()
On Error GoTo EList1
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    If ListView1.SelectedItem.SmallIcon <> 1 Then Exit Sub
    
    If ListView1.SelectedItem.Tag = "" Then Exit Sub
    
    
    'Solo sirve para los archivos
    Shell "notepad " & ListView1.SelectedItem.Tag, vbNormalFocus
            
            
    Exit Sub
EList1:
    MuestraError Err.Number, "Mostrando archivo LOG"
End Sub



Private Function EliminarArchivos(HacerPregunta As Boolean) As Boolean
Dim Cad As String
Dim I As Integer

    On Error GoTo EElim
    EliminarArchivos = False
    If ListView1.SelectedItem Is Nothing Then Exit Function
    If ListView1.SelectedItem.SmallIcon <> 1 Then Exit Function
    
    If ListView1.SelectedItem.Tag = "" Then Exit Function

    If HacerPregunta Then
        Cad = "Seguro que desea eliminar los datos relacionados con : " & ListView1.SelectedItem.Text & " - " & ListView1.SelectedItem.SubItems(1) & "?"
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If

    Cad = ListView1.SelectedItem.Tag
    I = InStr(1, Cad, ".LOG")
    If I = 0 Then
        MsgBox "Error obteniendo LOG", vbExclamation
        Exit Function
    End If
    
    Cad = Mid(Cad, 1, I - 1)
    I = InStr(1, Cad, ".Z")
    If I = 0 Then
        MsgBox "Error obteniendo Z", vbExclamation
        Exit Function
    End If
    
    Cad = Mid(Cad, 1, I - 1) & ".?" & Mid(Cad, I + 2)
    I = 1
    Kill Cad
    I = 2
    Kill ListView1.SelectedItem.Tag
    EliminarArchivos = True
    Exit Function
EElim:
    If I = 2 Then
        Cad = ListView1.SelectedItem.Tag
    Else
        If I <> 1 Then Cad = ""
    End If
    MuestraError Err.Number, "Eliminar archivos" & vbCrLf & Cad, Err.Description
End Function
