VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMemoria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurador memoria"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "frmMemoria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameActualizar 
      BorderStyle     =   0  'None
      Height          =   4035
      Left            =   1800
      TabIndex        =   7
      Top             =   1500
      Width           =   6435
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Aceptar"
         Height          =   435
         Index           =   0
         Left            =   3780
         TabIndex        =   13
         Top             =   3180
         Width           =   1155
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Cancelar"
         Height          =   435
         Index           =   1
         Left            =   5040
         TabIndex        =   15
         Top             =   3180
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   120
         MaxLength       =   100
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1920
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FEF7E4&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1080
         Width           =   915
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1080
         Width           =   1995
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   120
         MaxLength       =   100
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2640
         Width           =   6015
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   3915
         Left            =   60
         Top             =   60
         Width           =   6315
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   5895
      End
      Begin VB.Label Label4 
         Caption         =   "Descripción"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Código"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Valor"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo"
         Height          =   195
         Index           =   3
         Left            =   1320
         TabIndex        =   12
         Top             =   840
         Width           =   1155
      End
   End
   Begin VB.Frame FrameGenerador 
      Height          =   7455
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9435
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmMemoria.frx":000C
         Left            =   4620
         List            =   "frmMemoria.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   7020
         Width           =   2235
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   8280
         TabIndex        =   23
         Top             =   6960
         Width           =   975
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar"
         Height          =   375
         Index           =   0
         Left            =   7200
         TabIndex        =   22
         Top             =   6960
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   6375
         Left            =   180
         TabIndex        =   20
         Top             =   480
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   11245
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Datos para la memoria"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.Frame FramePpal 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9435
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8820
         Top             =   60
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
               Picture         =   "frmMemoria.frx":002D
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMemoria.frx":0347
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   8100
         TabIndex        =   5
         Top             =   6960
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Listado"
         Height          =   375
         Index           =   0
         Left            =   6960
         TabIndex        =   4
         Top             =   6960
         Width           =   1035
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1635
         Left            =   180
         TabIndex        =   3
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   2884
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7673
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Valor"
            Object.Width           =   6350
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4395
         Left            =   180
         TabIndex        =   6
         Top             =   2460
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   7752
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7673
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Valor"
            Object.Width           =   6350
         EndProperty
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   2
         Left            =   2160
         Picture         =   "frmMemoria.frx":0661
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   1
         Left            =   1800
         Picture         =   "frmMemoria.frx":0763
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   1500
         Picture         =   "frmMemoria.frx":0865
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   2460
         Picture         =   "frmMemoria.frx":0967
         Top             =   180
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   2100
         Picture         =   "frmMemoria.frx":0A69
         Top             =   180
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   1740
         Picture         =   "frmMemoria.frx":0B6B
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fórmulas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   180
         TabIndex        =   2
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Parametros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmMemoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public opcion As Byte
'           0.-  Configurador
'           1.-  Generador de los datos para luego mandar a cambiar valores




'-------------------------------------------------------------------
' Los parametros son textos. Seran los MinParam
' tiene un valor que sera ofertado. pero se podra modificar
' Las formulas son o bien ctas o bien formaulas de sumas entre ellas
'-------------------------------------------------------------------



Private Const MinParam = 3  'parametros k vienen "de serie"
Dim PrimeraVez As Boolean
Dim ItmX As ListItem
Dim Rs As ADODB.Recordset





Private Sub cmdGenerar_Click(Index As Integer)
Dim Cad As String
Dim T1 As Single
Dim Salir As Boolean
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    On Error GoTo EGen
    
    
    'Comprobamos que esta el archivo para ser llamado
    If Dir(App.Path & "\A_word.exe") = "" Then
        MsgBox "No existe el programa para generar la memoria", vbExclamation
        Exit Sub
    End If
    
    If Dir(App.Path & "\memoria", vbDirectory) = "" Then
        MsgBox "No existe la carpeta con los .DAT para la generacion de la memoria.", vbExclamation
        Exit Sub
    End If
    
    If Dir(App.Path & "\Memoria.doc") <> "" Then Kill App.Path & "\Memoria.doc"
    
    
    'Como llamar al programa
    'Cad = "/F:C:\Programas\WORD\memoria\normal.dat"
    'Cad = Cad & " /U:1003 /P:aritel"
    
    Cad = "/F:" & App.Path & "\memoria\"
    If Combo2.ListIndex = 1 Then
        Cad = Cad & "Abreviad.dat"
    Else
        Cad = Cad & "Normal.dat"
    End If
    
    Cad = Cad & " /U:" & vUsu.Codigo
'-- vconfig comentado
'    Cad = Cad & " /P:" & vConfig.password
    
    'Ejecutamos
    Screen.MousePointer = vbHourglass
    Cad = App.Path & "\A_word.exe " & Cad
    Shell Cad, vbNormalFocus
    espera 2
    T1 = Timer
    Salir = False
    Cad = App.Path & "\Semaforo.flg"
    Do
        If Dir(Cad) = "" Then Salir = True
        If (Timer - T1 > 60) Then Salir = True
    Loop Until Salir
    
    Cad = "Se ha generado la memoria." & vbCrLf & vbCrLf
    Cad = Cad & App.Path & "\memoria.doc"
    If Dir(App.Path & "\Memoria.doc") <> "" Then MsgBox Cad, vbInformation
        
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdNuevo_Click(Index As Integer)
    If Index = 1 Then
        PreparaFrameAnyadir False
        Me.Refresh
        Exit Sub
    End If
    
    'Combo tiene k tener valor
    If Combo1.ListIndex < 0 Then
        MsgBox "Selecciione un tipo de operacion", vbExclamation
        Exit Sub
    End If
        
    Text1(0).Text = Trim(Text1(0).Text)
    Text1(2).Text = Trim(Text1(2).Text)
    If Text1(0).Text = "" Or Text1(2).Text = "" Then
        If MsgBox("Deberia poner valor para todas las opciones. ¿Desea continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    'Comprobaremos los valores
    If Not DatosOK Then Exit Sub
    
    'Ahora o modificaremos, o insertaremos
    If Label3.Tag = "N" Then
        If InsertarNuevo Then
            'Cargamos y buscamos el nuevo
            Screen.MousePointer = vbHourglass
            CargaList (Not Combo1.Enabled)
            'Lo seleccionamos
            If Combo1.Enabled Then
                'Es FORMUALAS
                With ListView2
                    For Index = 1 To .ListItems.Count
                        If .ListItems(Index).Text = Text1(1).Text Then
                            .SelectedItem = .ListItems(Index)
                            .SelectedItem.EnsureVisible
                            Exit For
                        End If
                    Next Index
                End With
            Else
                'Es parametros
                With ListView1
                    For Index = 1 To .ListItems.Count
                        If .ListItems(Index).Text = Text1(1).Text Then
                            .SelectedItem = .ListItems(Index)
                            .SelectedItem.EnsureVisible
                            Exit For
                        End If
                    Next Index
                End With
            End If
            
            
            
            
            Screen.MousePointer = vbDefault
        Else
            Exit Sub
        End If
    Else
        If ModificarValores Then
            'Los cambiamos en el itm
            If Combo1.Enabled Then
                'Es FORMUALAS
                With ListView2.SelectedItem
                    .SubItems(1) = Text1(0).Text
                    .SubItems(2) = Combo1.Text
                    .SubItems(3) = Text1(2).Text
                End With
            Else
                'Es parametros
                With ListView1.SelectedItem
                    '.SubItems(1) = Text1(0).Text
                    '.SubItems(2) = Combo1.Text
                    .SubItems(3) = Text1(2).Text
                End With
            End If
            Me.Refresh
        Else
            Exit Sub
        End If
    End If
    PreparaFrameAnyadir False
    Me.Refresh
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Cad As String
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Conn.Execute "Delete from usuarios.zmemoria where codusu = " & vUsu.Codigo
    Cad = " INSERT INTO Usuarios.zmemoria "   '(codusu, codigo, parame, descripcion, valortexto, valornumero) VALUES ( "
    Cad = Cad & " SELECT " & vUsu.Codigo & ",memoria.codigo,parametros,memoria.valor,"
    Cad = Cad & " memoria.descripcion,tipomemoria.descripcion,NULL"
    Cad = Cad & " FROM memoria,tipomemoria where  memoria.tipo=tipomemoria.codigo "
    Conn.Execute Cad
    With frmImprimir
        .OtrosParametros = ""
        .NumeroParametros = 0
        .FormulaSeleccion = "{zmemoria.codusu}=" & vUsu.Codigo
        .SoloImprimir = False
        'Opcion dependera del combo
        .opcion = 55
        .Show vbModal
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        Screen.MousePointer = vbHourglass
        PrimeraVez = False
        If opcion = 0 Then
            CargaList False
            CargaList True
        Else
            Combo2.ListIndex = 0
            'Calcular datos
            RealizarCalculoDatos
            CargaDatos
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon

    FrameActualizar.Visible = False
    FrameGenerador.Visible = opcion = 1
    Me.FramePpal.Visible = opcion = 0
    If opcion = 0 Then
        Caption = "Configurador memoria ejercicio"
        CargaCombo
        ListView1.SmallIcons = Me.ImageList1
        ListView2.SmallIcons = Me.ImageList1
    Else
        Caption = "Generador memoria ejercicio"
        ListView3.SmallIcons = Me.ImageList1
    End If
    PrimeraVez = True
End Sub

Private Sub CargaCombo()
    Set Rs = New ADODB.Recordset
    On Error GoTo ECarga
    Rs.Open "Select * from tipomemoria order by Descripcion", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Combo1.Clear
    While Not Rs.EOF
        Combo1.AddItem Rs!Descripcion
        Combo1.ItemData(Combo1.NewIndex) = Rs!Codigo
        Rs.MoveNext
    Wend
    Rs.Close
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando combo"
    Set Rs = Nothing
End Sub


Private Sub CargaList(Parametros As Boolean)
Dim Cad As String
Dim i As Integer
    Set Rs = New ADODB.Recordset
    Cad = "Select memoria.codigo,memoria.valor,memoria.descripcion,"
    Cad = Cad & "tipomemoria.descripcion,memoria.tipo FROM memoria,tipomemoria where "
    Cad = Cad & " memoria.tipo=tipomemoria.codigo"
    'Solo los de parametros
    Cad = Cad & " AND parametros="
    If Parametros Then
        Cad = Cad & "1"
        i = 1
        ListView1.ListItems.Clear
    Else
        Cad = Cad & "0"
        ListView2.ListItems.Clear
        i = 2
    End If
    Cad = Cad & " ORDER BY memoria.codigo"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        If Parametros Then
            Set ItmX = ListView1.ListItems.Add(, , Rs!Codigo)
        Else
            Set ItmX = ListView2.ListItems.Add(, , Rs!Codigo)
        End If
        ItmX.SubItems(1) = Rs.Fields(2)
        ItmX.SubItems(2) = Rs.Fields(3)
        ItmX.SubItems(3) = Rs.Fields(1)
        'En el tag ponemos el tipo
        ItmX.Tag = Rs!Tipo
        ItmX.SmallIcon = i
        Rs.MoveNext
    Wend
    Rs.Close
End Sub

Private Sub PreparaFrameAnyadir(Vsible As Boolean)
    Command1(0).Enabled = Not Vsible
    Command1(1).Enabled = Not Vsible
    FrameActualizar.Visible = Vsible
    FrameActualizar.Enabled = Vsible
    FramePpal.Enabled = Not Vsible
End Sub

Private Sub Nuevo(Parametro As Boolean)
    'El label
    Label3.Tag = "N"
    Label3.Caption = "NUEVO "
    If Parametro Then
        Label3.Caption = Label3.Caption & " parametro"
        Label3.ForeColor = &H800000
        'El como NO se toca
        Combo1.ListIndex = 4
        Combo1.Enabled = False
    Else
        Combo1.ListIndex = -1
        Label3.ForeColor = &H80&
        Combo1.Enabled = True
    End If
    ObtenerCodigo Parametro
    Text1(0).Text = ""
    Text1(2).Text = ""
    Text1(0).Enabled = True
    Text1(0).BackColor = &H80000005
    PreparaFrameAnyadir True
    FrameActualizar.Tag = Parametro
    Text1(0).SetFocus
End Sub

        'En cadena pasamos, codigo ,descrip, valor, tipo
Private Sub Modificar(Parametro As Boolean, ByRef Itm As ListItem)
Dim i As Integer
    'El label
    Label3.Tag = "M"
    Label3.Caption = "MODIFICAR "
    If Parametro Then
        Label3.Caption = Label3.Caption & " parametros"
        Label3.ForeColor = &H800000
        Combo1.Enabled = False
    Else
        Label3.ForeColor = &H80&
        Combo1.Enabled = True
    End If
    'Resto valores
    Text1(1).Text = Itm.Text
    Text1(0).Text = Itm.SubItems(1)
    Text1(2).Text = Itm.SubItems(3)
    Combo1.ListIndex = -1
    For i = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(i) = Itm.Tag Then
            'Este es el nodo
            Combo1.ListIndex = i
            Exit For
        End If
    Next i
    'Si es parametro, y no es nuevo, entocnces no puede modificar el texto
    Text1(0).Enabled = True
    Text1(0).BackColor = &H80000005
    If Parametro Then
        If MinParam >= Val(Text1(1).Text) Then
            Text1(0).Enabled = False
            Text1(0).BackColor = &H80000018   '
        End If
    End If
    PreparaFrameAnyadir True
End Sub


Private Function Eliminar(Parametros As Boolean, ByRef Itm As ListItem) As Boolean
Dim Cad As String
    On Error GoTo EElim
    Eliminar = False
    If Parametros Then
        If Val(Itm.Text) <= MinParam Then
            MsgBox "Los parametros menores de " & MinParam & " no pueden ser eliminados", vbExclamation
            Exit Function
        End If
    End If
    Cad = "¿Seguro que desea eliminar el nodo " & Itm.Text & " - " & Itm.SubItems(1) & "?"
    If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
        Cad = "Delete from memoria where codigo =" & Itm.Text
        Cad = Cad & " AND parametros = "
        If Parametros Then
            Cad = Cad & "1"
        Else
            Cad = Cad & "0"
        End If
        Conn.Execute Cad
        Eliminar = True
    End If
    Exit Function
EElim:
    MuestraError Err.Number, "Eliminar" & Err.Description
End Function




Private Sub ObtenerCodigo(EsParametro As Boolean)
Dim i As Integer
Dim Cad As String
        
    Cad = "Select max(codigo) from memoria where parametros="
    If EsParametro Then
        Cad = Cad & "1"
    Else
        Cad = Cad & "0"
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 1
    If Not Rs.EOF Then i = DBLet(Rs.Fields(0), "N") + 1
    Rs.Close
    Set Rs = Nothing
    'Lo ponemos en el text
    Text1(1).Text = i
End Sub

Private Sub Image1_Click(Index As Integer)
    If Index > 0 Then
        'Modificar eliminar
        If ListView1.ListItems.Count = 0 Then
            MsgBox "No hay ningun nodo para modificar/eliminar", vbExclamation
            Exit Sub
        End If
        
        If ListView1.SelectedItem Is Nothing Then
            MsgBox "Seleccione un nodo", vbExclamation
            Exit Sub
        End If
    End If
    Select Case Index
    Case 0
        Nuevo True
    Case 1
        Modificar True, ListView1.SelectedItem
    Case 2
        If Eliminar(True, ListView1.SelectedItem) Then
            'Lo eliminamos de aqui tb
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
        End If
    End Select
End Sub

Private Sub Image2_Click(Index As Integer)
   If Index > 0 Then
        'Modificar eliminar
        If ListView2.ListItems.Count = 0 Then
            MsgBox "No hay ningun nodo para modificar/eliminar", vbExclamation
            Exit Sub
        End If
        
        If ListView2.SelectedItem Is Nothing Then
            MsgBox "Seleccione un nodo", vbExclamation
            Exit Sub
        End If
    End If
    Select Case Index
    Case 0
        Nuevo False
    Case 1
        Modificar False, ListView2.SelectedItem
    Case 2
        If Eliminar(False, ListView2.SelectedItem) Then
            'Lo eliminamos de aqui tb
            ListView2.ListItems.Remove ListView2.SelectedItem.Index
        End If
    End Select
End Sub



Private Function DatosOK() As Boolean
    DatosOK = False
    Select Case Combo1.ListIndex
    Case 0
        'ASiento de apertura
        If CadenaCorrecta(False) Then DatosOK = True
            
    Case 1
        'Campo calculado. Es una formula
        If CadenaCorrecta(True) Then
           If ComprobarFormula Then DatosOK = True
        End If
        
    Case 2
        'Ejercicio actual
        If CadenaCorrecta(False) Then DatosOK = True
        
        
    Case 3
        'Ejercicio siguiente
        If CadenaCorrecta(False) Then DatosOK = True
    Case 4
        'Textos
        'siempre Correcto
        DatosOK = True
    End Select
End Function


Private Function CadenaCorrecta(EsFormula As Boolean) As Boolean
Dim i As Integer
    CadenaCorrecta = False
    If Len(Text1(2).Text) > 0 Then
        For i = 1 To Len(Text1(2).Text)
            Select Case Mid(Text1(2).Text, i, 1)
            Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
                'Correcto
'            Case "(", ")"
'                'correcto
            Case "+", "-"  '"/", "*"
                'correcto
            Case " "
                'Correcto
            Case "$"
                If EsFormula Then
                    'correcto
                Else
                    MsgBox "No se puede referenciar parametros y formulas entre si", vbExclamation
                    Exit Function
                End If
                
            Case Else
                MsgBox "Caracater incorrecto: " & Mid(Text1(2).Text, i, 1), vbExclamation
                Exit Function
            End Select
        Next i
        CadenaCorrecta = True
    Else
        CadenaCorrecta = True
    End If
End Function


Private Function ComprobarFormula() As Boolean
Dim Cad As String
Dim AUx As String
Dim i As Integer
Dim SiguienteDollar As Boolean
Dim Prov As String
Dim J As Integer

    Cad = Text1(2).Text
    Cad = Trim(Cad)
    ComprobarFormula = True
    If Cad = "" Then Exit Function
    
    ComprobarFormula = False
    
    'Quitamos los espacios en blanco
    AUx = ""
    For i = 1 To Len(Cad)
        If Mid(Cad, i, 1) <> " " Then AUx = AUx & Mid(Cad, i, 1)
    Next i
    Cad = AUx
    
    
    'QUitamos los parentesis, si los tiene
    AUx = ""
    For i = 1 To Len(Cad)
        If Mid(Cad, i, 1) <> "(" And Mid(Cad, i, 1) <> ")" Then AUx = AUx & Mid(Cad, i, 1)
    Next i
    Cad = AUx
    
    'Vemos si empiza por numero
    AUx = Mid(Cad, 1, 1)
    i = Asc(AUx)
    If i >= 48 Or i <= 57 Then
        'Es numero
        Cad = "+" & Cad
    End If
    
    'Vamos por grupos. Es decir tendremos la
    'cadena de esta forma
    '   +$12+$18*$3
    '
    ' Yo busco signo de operacion. Nada mas lo encuentre, lo siguiente debe ser el $
    SiguienteDollar = False
    For i = 1 To Len(Cad)
        AUx = Mid(Cad, i, 1)
        'Case "+", "*", "/", "-"
        If AUx = "+" Or AUx = "*" Or AUx = "/" Or AUx = "-" Then
            SiguienteDollar = True
        Else
            If SiguienteDollar Then
                If AUx = "$" Then
                    'Correcto
                    SiguienteDollar = False
                Else
                    MsgBox "No se pude mezclar referencias a cuentas con formulas", vbExclamation
                    Exit Function
                End If
            End If
        End If
    Next i
    
    'Comprobamos la referencia
    SiguienteDollar = False
    Prov = ""
    For i = 1 To Len(Cad)
        AUx = Mid(Cad, i, 1)
        'Case "+", "*", "/", "-"
        If AUx = "+" Or AUx = "*" Or AUx = "/" Or AUx = "-" Then
            If Prov <> "" Then
                'Comprobamos k exite la referencia
                If Not ExisteReferencia(Prov) Then Exit Function
            End If
            
            'Reinicializamos para que siga comprobando
            SiguienteDollar = True
            Prov = ""
        Else
            If SiguienteDollar Then
                If AUx = "$" Then
                    'Correcto
                    SiguienteDollar = False
                Else
                    MsgBox "No se pude mezclar referencias a cuentas con formulas", vbExclamation
                    Exit Function
                End If
            Else
                Prov = Prov & AUx
            End If
        End If
    Next i
    'El ultimo bloque
    If Prov <> "" Then
        If Not ExisteReferencia(Prov) Then Exit Function
    End If
    
    '---------------------------------
    'Comprobamos la parentizacion
    'Buscamos (
    J = 0
    For i = 1 To Len(Cad)
        If Mid(Cad, i, 1) = "(" Then J = J + 1
    Next i
    Prov = CStr(J)
    'Buscamos )
    J = 0
    For i = 1 To Len(Cad)
        If Mid(Cad, i, 1) = ")" Then J = J + 1
    Next i
    If Val(Prov) <> J Then
        MsgBox "No coincide el numero de parentesis"
        Exit Function
    End If
    
    
    ComprobarFormula = True
End Function


Private Function ExisteReferencia(ByRef Cad As String) As Boolean
Dim J As Integer
Dim B As Boolean

    ExisteReferencia = False
    B = False
    For J = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(J) = Cad Then
            B = True
            Exit For
        End If
    Next J
    If B = False Then
        MsgBox "No se ha encontrado la formula $" & Cad, vbExclamation
        Exit Function
    Else
        ExisteReferencia = True
    End If

End Function



Private Function ModificarValores() As Boolean
Dim Cad As String

    On Error GoTo EmodificarValores
    ModificarValores = False
    
    Cad = "UPDATE memoria SET "
    'VALOR
    Cad = Cad & " Valor='" & Text1(2).Text & "'"
    'DESCRIPCION
    Cad = Cad & ", Descripcion='" & Text1(0).Text & "'"
    'Tipo
    Cad = Cad & ", Tipo=" & Combo1.ItemData(Combo1.ListIndex)
    
    Cad = Cad & " WHERE Codigo = " & Text1(1).Text & " And parametros = "
    'Si es paramtro
    If Combo1.Enabled = False Then
        Cad = Cad & "1"
    Else
        Cad = Cad & "0"
    End If
    Conn.Execute Cad
    
    ModificarValores = True
    Exit Function
EmodificarValores:
    MuestraError Err.Number, Err.Description
    
End Function


Private Function InsertarNuevo() As Boolean
Dim Cad As String
    On Error GoTo EInsertarNuevo
    InsertarNuevo = False
    Cad = "INSERT INTO memoria (codigo, parametros, valor, descripcion, tipo) VALUES ("
    'Codigo
    Cad = Cad & Text1(1).Text & ","
    'Si es paramtro
    If Combo1.Enabled = False Then
        Cad = Cad & "1"
    Else
        Cad = Cad & "0"
    End If
    'VALOR
    Cad = Cad & ",'" & Text1(2).Text & "','"
    'DESCRIPCION
    Cad = Cad & Text1(0).Text & "',"
    'Tipo
    Cad = Cad & Combo1.ItemData(Combo1.ListIndex) & ")"
    
    'Ejecutamos
    Conn.Execute Cad
    
    InsertarNuevo = True
    Exit Function
EInsertarNuevo:
    MuestraError Err.Number, "Insertar. " & Err.Description
End Function



Private Sub ListView1_DblClick()
    If Not (ListView1.SelectedItem Is Nothing) Then Image1_Click (1)
End Sub


Private Sub ListView2_DblClick()
    If Not (ListView2.SelectedItem Is Nothing) Then Image2_Click (1)
End Sub




'----------------------------------------------------
Private Sub RealizarCalculoDatos()
Dim Insert As String
Dim Cad  As String
Dim Importe As Currency
Dim Fecha As Date
Dim Cerrado As Boolean

    'Borramos los campos de la tabla temporal
    Conn.Execute "delete from Usuarios.zmemoria where codusu =" & vUsu.Codigo
    
    'Cadena INSERT para todos
    Insert = "INSERT INTO Usuarios.zmemoria (codusu, codigo, parame, descripcion, valortexto, texto2, valornumero) VALUES (" & vUsu.Codigo & ","
    '--- valortexto, texto2, valornumero
    Set Rs = New ADODB.Recordset
    
    
    'Hacemos una primera insercion con los datos de parametros
    Cad = "Select * from memoria where parametros=1"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Cad = Rs!Codigo & "," & Rs!Parametros & ",'" & Rs!Descripcion & "','"
        Cad = Cad & Rs!Valor & "','T',0)"  'TEXTO
        Conn.Execute Insert & Cad
        
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    'FORMULAS de TEXTO
    Cad = "Select * from memoria where parametros=0 and tipo=5"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Cad = Rs!Codigo & "," & Rs!Parametros & ",'" & Rs!Descripcion & "','"
        Cad = Rs!Valor & "','T',0)"
        Conn.Execute Cad
        
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    
    'Algunos datos antes de empezar
    Fecha = DateAdd("yyyy", -1, vParam.fechaini)
    Cad = "Select distinct(numdiari) from hcabapu WHERE fechaent ='" & Format(Fecha, FormatoFecha) & "'"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cerrado = True
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            If Rs.Fields(0) > 0 Then Cerrado = False
        End If
    End If
    Rs.Close
    
    
    'Generamos los datos del asiento de apertura
    Cad = "Select * from memoria where parametros=0 and tipo=3"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Cad = Rs!Codigo & "," & Rs!Parametros & ",'" & Rs!Descripcion & "','"
        Importe = ProcesarFormulasCuentas(Rs!Valor, 3, Fecha, Cerrado)
        'Tenemos k obtener el valor
        Cad = Cad & Format(Importe, FormatoImporte) & "','N'," & TransformaComasPuntos(CStr(Importe))
        Cad = Cad & ")"
        Conn.Execute Insert & Cad
                    
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    
    'Cuentas del curso actual de la memoria
    Cad = "Select * from memoria where parametros=0 and tipo=1"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Cad = Rs!Codigo & "," & Rs!Parametros & ",'" & Rs!Descripcion & "','"
        If Rs!Valor <> "" Then
            Importe = ProcesarFormulasCuentas(Rs!Valor, 1, Fecha, Cerrado)
        Else
            Importe = 0
        End If
        'Tenemos k obtener el valor
        Cad = Cad & Format(Importe, FormatoImporte) & "','N'," & TransformaComasPuntos(CStr(Importe)) & ")"

        Conn.Execute Insert & Cad
                    
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    
    'Algunos datos para el año anterior a la memoria
    Fecha = DateAdd("yyyy", -2, vParam.fechaini)
    Cad = "Select distinct(numdiari) from hcabapu WHERE fechaent>='" & Format(Fecha, FormatoFecha) & "'"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cerrado = True
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            If Rs.Fields(0) > 0 Then Cerrado = False
        End If
    End If
    Rs.Close
    
    
    'Curso Anterior
    Cad = "Select * from memoria where parametros=0 and tipo=2"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Cad = Rs!Codigo & "," & Rs!Parametros & ",'" & Rs!Descripcion & "','"
        Importe = ProcesarFormulasCuentas(Rs!Valor, 2, Fecha, Cerrado)
        'Tenemos k obtener el valor
        Cad = Cad & Format(Importe, FormatoImporte) & "','N'," & TransformaComasPuntos(CStr(Importe)) & ")"
        Conn.Execute Insert & Cad
                    
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    
    
    'Ahora haremos la formula
    Cad = "Select * from memoria where parametros=0 and tipo=4"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Cad = Rs!Codigo & "," & Rs!Parametros & ",'" & Rs!Descripcion & "','"
        Importe = ProcesarFormulasCuentas(Rs!Valor, 4, Fecha, Cerrado)
        Cad = Cad & Format(Importe, FormatoImporte) & "','N'," & TransformaComasPuntos(CStr(Importe)) & ")"
        Conn.Execute Insert & Cad
                    
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    
End Sub



'-------------------------------------------------
'
'   Opciones    0-
'               1-
'               2
'               3
'               4
'
Private Function ProcesarFormulasCuentas(CADENA As String, opcion As Byte, FechaInicio As Date, EnCerrado As Boolean) As Currency
Dim i As Integer
Dim J As Integer
'Dim K As Integer
Dim Cad As String
Dim Importe As Currency
Dim Impo As Currency
Dim B As Boolean

    On Error GoTo EProcesarFormulasCuentas
    ProcesarFormulasCuentas = 0
    Importe = 0
    
'''''    'Comprobamos los parentesis
'''''    J = 0
'''''    For I = 1 To Len(Cadena)
'''''        If Mid(Cadena, I, 1) = ")" Then J = J + 1
'''''    Next I
'''''    Cad = CStr(J)
'''''    J = 0
'''''    For I = 1 To Len(Cadena)
'''''        If Mid(Cadena, I, 1) = ")" Then J = J + 1
'''''    Next I
'''''    If Val(Cad) <> J Then
'''''        'MsgBox "No coincide el numero de parentesis", vbExclamation
'''''        Exit Function
'''''    End If
'''''
'''''
'''''    'Vemos is tiene parentesis
'''''    B = J > 0 '->tiene
'''''    While B
'''''        'Tiene parentesis. Habra que mandar a calcualar con la subcadna k gnerea
'''''        I = InStr(1, Cadena, ")")
'''''        J = 0
'''''        For K = I To 1 Step -1
'''''            If Mid(Cadena, K, 1) = "(" Then
'''''                J = K
'''''                Exit For
'''''            End If
'''''        Next K
'''''
'''''        'Si j>0 entonces ha encontrado la subcadena con parentesis
'''''        Cad = Mid(Cadena, J + 1, I - J - 1)
'''''        Impo = ProcesarFormulasCuentas(Cad)
'''''        Cad = Mid(Cadena, 1, J - 1) & " @" & CStr(Impo) & Mid(Cadena, I + 1)
'''''        Cadena = Cad
'''''
'''''        'Ahora vemos si kedan parentesis
'''''        B = (InStr(1, Cadena, "(") > 0)
'''''
'''''    Wend
    
    'Ya hemos kitado los parentesis. Ahora recorremos la cadena para obtener el valor
    'primero vemos si el primer caracter es un $ o una @ de quitar parentesis
    CADENA = "+" & CADENA
    
    While CADENA <> ""
                                                    
        B = False
        For i = 2 To Len(CADENA)
            Cad = Mid(CADENA, i, 1)
            'If Cad = "+" Or Cad = "-" Or Cad = "\" Or Cad = "/" Or Cad = "*" Then
            If Cad = "+" Or Cad = "-" Then
                B = True
                Exit For
            End If
        Next i
        If B Then
            Cad = Mid(CADENA, 1, i - 1)
            CADENA = Mid(CADENA, i)
        Else
            'Solo keda un bloque
            Cad = CADENA
            CADENA = ""
        End If
        
        'Procesar bloque
        'Signo
        If Mid(Cad, 1, 1) = "-" Then
            i = -1
        Else
            i = 1
        End If
        Cad = Mid(Cad, 2)
        Select Case opcion
        Case 1, 2, 3
            '1 Actual
            '2 Anterior
            '3 ASiento apertura
            If Cad <> "" Then _
                Impo = CalcularImporteCta(FechaInicio, Cad, EnCerrado, opcion)
                
        Case 4
            'Formula
            If Cad <> "" Then _
                Impo = DevuelveValorParaFormula(Cad)
        Case Else
            Impo = 0
        End Select
        Impo = i * Impo
        Importe = Importe + Impo
    Wend
    ProcesarFormulasCuentas = Importe
    Exit Function
EProcesarFormulasCuentas:
    MuestraError Err.Number, "Procesar Formulas Cuentas"
End Function


Private Function DevuelveValorParaFormula(ByRef CADENA As String) As Currency
Dim RT As ADODB.Recordset
Dim Cad As String

    DevuelveValorParaFormula = 0
    Set RT = New ADODB.Recordset
    Cad = "Select ValorNumero from Usuarios.zmemoria where codusu= " & vUsu.Codigo & " AND parame = 0"
    Cad = Cad & " AND codigo = " & Mid(CADENA, 2)  'le kitamos el $
    RT.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then
            DevuelveValorParaFormula = RT.Fields(0)
        End If
    End If
    RT.Close
    Set RT = Nothing
End Function



Private Sub CargaDatos()
Dim Cad As String
Dim i As Integer
    Set Rs = New ADODB.Recordset
    Cad = "Select * from Usuarios.zmemoria WHERE codusu = " & vUsu.Codigo
    Cad = Cad & " ORDER BY codigo"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Set ItmX = ListView3.ListItems.Add(, , Rs!Codigo)
        ItmX.SubItems(1) = Rs.Fields(3)
        ItmX.SubItems(2) = Rs.Fields(4)
        'En el tag ponemos el tipo
        i = Rs!parame + 1
        ItmX.SmallIcon = i
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub

