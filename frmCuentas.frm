VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCuentas 
   Caption         =   "Mantenimiento de cuentas"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   Icon            =   "frmCuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   630
      Left            =   6135
      TabIndex        =   28
      Top             =   540
      Width           =   1890
      Begin VB.CheckBox chkUltimo 
         Caption         =   "Cta último nivel"
         Height          =   300
         Left            =   105
         TabIndex        =   29
         Top             =   225
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Cancel          =   -1  'True
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7005
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5685
      TabIndex        =   11
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7005
      TabIndex        =   12
      Top             =   6000
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   120
      TabIndex        =   26
      Top             =   5880
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   255
         Width           =   2955
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos cuentas con apuntes directos"
      Height          =   4590
      Left            =   135
      TabIndex        =   16
      Top             =   1200
      Width           =   7905
      Begin VB.CheckBox Check1 
         Caption         =   "Modelo 347"
         Height          =   225
         Left            =   5385
         TabIndex        =   30
         Tag             =   "E-Mail|T|S|||cuentas|model347|||"
         Top             =   660
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   225
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "Razón social|T|S|||cuentas|razosoci|||"
         Text            =   "123456789012345678901234567890"
         Top             =   615
         Width           =   3105
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   240
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Domicilio|T|S|||cuentas|dirdatos|||"
         Text            =   "Text1"
         Top             =   1305
         Width           =   3120
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   240
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "Cod. Postal|T|S|||cuentas|codposta|||"
         Text            =   "Text1"
         Top             =   2085
         Width           =   840
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   3570
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Población|T|S|||cuentas|despobla|||"
         Text            =   "Text1"
         Top             =   1305
         Width           =   3120
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   1260
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Provincia|T|S|||cuentas|desprovi|||"
         Text            =   "Text1"
         Top             =   2070
         Width           =   2130
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   3555
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "NIF|T|S|||cuentas|nifdatos|||"
         Text            =   "Text1"
         Top             =   615
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   3555
         MaxLength       =   40
         TabIndex        =   8
         Tag             =   "E-Mail|T|S|||cuentas|maidatos|||"
         Text            =   "Text1"
         Top             =   2055
         Width           =   3765
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   9
         Left            =   240
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "Direccion web|T|S|||cuentas|webdatos|||"
         Text            =   "Text1"
         Top             =   2790
         Width           =   4800
      End
      Begin VB.TextBox Text1 
         Height          =   795
         Index           =   10
         Left            =   240
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   10
         Tag             =   "Observaciones|T|S|||cuentas|obsdatos|||"
         Text            =   "frmCuentas.frx":000C
         Top             =   3465
         Width           =   7065
      End
      Begin VB.Label Label1 
         Caption         =   "Razón social"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   25
         Top             =   330
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "N.I.F."
         Height          =   255
         Index           =   7
         Left            =   3555
         TabIndex        =   24
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   23
         Top             =   1065
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "C.Postal"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Top             =   1830
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   5
         Left            =   3570
         TabIndex        =   21
         Top             =   1065
         Width           =   3465
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         Height          =   255
         Index           =   6
         Left            =   1260
         TabIndex        =   20
         Top             =   1830
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "MAIL"
         Height          =   255
         Index           =   8
         Left            =   3555
         TabIndex        =   19
         Top             =   1830
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Dirección web"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   18
         Top             =   2550
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   17
         Top             =   3210
         Width           =   1905
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   210
      MaxLength       =   15
      TabIndex        =   0
      Tag             =   "Codigo cuenta|T|N|||cuentas|codmacta||S|"
      Text            =   "Text1"
      Top             =   750
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1815
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Denominación cuenta|T|N|||cuentas|nommacta|||"
      Text            =   "Text1"
      Top             =   765
      Width           =   3900
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   630
      Left            =   150
      Top             =   5805
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1111
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   4560
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":0012
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":0124
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":0236
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":0348
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":045A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":056C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":0E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":1720
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":1FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":28D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":31AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":3600
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":3712
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":3824
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":3936
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":3FB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Denominación"
      Height          =   255
      Index           =   1
      Left            =   1830
      TabIndex        =   15
      Top             =   525
      Width           =   3465
   End
   Begin VB.Label Label1 
      Caption         =   "Cod."
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   14
      Top             =   525
      Width           =   1215
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private NumRegistro As Long
Private kCampo As Integer
Private TotalReg As Long
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean


Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                'MsgBox "Registro insertado.", vbInformation
                PonerModo 0
                lblIndicador.Caption = "Insertado"
            End If
        End If
    Case 4
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    If SituarData1 Then
                        PonerModo 2
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
    Case 1
        HacerBusqueda
    End Select
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1, 3
    LimpiarCampos
    PonerModo 0
Case 4
    'Modificar
    lblIndicador.Caption = NumRegistro & " de " & TotalReg
    PonerModo 2
End Select

End Sub


' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)
Private Function SituarData1() As Boolean
    Dim SQL As String
    On Error GoTo ESituarData1
            'Actualizamos el recordset
            Data1.Refresh
            '#### A mano.
            'El sql para que se situe en el registro en especial es el siguiente
            SQL = " codempre = '" & Text1(0).Text & "'"
            Data1.Recordset.Find SQL
            If Data1.Recordset.EOF Then GoTo ESituarData1
            SituarData1 = True
        Exit Function
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData1 = False
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Aceptar"
    PonerModo 3
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    'Disablamos el frame1
    Frame1.Visible = False
    '###A mano
    Text1(0).SetFocus
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        lblIndicador.Caption = "Búsqueda"
        PonerModo 1
        '### A mano
        '################################################
        'Si pasamos el control aqui lo ponemos en amarillo
        Text1(0).SetFocus
        Text1(0).BackColor = vbYellow
        Else
            HacerBusqueda
            If TotalReg = 0 Then
                 '### A mano
                Text1(kCampo).Text = ""
                Text1(kCampo).BackColor = vbYellow
                Text1(kCampo).SetFocus
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla
        PonerCadenaBusqueda
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
Select Case Index
    Case 0
        Data1.Recordset.MoveFirst
        NumRegistro = 1
    Case 1
        Data1.Recordset.MovePrevious
        NumRegistro = NumRegistro - 1
        If Data1.Recordset.BOF Then
            Data1.Recordset.MoveFirst
            NumRegistro = 1
        End If
    Case 2
        Data1.Recordset.MoveNext
        NumRegistro = NumRegistro + 1
        If Data1.Recordset.EOF Then
            Data1.Recordset.MoveLast
            NumRegistro = TotalReg
        End If
    Case 3
        Data1.Recordset.MoveLast
        NumRegistro = TotalReg
End Select
PonerCampos
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Modificar"
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
    Text1(0).Locked = True
    Text1(0).BackColor = &H80000018
    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
End Sub

Private Sub BotonEliminar()
    Dim Cad As String
    Dim i As Integer
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    '### a mano
    Cad = "Seguro que desea eliminar de la BD el registro:"
    Cad = Cad & vbCrLf & "Cod. empresa: " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Nombre empresa: " & Data1.Recordset.Fields(1)
    i = MsgBox(Cad, vbQuestion + vbYesNo)
    'Borramos
    If i = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        Data1.Recordset.Delete
        Data1.Refresh
        If Data1.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCampos
            PonerModo 0
            Else
                If NumRegistro = TotalReg Then
                        'He borrado el que era el ultimo
                        Data1.Recordset.MoveLast
                        NumRegistro = NumRegistro - 1
                        Else
                            For i = 1 To NumRegistro - 1
                                Data1.Recordset.MoveNext
                            Next i
                End If
                TotalReg = TotalReg - 1
                PonerCampos
        End If
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number > 0 Then MsgBox Err.Number & " - " & Err.Description
End Sub




Private Sub cmdRegresar_Click()
Dim Cad As String
Dim i As Integer
Dim j As Integer
Dim aux As String

If Data1.Recordset.EOF Then
    MsgBox "Ningún registro devuelto.", vbExclamation
    Exit Sub
End If

Cad = ""
i = 0
Do
    j = i + 1
    i = InStr(j, DatosADevolverBusqueda, "|")
    If i > 0 Then
        aux = Mid(DatosADevolverBusqueda, j, i - j)
        j = Val(aux)
        Cad = Cad & Text1(j).Text & "|"
    End If
Loop Until i = 0
RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim ancho As Single
    LimpiarCampos

    '## A mano
    NombreTabla = "cuentas"
    Ordenacion = " ORDER BY codmacta"
        
    'Para que no puedan introduci valores mayores que el numero de cuenta permitido
    Text1(0).MaxLength = vEmpresa.DigitosUltimoNivel
    ancho = 120 * vEmpresa.DigitosUltimoNivel
    Text1(0).Width = ancho
    ancho = Text1(0).Left + ancho + 150
    Label1(1).Left = ancho
    Text1(1).Left = ancho
    'Para todos
    Me.Data1.Password = vUsu.Passwd
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    Data1.RecordSource = "Select * from " & NombreTabla
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
        Else
        PonerModo 1
        '### A mano
        Text1(0).BackColor = vbYellow
    End If

End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    chkUltimo.Value = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim aux As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        'If CadB <> "" Then CadB = CadB & " AND "
        'CadB = CadB & Aux
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
BotonAnyadir
End Sub

Private Sub mnSalir_Click()
Screen.MousePointer = vbHourglass
Unload Me
End Sub

Private Sub mnVerTodos_Click()
BotonVerTodos
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Modo = 1 Then
        Text1(Index).BackColor = vbYellow
        Else
            Text1(Index).SelStart = 0
            Text1(Index).SelLength = Len(Text1(Index).Text)
    End If
End Sub


'Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        KeyCode = 0
'        SendKeys "{TAB}"
'    End If
'End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
    Dim i As Integer
    Dim SQL As String
    Dim mTag As CTag
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = &H80000018
    End If
    FormateaCampo Text1(Index)  'Formateamos el campo si tiene valor
    'Si queremos hacer algo ..
    Select Case Index
        Case 0
            PierdeFocoCodigoCuenta
        Case 1
        
        '....
    End Select
    '---
End Sub

Private Sub HacerBusqueda()
Dim Cad As String
Dim CadB As String
CadB = ObtenerBusqueda(Me)

If chkVistaPrevia = 1 Then
    MandaBusquedaPrevia CadB
    Else
        'Se muestran en el mismo form
        If CadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 30, "Cuenta")
        Cad = Cad & ParaGrid(Text1(1), 70, "Denominación")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSql = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1"
            frmB.vTitulo = "CUENTAS"
            frmB.vSelElem = 1
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                Text1(kCampo).SetFocus
            End If
        End If
End Sub



Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

Data1.RecordSource = CadenaConsulta
Data1.Refresh
If Data1.Recordset.RecordCount <= 0 Then
    MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
    Screen.MousePointer = vbDefault
    TotalReg = 0
    Exit Sub

    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        TotalReg = Data1.Recordset.RecordCount
        NumRegistro = 1
        PonerCampos
End If


Screen.MousePointer = vbDefault
Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
    Dim i As Integer
    Dim mTag As CTag
    Dim SQL As String
    If Data1.Recordset.EOF Then Exit Sub
    'Este form no llama a la funcion generecia por que es un form particular
    'PonerCamposForma Me, Data1
    
    Text1(0).Text = Data1.Recordset!codmacta
    Text1(1).Text = Data1.Recordset!nommacta
    Check1.Value = 0
    If Data1.Recordset!apudirec = "S" Then
        chkUltimo.Value = 1
        PonerCamposForma Me, Data1
        
        
        
        
    Else
        chkUltimo.Value = 0
    End If
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = NumRegistro & " de " & TotalReg
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim i As Integer
    Dim b As Boolean
    If Modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
        For i = 0 To Text1.Count - 1
            'Text1(i).BackColor = vbWhite
            Text1(0).BackColor = &H80000018
        Next i
        'chkVistaPrevia.Visible = False
    End If
    Modo = Kmodo
    'chkVistaPrevia.Visible = (Modo = 1)
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    DespalzamientoVisible b
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(8).Enabled = b
    mnEliminar.Enabled = b
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = b
    Else
        cmdRegresar.Visible = False
    End If
    
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.Visible = b Or Modo = 1
    cmdCancelar.Visible = b Or Modo = 1
    mnOpciones.Enabled = Not b
    Toolbar1.Buttons(6).Enabled = Not b
    Toolbar1.Buttons(1).Enabled = Not b
    Toolbar1.Buttons(2).Enabled = Not b
    
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    b = (Modo = 2) Or Modo = 0
    For i = 0 To Text1.Count - 1
        Text1(i).Locked = b
        If b Then
            Text1(i).BackColor = &H80000018
        ElseIf Modo <> 1 Then
            Text1(i).BackColor = vbWhite
        End If
    Next i
    
End Sub


Private Function DatosOk() As Boolean
Dim Rs As ADODB.Recordset
Dim b As Boolean
b = CompForm(Me)
DatosOk = b
End Function


'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Sub SugerirCodigoSiguiente()

' Como es texto y no tiene por que ser numérico entonces
'no sugerimos nada

'''    Dim SQL As String
'''    Dim Rs As ADODB.Recordset
'''
'''    SQL = "Select Max(codprodu) from " & NombreTabla
'''    Text1(0).Text = 1
'''    Set Rs = New ADODB.Recordset
'''    Rs.Open SQL, Conn, , , adCmdText
'''    If Not Rs.EOF Then
'''        If Not IsNull(Rs.Fields(0)) Then
'''            Text1(0).Text = Rs.Fields(0) + 1
'''        End If
'''    End If
'''    Rs.Close
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    BotonBuscar
Case 2
    BotonVerTodos
Case 6
    BotonAnyadir
Case 7
    BotonModificar
Case 8
    BotonEliminar
Case 11
        'Imprimimos el listado

            frmImprimir.Opcion = 1
            frmImprimir.Show vbModal

Case 12
        frmImprimir.Opcion = 1
        Unload Me
Case 14 To 17
    Desplazamiento (Button.Index - 14)
'Case 20
'    'Listado en crystal report
'    Screen.MousePointer = vbHourglass
'    CR1.Connect = Conn
'    CR1.ReportFileName = App.Path & "\Informes\list_Inc.rpt"
'    CR1.WindowTitle = "Listado incidencias."
'    CR1.WindowState = crptMaximized
'    CR1.Action = 1
'    Screen.MousePointer = vbDefault


Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    Dim i
    For i = 14 To 17
        Toolbar1.Buttons(i).Visible = bol
    Next i
End Sub


Private Sub PierdeFocoCodigoCuenta()
If Text1(0).Text = "" Then Exit Sub
'Si no compruebo que es un campo numerico
If Not IsNumeric(Text1(0).Text) Then
    MsgBox "El código de cuenta es un campo numérico", vbExclamation
    Exit Sub
End If

'Vemos si a puesto el punto para rellenar
Text1(0).Text = RellenaCodigoCuenta(Text1(0).Text)

If Len(Text1(0).Text) > vEmpresa.DigitosUltimoNivel Then
    MsgBox "El número maximo de digitos para las cuentas es de " & vEmpresa.DigitosUltimoNivel & _
        vbCrLf & "La cuenta que ha puesto tiene " & Len(Text1(0).Text), vbExclamation
    Exit Sub
End If

'Ponemos , si es de ultimo nivel habilitados los campos
Frame1.Visible = (Len(Text1(0).Text) = vEmpresa.DigitosUltimoNivel)
End Sub
