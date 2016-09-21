VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBuscaGrid 
   Caption         =   "Búsqueda"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   Icon            =   "frmBuscaGrid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2880
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   435
      Left            =   4440
      TabIndex        =   2
      Top             =   5100
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   435
      Left            =   6000
      TabIndex        =   3
      Top             =   5100
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7275
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmBuscaGrid.frx":1CFA
      Height          =   3675
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6482
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Leyendo datos servidor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   2520
   End
   Begin VB.Label Label3 
      Caption         =   "Búsqueda"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "TITULO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cargando datos ...."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1560
      TabIndex        =   5
      Top             =   2640
      Width           =   3675
   End
End
Attribute VB_Name = "frmBuscaGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Selecionado(CadenaDevuelta As String)

'Variables publicas para montar datos
Public vTabla As String
Public vCampos As String 'columnas en la tabla.Empipados
Public vSelElem As Integer
Public vTitulo As String
Public vSQL As String
'Dentro de campos vendra cada grupo separado por ·
'Y cada grupo sera Desc|Tabla|Tipo|Porcentaje de ancho
Public vDevuelve As String 'Empipados los campos que devuelve



'Variables privadas
Dim PrimeraVez As Boolean
Dim SQL As String
'Las redimensionaremos
Dim TotalArray As Integer
Dim Cabeceras() As String
Dim CabTablas() As String
Dim CabAncho() As Single
Dim TipoCampo() As String
Private Busca As Boolean
Private DbClick As Boolean



Private Sub adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Busca Then
        Busca = False
        Exit Sub
    End If
    DbClick = True
    If Adodc1.Recordset.BOF Then
        If Adodc1.Recordset.RecordCount > 0 Then Adodc1.Recordset.MoveFirst
    End If
    If Adodc1.Recordset.RecordCount > 0 Then
        Text1.Text = DBLet(Adodc1.Recordset.Fields(CabTablas(vSelElem)))
    End If
End Sub

Private Sub cmdRegresar_Click()
Dim vDes As String
Dim i, J As Integer
Dim V As String

If Adodc1.Recordset Is Nothing Then Exit Sub
If Adodc1.Recordset.EOF Then Exit Sub
i = 0
vDes = ""
Do
    J = i + 1
    i = InStr(J, vDevuelve, "|")
    If i > 0 Then
        V = Mid(vDevuelve, J, i - J)
        If V <> "" Then
            If IsNumeric(V) Then
                If Val(V) <= TotalArray Then vDes = vDes & Adodc1.Recordset(CabTablas(Val(V))) & "|"
            End If
        End If
    End If
Loop Until i = 0
RaiseEvent Selecionado(vDes)
Unload Me
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub DataGrid1_DblClick()
If Adodc1.Recordset Is Nothing Then Exit Sub
If Adodc1.Recordset.EOF Then Exit Sub
cmdRegresar_Click
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim Cad As String

If Adodc1.Recordset Is Nothing Then Exit Sub
If Adodc1.Recordset.EOF Then Exit Sub
If vSelElem = ColIndex Then Exit Sub
Cad = "¿Desea reordenar por el concepto " & DataGrid1.Columns(ColIndex).Caption & "?"
If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
If ColIndex <= TotalArray Then
    Me.Refresh
    Screen.MousePointer = vbHourglass
    vSelElem = ColIndex
    CargaGrid
    Screen.MousePointer = vbDefault
    Else
    MsgBox "Error cargando tabla. Imposible ordenacion", vbCritical
End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub Form_Activate()
Dim Ok As Boolean
If PrimeraVez Then
    PrimeraVez = False
    Me.Refresh
    Screen.MousePointer = vbHourglass
    Ok = ObtenerTamanyosArray
    If Ok Then Ok = SeparaCampos
    If Not Ok Then
        'Error en SQL
        'Salimos
        Unload Me
        Exit Sub
    End If
    CargaGrid
    Label4.Visible = False
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Label4.Visible = True
DataGrid1.Enabled = False
PrimeraVez = True
Label1.Caption = vTitulo
DbClick = True
'Adodc1.password = vUsu.Passwd
End Sub



Private Function SeparaCampos() As Boolean
Dim Cad As String
Dim Grupo As String
Dim i As Integer
Dim J As Integer
Dim C As Integer 'Contrador dentro del array

SeparaCampos = False
i = 0
C = 0
Do
    J = i + 1
    i = InStr(J, vCampos, "·")
    If i > 0 Then
        Grupo = Mid(vCampos, J, i - J)
        'Y en la martriz
        InsertaGrupo Grupo, C
        C = C + 1
    End If
Loop Until i = 0
SeparaCampos = True
End Function

Private Sub InsertaGrupo(Grupo As String, Contador As Integer)
Dim i As Integer
Dim J As Integer
Dim Cad As String
J = 0


    Cad = ""
    
    'Cabeceras
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        Cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        J = 1
    End If
    Cabeceras(Contador) = Cad
    
    'TAblas BD
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        Cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        Else
            Cad = ""
            Grupo = ""
    End If
    
    CabTablas(Contador) = Cad
    
    'El tipo
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        Cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        Else
            Cad = ""
            Grupo = ""
    End If
    
    TipoCampo(Contador) = Cad
    
    'Por ultimo
    'ANCHO
    If Grupo = "" Then Grupo = 0
    CabAncho(Contador) = Grupo
End Sub

Private Function ObtenerTamanyosArray() As Boolean
Dim i As Integer
Dim J As Integer
Dim Grupo As String

ObtenerTamanyosArray = False
'Primero a los campos de la tabla
TotalArray = -1
J = 0
Do
    i = J + 1
    J = InStr(i, vCampos, "·")
    If J > 0 Then TotalArray = TotalArray + 1
Loop Until J = 0
If TotalArray < 0 Then Exit Function
'Las redimensionaremos
ReDim Cabeceras(TotalArray)
ReDim CabTablas(TotalArray)
ReDim CabAncho(TotalArray)
ReDim TipoCampo(TotalArray)
ObtenerTamanyosArray = True
End Function


Private Sub CargaGrid()
Dim Cad As String
Dim i As Integer
Dim anc As Single


'On Error GoTo ECargaGRid '##QUITAR
'Generamos SQL
Cad = ""
For i = 0 To TotalArray
    If Cad <> "" Then Cad = Cad & ","
    Cad = Cad & CabTablas(i)
Next i
Cad = "SELECT " & Cad & " FROM " & vTabla
If vSQL <> "" Then Cad = Cad & " WHERE " & vSQL
Cad = Cad & " ORDER BY " & CabTablas(vSelElem)

DataGrid1.AllowRowSizing = False
Adodc1.ConnectionString = Conn
Adodc1.RecordSource = Cad
Adodc1.Refresh

DataGrid1.Visible = True
'Cargamos el grid
anc = DataGrid1.Width - 640

For i = 0 To TotalArray
    
    DataGrid1.Columns(i).Caption = Cabeceras(i)
    If CabAncho(i) = 0 Then
        DataGrid1.Columns(i).Visible = False
        Else
        DataGrid1.Columns(i).Width = anc * (CabAncho(i) / 100)
    End If
Next i


'Habilitamos el text1 para que escriban
DataGrid1.Enabled = True
Text1.Enabled = True
If Not Adodc1.Recordset.EOF Then
    'Le ponemos el 1er registro
    cmdRegresar.Enabled = True
    Cad = CabTablas(vSelElem)
    Text1.Text = DBLet(Adodc1.Recordset(Cad))
    Text1.SetFocus
    Else
        cmdRegresar.Enabled = False
        cmdSalir.SetFocus
End If
Exit Sub
ECargaGRid:
    MuestraError Err.Number, "Carga grid." & vbCrLf & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DataGrid1.Enabled = False
End Sub



Private Sub Text1_Change()
Dim SQLDBGRID As String

    If DbClick Then
        DbClick = False
        Exit Sub
    End If

    Busca = True
    SQLDBGRID = CabTablas(vSelElem)
    Select Case TipoCampo(vSelElem)
        Case "N"
            If Not IsNumeric(Text1.Text) Then
                If Adodc1.Recordset.RecordCount > 0 Then Adodc1.Recordset.MoveFirst
                Exit Sub
            End If
            SQLDBGRID = SQLDBGRID & " >= " & Trim(Text1)
        Case "T"
            SQLDBGRID = SQLDBGRID & " >= '" & Trim(Text1) & "'"
        Case Else
            Exit Sub
    End Select
    Screen.MousePointer = vbHourglass
    Adodc1.Recordset.Find SQLDBGRID, , adSearchForward, 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdRegresar_Click
End If
End Sub


' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_LostFocus()
  WheelUnHook
End Sub

