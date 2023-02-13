VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmImportacionNavFra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Totales factura consum"
   ClientHeight    =   10530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   ClipControls    =   0   'False
   Icon            =   "frmNavaImportacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   23
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   19
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtDatos 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   16
      Text            =   "Text5"
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3300
      TabIndex        =   12
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   320
      Index           =   3
      Left            =   8760
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   10
      Text            =   "existencia"
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   320
      Index           =   2
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   9
      Text            =   "existencia"
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   320
      Index           =   1
      Left            =   7680
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   8
      Text            =   "existencia"
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   320
      Index           =   0
      Left            =   6600
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   0
      Text            =   "existencia"
      Top             =   5760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   9960
      Width           =   3255
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   10080
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8760
      TabIndex        =   2
      Top             =   10080
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1440
      Top             =   10080
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   582
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmNavaImportacion.frx":000C
      Height          =   7695
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   13573
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   4440
      Top             =   9840
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   582
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
      Caption         =   "data2"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmNavaImportacion.frx":0020
      Height          =   1815
      Left            =   5160
      TabIndex        =   11
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Observa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   4
      Left            =   1800
      TabIndex        =   24
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Numero factura"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   1680
      TabIndex        =   20
      Top             =   720
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fec. factura"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Centro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fec. Recepcion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   11
      Left            =   3300
      TabIndex        =   14
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label lblInfInv 
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   9960
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   10500
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmImportacionNavFra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TotalLineas As Currency


Private WithEvents frmPre As frmAsiPre
Attribute frmPre.VB_VarHelpID = -1
Private Modo As Byte
Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
Dim PrimeraVez As Boolean
Dim cad As String


Private Sub cmdAceptar_Click()
Dim Bases As Currency
Dim Ivas As Currency

  
    'Fechas dentro de ejerciccios y period NO liquidado
    If Not DatosCorrecto Then Exit Sub
  
  
  
    Set miRsAux = New ADODB.Recordset
   
    cad = ""
    Bases = 0: Ivas = 0
    data2.Recordset.MoveFirst
    While Not data2.Recordset.EOF
        cad = cad & vbCrLf & "IVA: " & data2.Recordset.Fields(0) & "% " & vbCrLf
        cad = cad & "Base: " & data2.Recordset.Fields(1) & vbCrLf
        cad = cad & "Imp iva: " & data2.Recordset.Fields(2) & vbCrLf
        Bases = Bases + data2.Recordset.Fields(1)
        Ivas = Ivas + data2.Recordset.Fields(2)
        
        data2.Recordset.MoveNext
    Wend
    data2.Recordset.MoveFirst
    
    
    cad = cad & vbCrLf & vbCrLf
    cad = cad & "TOTAL" & vbCrLf & "               " & Bases & " +  " & Ivas & " = " & Bases + Ivas & vbCrLf
    
    cad = String(40, "*") & cad & String(40, "*") & vbCrLf & "¿Generar factura?"
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
    
    
        
        
    CadenaDesdeOtroForm = Text3.Text & "|" & txtDatos(5).Text & "|" & txtDatos(0).Text & "|"
    
    Me.Tag = 0
    Unload Me
    
    
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
   
    
    If MsgBox("Desea cancelar el proceso?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    CadenaDesdeOtroForm = ""
    Me.Tag = 0
    Unload Me
  
   
End Sub





Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF And gridCargado And Modo = 4 Then
       CargaTxtAux True, False
       'txtAux.SelStart = Len(Me.txtAux.Text)
       
       txtAux(0).SetFocus
       
       txtAux(0).SelStart = 0
       txtAux(0).SelLength = Len(Me.txtAux(0).Text)
       txtAux(0).Refresh
       PonIndicador
    End If
End Sub

Private Sub PonIndicador()
On Error Resume Next
       Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
       If Err.Number <> 0 Then
            Me.lblIndicador.Caption = ""
            Err.Clear
        End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
         PrimeraVez = False
         CargaDatosFra
         CargaTxtAux True, False
         PonFoco txtAux(0)
    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
    
 
    
    Me.Tag = 1 'NO se puede cerrar mas que de boton
    Modo = 4
    
    CadenaDesdeOtroForm = ""
    
    If vParam.autocoste Then
        Me.Width = 10315
        
    End If
    DataGrid1.Width = Me.Width - 400
    Me.cmdCancelar.Left = Me.Width - 1365
    Me.cmdAceptar.Left = Me.Width - 2565
    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
    
    DataGrid1.RowHeight = 320
    DataGrid2.RowHeight = 320
    CargaGrid

    
    PrimeraVez = True
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid()
Dim K As Byte
Dim Tot As Currency

On Error GoTo ECarga

    gridCargado = False
    
    cad = "select seccion,descripcion,base,porceniva,iva,base+iva from "
    cad = cad & " importanatmptotal,importnavconceptos where seccion=concepto ORDER BY seccion, iva"

    Data1.ConnectionString = Conn
    Data1.RecordSource = cad
    Data1.CursorType = adOpenDynamic
    Data1.LockType = adLockPessimistic
    Data1.Refresh
   
    
    PrimeraVez = False
        

    
    DataGrid1.Columns(0).Caption = "Seccion"
    DataGrid1.Columns(0).Width = 800
    
    
        
    DataGrid1.Columns(1).Caption = "Descripcion"
    DataGrid1.Columns(1).Width = 3700
    
    
    
    
    For K = 2 To 5
        DataGrid1.Columns(K).Caption = RecuperaValor("Base|%IVA|Importe|Total|", K - 1)
        DataGrid1.Columns(K).Width = Val(RecuperaValor("1050|800|1000|1400|", K - 1))
        DataGrid1.Columns(K).NumberFormat = FormatoImporte
        DataGrid1.Columns(K).Alignment = dbgRight
    Next
    
    
    DataGrid1.ScrollBars = dbgAutomatic
    
    
    
    cad = "select  porceniva '%iva',sum(base) base,sum(iva) IVA,sum(base+iva) Total"
    cad = cad & " from importanatmptotal  group by 1"


    data2.ConnectionString = Conn
    data2.RecordSource = cad
    data2.CursorType = adOpenDynamic
    data2.LockType = adLockPessimistic
    data2.Refresh
    For K = 0 To 3
        'DataGrid1.Columns(K).Caption = RecuperaValor("Base|%IVA|Importe|Total|", K - 1)
        DataGrid2.Columns(K).Width = Val(RecuperaValor("650|1000|900|1050|", K + 1))
        DataGrid2.Columns(K).NumberFormat = FormatoImporte
        DataGrid2.Columns(K).Alignment = dbgRight
    Next
    
    txtDatos(4).Text = ""
    Tot = 0
    While Not data2.Recordset.EOF
        Tot = Tot + data2.Recordset!Total
        data2.Recordset.MoveNext
    Wend
    txtDatos(4).Text = Format(Tot, FormatoImporte)
    
    
    gridCargado = True
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, Limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim J As Byte


    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux(0).top = 290
        txtAux(0).visible = visible
        txtAux(1).visible = visible
    Else
        DeseleccionaGrid Me.DataGrid1
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.top + 220
        Else
            alto = DataGrid1.top + DataGrid1.RowTop(DataGrid1.Row) + 20
        End If
        
        For J = 0 To 3
            If Limpiar Then
                txtAux(J).Text = ""
            Else
                txtAux(J).Text = Format(Data1.Recordset.Fields(J + 2), FormatoImporte)
            End If
            txtAux(J).top = alto
            txtAux(J).Locked = J > 0
            txtAux(J).Height = DataGrid1.RowHeight
            txtAux(J).Left = DataGrid1.Columns(J + 2).Left + 130
            txtAux(J).Width = DataGrid1.Columns(J + 2).Width - 10
            txtAux(J).visible = visible
        Next
        

        
        
        
        
        
        
        
    End If
'    PonFoco txtAux(0)
    
'    If visible Then
'        txtAux(0).TabIndex = 2
'      '  txtAux.SelStart = 0
'       ' txtAux.SelLength = Len(txtAux.Text)
'    Else
'        txtAux(0).TabIndex = 5
'    End If
End Sub






Private Sub Form_Unload(Cancel As Integer)
    If Me.Tag = 1 Then Cancel = 1 'o aceptar o cancelar
End Sub

Private Sub frmPre_DatoSeleccionado(CadenaSeleccion As String)
    cad = CadenaSeleccion
End Sub


Private Sub txtaux_GotFocus(Index As Integer)
    txtAux(Index).SelStart = 0
    txtAux(Index).SelLength = Len(txtAux(Index).Text)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    If KeyCode = 38 Or KeyCode = 40 Then
        ModificarExistencia
    End If

    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Bookmark > 0 Then
                    DataGrid1.Bookmark = DataGrid1.Bookmark - 1
                    PonFoco txtAux(Index)
                End If
        
        Case 40 'Desplazamiento Flecha Hacia Abajo
                PasarSigReg
                Me.txtAux(Index).SelStart = 0
                Me.txtAux(Index).SelLength = Len(Me.txtAux(Index).Text)
                'txtaux.Refresh
    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)

'    KEYpress KeyAscii
    
   If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        ModificarExistencia
        
        PasarSigReg
        
   ElseIf KeyAscii = 27 Then
        cmdCancelar_Click 'ESC
   End If
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim Importe As Currency
    If Screen.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    With txtAux(Index)
        .Text = Trim(.Text)
        If .Text = "" Then
            .Text = "0,00"
        Else
                If Not EsNumerico(.Text) Then
                    MsgBox "Importes deben ser numéricos.", vbExclamation
                    On Error Resume Next
                    .Text = "0,00"
                    PonFoco txtAux
                    Exit Sub
                End If
                
                
                'Es numerico
                If InStr(1, .Text, ",") > 0 Then
                    cad = ImporteFormateado(.Text)
                Else
                    cad = TransformaPuntosComas(.Text)
                End If
                If CadenaCurrency(cad, Importe) Then .Text = Format(Importe, "0.00")
                    
                
        
        End If
    End With

End Sub






Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim B As Boolean
       
    Modo = Kmodo
  
    

    B = (Modo = 0) Or (Modo = 2)
    PonerBotonCabecera B
   
    Select Case Kmodo
'    Case 0    'Modo Inicial
'        PonerBotonCabecera True
'        lblIndicador.Caption = ""
        
    Case 1 'Modo Buscar
'        PonerBotonCabecera False
      
'        lblIndicador.Caption = "BÚSQUEDA"
'    Case 2    'Visualización de Datos
'        PonerBotonCabecera True
'    Case 3 'Insertar Datos en el Datagrid
'        PonerBotonCabecera False 'Poner Aceptar y Cancelar Visible
'        lblIndicador.Caption = "MODIFICAR"
    End Select

    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub







Private Sub BotonModificar()
    If Data1.Recordset.EOF Then Exit Sub
    PonerModo 4
    CargaTxtAux True, True
    PonleFoco txtAux
End Sub


Private Function DatosOK() As Boolean
'Solo se actualiza el campo de Existencia Real
    txtAux(0).Text = Trim(txtAux(0).Text)
    DatosOK = False
    If txtAux(0).Text <> "" Then
        If EsNumerico(txtAux(0).Text) Then DatosOK = True
    End If
End Function


Private Sub PonerBotonCabecera(B As Boolean)
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    If B Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerOpcionesMenu()
    'PonerOpcionesMenuGeneral Me
End Sub




Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If DataGrid1.Bookmark < Data1.Recordset.RecordCount Then
'        DataGrid1.Row = DataGrid1.Row + 1
        DataGrid1.Bookmark = DataGrid1.Bookmark + 1
        PonleFoco Me.txtAux(0)
    ElseIf DataGrid1.Bookmark = Data1.Recordset.RecordCount Then
       PonleFoco cmdAceptar
    End If
    

End Sub


Private Function ModificarExistencia() As Boolean
Dim NumReg As Long


    If DatosOK Then
        
        If ActualizarExistencia() Then
            
            NumReg = Data1.Recordset.AbsolutePosition
            CargaGrid
            
                    
            If NumReg < Data1.Recordset.RecordCount Then
                Data1.Recordset.Move NumReg - 1
            Else
                Data1.Recordset.MoveLast
            End If
        End If

            
            
            ModificarExistencia = True
    Else
            ModificarExistencia = False
  
    End If
End Function




Private Function ActualizarExistencia() As Boolean
'Actualiza la cantidad de stock Inventariada (Existencia Real en Almacen)
Dim Sql As String
Dim Cantidad As Currency


    On Error GoTo EActualizar

    
        

    If InStr(1, txtAux(0).Text, ",") > 0 Then
        cad = ImporteFormateado(txtAux(0).Text)
    Else
        cad = TransformaPuntosComas(txtAux(0).Text)
    End If
    Cantidad = TransformaPuntosComas(cad)
    
   
    
    If Cantidad <> Data1.Recordset!Base Then
    
        
        Sql = "UPDATE importanatmptotal SET base =  " & TransformaComasPuntos(CStr(Cantidad))
        Cantidad = (Cantidad * Data1.Recordset!porceniva) / 100
        Cantidad = Round2(Cantidad, 2)
        Sql = Sql & " , iva=" & TransformaComasPuntos(CStr(Cantidad))
        Sql = Sql & " , Modificad=1"
        Sql = Sql & " WHERE seccion = " & Data1.Recordset!seccion & " AND porceniva="
        Sql = Sql & TransformaComasPuntos(CStr(Data1.Recordset!porceniva))
        Conn.Execute Sql
        
        
       
        
    End If
        
EActualizar:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
         MuestraError Err.Number, Sql, Err.Description
         ActualizarExistencia = False
    Else
        ActualizarExistencia = True
    End If
End Function



Private Sub CargaDatosFra()
    Set miRsAux = New ADODB.Recordset
    cad = "select tienda,descripcion ,fechafac, numfac from importnavtmp,importnavcentros where tienda=codcentro and secuencial=1"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        Me.txtDatos(0).Text = "Error"
        Me.txtDatos(1).Text = "Error"
        Me.txtDatos(2).Text = "Error"
        Me.txtDatos(3).Text = "Error"
        Me.Text3.Text = txtDatos(2).Text
        Me.txtDatos(5).Text = "Error"
        Me.cmdAceptar.Enabled = False
    Else
        Me.txtDatos(0).Text = miRsAux.Fields(0)
        Me.txtDatos(1).Text = miRsAux.Fields(1)
        Me.txtDatos(2).Text = Format(miRsAux.Fields(2), "dd/mm/yyyy")
        Me.txtDatos(3).Text = miRsAux.Fields(3)
        Me.Text3.Text = txtDatos(2).Text
        Me.txtDatos(5).Text = miRsAux.Fields(0) & " " & miRsAux.Fields(1)
        Me.cmdAceptar.Enabled = True
    End If
    miRsAux.Close
    
End Sub


Private Function DatosCorrecto() As Boolean
    
    DatosCorrecto = False
    
    If Not EsFechaOK(Text3) Then
        MsgBox "Fecha incorrecta: " & Text3.Text, vbExclamation
        Text3.Text = ""
        PonFoco Text3
        Exit Function
    End If
            
            
    'Vamos a ver si NO existe una factura de CONSUM con ese numero de facvura
    cad = DevuelveDesdeBD("Ctaconsum", "importnavparam ", "1", "1")
    cad = "codmacta = '" & cad & "' AND fecharec>='" & Year(vParam.fechaini) & "-01-01' AND numfactu"
    cad = DevuelveDesdeBD("numfactu", "factpro", cad, Me.txtDatos(3).Text, "T")
    If cad <> "" Then
        cad = "Ya existe la factura de consum nº:  " & Me.txtDatos(3).Text & vbCrLf & vbCrLf
        cad = cad & vbCrLf & "¿Continuar?"
        If MsgBox(cad, vbQuestion + vbYesNo) <> vbYes Then Exit Function
    End If
        
            
            
    'Hay que comprobar que las fechas estan
    'en los ejercicios y si
    '       0 .- Año actual
    '       1 .- Siguiente
    '       2 .- Anterior al inicio
    '       3 .- Posterior al fin
'    FechaCorrecta2 (CDate(Text3.Text))
'    If ModificandoLineas > 1 Then
'        If ModificandoLineas = 2 Then
'            RC = varTxtFec
'        Else
'            If ModificandoLineas = 3 Then
'                RC = "ya esta cerrado"
'            Else
'                RC = " todavia no ha sido abierto"
'            End If
'            RC = "La fecha pertenece a un ejercicio que " & RC
'        End If
'        MsgBox RC, vbExclamation
'    End If
    If FechaCorrecta2(CDate(Text3.Text)) > 1 Then
        MsgBox "Fecha fuera de ejercicio", vbExclamation
        Exit Function
    End If

    
    'Primero pondremos la fecha a año periodo
    
    If vParam.periodos = 0 Then
        'Trimestral
        NumRegElim = ((Month(CDate(Text3.Text)) - 1) \ 3) + 1
    Else
        NumRegElim = Month(CDate((Text3.Text)))
    End If
    
    If Year(CDate(Text3.Text)) < vParam.Anofactu Then
       NumRegElim = 0
    Else
        If Year(CDate(Text3.Text)) = vParam.Anofactu Then
            'El mismo año. Comprobamos los periodos
            If vParam.perfactu >= NumRegElim Then NumRegElim = 0
        End If
    End If
    
    If NumRegElim = 0 Then
        
        cad = "La factura corresponde a un periodo ya liquidado. " & vbCrLf
        cad = cad & vbCrLf & " ¿Desea continuar igualmente ?"
      
        If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
        
    
    
    
    DatosCorrecto = True
End Function
