VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTESRefTalonPagos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Introducci�n de Referencia de Talones"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   10770
   Icon            =   "frmTESRefTalonPagos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   15
      Tag             =   "Referencia|T|S|||tmppagos2|codmacta|||"
      Text            =   "1234567"
      Top             =   4440
      Width           =   1605
   End
   Begin VB.TextBox txtAux 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   720
      MaxLength       =   30
      TabIndex        =   14
      Tag             =   "Serie|T|S|||tmppagos2|numserie||S|"
      Top             =   4440
      Width           =   795
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   12
      Top             =   60
      Width           =   1095
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   13
         Top             =   210
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3600
      MaxLength       =   7
      TabIndex        =   3
      Tag             =   "Orden|N|S|||tmppagos2|numorden|######0|S|"
      Text            =   "1234567"
      Top             =   4440
      Width           =   915
   End
   Begin VB.TextBox txtAux 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Fecha Vto|F|S|||tmppagos2|fecefect|dd/mm/yyyy|S|"
      Text            =   "1234567890"
      Top             =   4440
      Width           =   1065
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   5850
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "Referencia|T|S|||tmppagos2|reftalonpag|||"
      Text            =   "1234567"
      Top             =   4440
      Width           =   1605
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1590
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "Factura|T|N|||tmppagos2|numfactu||S|"
      Top             =   4440
      Width           =   705
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "FecFactu|F|N|||tmppagos2|fecfactu|dd/mm/yyyy|S|"
      Top             =   4440
      Width           =   825
   End
   Begin VB.TextBox txtAux 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   250
      TabIndex        =   0
      Tag             =   "Usuario|N|N|||tmppagos2|codusu||S|"
      Top             =   4440
      Width           =   525
   End
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
      Left            =   8400
      TabIndex        =   6
      Tag             =   "   "
      Top             =   5550
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
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
      Left            =   9540
      TabIndex        =   7
      Top             =   5535
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmTESRefTalonPagos.frx":000C
      Height          =   4410
      Left            =   120
      TabIndex        =   10
      Top             =   885
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   7779
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   9540
      TabIndex        =   11
      Top             =   5550
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   5430
      Width           =   2385
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   40
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   8040
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmTESRefTalonPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1


'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadB As String

Dim Ordenacion As String

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la b�squeda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edici� del camp
'   3.-  Inserci� de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Integer

Dim FechaAnt As String
Dim Ok As Boolean
Dim CadB1 As String
Dim FILTRO As Byte
Dim SQL As String


Dim CadB2 As String
Dim RefAnt As String



Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If

    
    For i = 0 To txtaux.Count - 1
        txtaux(i).visible = (Modo = 1)
        txtaux(i).Enabled = (Modo = 1)
    Next i
    
    i = 6
        txtaux(i).visible = (Modo = 1 Or Modo = 4)
        txtaux(i).Enabled = (Modo = 1 Or Modo = 4)
   
    
    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.Adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim B As Boolean

    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(1).Enabled = B
    Me.mnModificar.Enabled = B
    
    
End Sub


Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 870 '670 '545
    End If

    'Llamamos al form
    txtaux(0).Text = DataGrid1.Columns(0).Text
    txtaux(1).Text = DataGrid1.Columns(1).Text
    txtaux(2).Text = DataGrid1.Columns(2).Text
    txtaux(3).Text = DataGrid1.Columns(3).Text
    txtaux(4).Text = DataGrid1.Columns(4).Text
    txtaux(7).Text = DataGrid1.Columns(5).Text
    txtaux(5).Text = DataGrid1.Columns(6).Text
    txtaux(6).Text = DataGrid1.Columns(7).Text
    If txtaux(6).Text = "" Then txtaux(6).Text = RefAnt

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonFoco txtaux(6)
    Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For i = 0 To txtaux.Count - 1
        txtaux(i).top = alto
    Next i
    ' ### [Monica] 12/09/2006
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub cmdAceptar_Click()
    Dim i As String
    Dim NReg As Long
    Dim SQL As String
    Dim Sql2 As String
    
    Select Case Modo
        Case 1 'BUSQUEDA
            
        Case 3 'INSERTAR
            If DatosOK Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
                        If Not Adodc1.Recordset.EOF Then
                            Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " =" & NuevoCodigo)
                        End If
                        cmdRegresar_Click
                    Else
'                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            Ok = False
            If DatosOK Then
                If ModificaDesdeFormulario2(Me, 0) Then
                    Ok = True
                    FechaAnt = txtaux(4).Text
                    TerminaBloquear
                    i = Adodc1.Recordset.AbsolutePosition
                    
                    PonerModo 2
                    CargaGrid "" 'CadB
                    If i = Me.Adodc1.Recordset.RecordCount Then
                        Me.Adodc1.Recordset.MoveLast
                    Else
                        Me.Adodc1.Recordset.Move i - 1
                    End If
                    
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'b�squeda
            CargaGrid CadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
        Case 4 'modificar
            TerminaBloquear
    End Select
    
    PonerModo 2
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If Adodc1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            cad = cad & Adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim cad As String

If Adodc1.Recordset Is Nothing Then Exit Sub
If Adodc1.Recordset.EOF Then Exit Sub

Me.Refresh
Screen.MousePointer = vbHourglass

Ordenacion = "ORDER BY " & DataGrid1.Columns(0).DataField

CargaGrid CadB

Screen.MousePointer = vbDefault
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 2 Then PonerContRegIndicador 'lblIndicador, adodc1, CadB
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'            BotonAnyadir
        Else
            PonerModo 2
             If Me.CodigoActual <> "" Then
                SituarData Me.Adodc1, "hidrante='" & CodigoActual & "'", "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
Dim Sql2 As String

    PrimeraVez = True

    With Toolbar1
         .HotImageList = frmppal.imgListComun_OM
         .DisabledImageList = frmppal.imgListComun_BN
         .ImageList = frmppal.ImgListComun
         .Buttons(1).Image = 4
    End With


    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT aaa.codusu, aaa.numserie, aaa.numfactu, aaa.fecfactu, aaa.numorden, aaa.codmacta, aaa.fecefect, aaa.reftalonpag"
    CadenaConsulta = CadenaConsulta & " FROM tmppagos2 aaa "
    CadenaConsulta = CadenaConsulta & " WHERE codusu = " & vUsu.Codigo
    '************************************************************************
    
    Ordenacion = " ORDER BY aaa.fecefect "
    
    
    CadB = ""
    CargaGrid
    
    FechaAnt = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    
    
    If Modo = 4 Then TerminaBloquear
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtaux(4).Text = Format(vFecha, "dd/mm/yyyy") '<===
End Sub


Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If Adodc1.Recordset.EOF Then Exit Sub
    
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    'Preparamos para modificar
    '-------------------------
    If BLOQUEADesdeFormulario2(Me, Adodc1, 1) Then BotonModificar
End Sub


Private Sub mnSalir_Click()
    Unload Me
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            mnModificar_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSql As String)
    Dim SQL As String
    Dim tots As String
    Dim Sql2 As String
    
    If vSql <> "" Then
        SQL = CadenaConsulta & " AND " & vSql
    Else
        SQL = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    SQL = SQL & " " & Ordenacion
    '**************************************************************++
    
    
    CargaGridGnral Me.DataGrid1, Me.Adodc1, SQL, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "N||||0|;S|txtAux(1)|T|Serie|700|;S|txtAux(2)|T|Factura|1000|;S|txtAux(3)|T|Fecha Fra.|1250|;"
    tots = tots & "S|txtAux(4)|T|Orden|1000|;"
    tots = tots & "S|txtAux(7)|T|Cuenta|1500|;"
    tots = tots & "S|txtAux(5)|T|Fecha Vto|1250|;S|txtAux(6)|T|Referencia|2500|;"
'    tots = tots & "S|txtAux(7)|T|Banco|1500|;"
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(2).Alignment = dbgLeft
    DataGrid1.Columns(4).Alignment = dbgLeft
    DataGrid1.Columns(7).Alignment = dbgCenter
'    DataGrid1.Columns(8).Alignment = dbgLeft
    
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFocoLin txtaux(Index)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 1, 2 ' <1> = socio <2> = partida
            PonerFormatoEntero txtaux(Index)
        
        Case 3, 5 ' lectura anterior / lectura actual
            PonerFormatoEntero txtaux(Index)
             
        Case 4 ' fecha de lectura actual
            '[Monica]28/08/2013: no comprobamos que la fecha est� en la campa�a
            PonerFormatoFecha txtaux(Index)
            
        Case 6
            RefAnt = txtaux(Index)
       
    End Select
    
End Sub

Private Function DatosOK() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim SQL As String
Dim Mens As String
Dim FechaAnt As Date
Dim NroDig As Integer
Dim Inicio As Long
Dim Fin As Long
Dim Consumo As Long
Dim Limite As Long

    B = CompForm(Me)
    If Not B Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
         If ExisteCP(txtaux(0)) Then B = False
    End If
    
    If B And Modo = 4 Then
    
    End If
    
    
    
    DatosOK = B
End Function

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
'    PonerOpcionesMenuGeneralNew Me
End Sub



Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index = 6 Then ' estoy introduciendo la lectura
       If KeyAscii = 13 Then 'ENTER
            If Modo = 4 Then
                RefAnt = txtaux(6)
                
                
                cmdAceptar_Click
                'ModificarLinea
                
                If Ok Then PasarSigReg
                
                
                    
            End If
            If Modo = 1 Or Modo = 3 Then
                cmdAceptar.SetFocus
            End If
            
       ElseIf KeyAscii = 27 Then
            cmdCancelar_Click 'ESC
       End If
    Else
        KEYpress KeyAscii
    End If

End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    
'Alvan�ar/Retrocedir els camps en les fleches de despla�ament del teclat.
'    KEYdown KeyCode
    
    
    ' si no estamos en muestra salimos
    If Index <> 7 Then Exit Sub
    
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
            cmdAceptar_Click
            If Ok Then PasarAntReg
        Case 40 'Desplazamiento Flecha Hacia Abajo
            'ModificarExistencia
                
            RefAnt = txtaux(6)
            
            cmdAceptar_Click
'            ModificarLinea
            
            If Ok Then PasarSigReg
    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If Me.DataGrid1.Bookmark < Me.Adodc1.Recordset.RecordCount Then
        DataGrid1.Bookmark = DataGrid1.Bookmark + 1
        BotonModificar
        PonFoco txtaux(6)
    ElseIf DataGrid1.Bookmark = Adodc1.Recordset.RecordCount Then
        PonerFocoBtn cmdAceptar
    End If
End Sub

Private Sub PasarAntReg()
'Nos situamos en el siguiente registro
    If Me.DataGrid1.Bookmark > 1 Then
        DataGrid1.Bookmark = DataGrid1.Bookmark - 1
        BotonModificar
        PonFoco txtaux(6)
    ElseIf DataGrid1.Bookmark = 1 Then
        BotonModificar
        PonFoco txtaux(6)
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


