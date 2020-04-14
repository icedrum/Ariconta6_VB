VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAsiLinAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lineas Asientos Predefinidos"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAsiLinAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame FrameBotonGnral 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   1185
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   13
         Top             =   180
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Leer desde Predefinidos"
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   1
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   5
      Text            =   "haber"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   1290
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   810
      Width           =   7515
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   4
      Text            =   "debe"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   7215
      Width           =   3255
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8610
      TabIndex        =   0
      Top             =   7260
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9810
      TabIndex        =   1
      Top             =   7260
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9810
      TabIndex        =   3
      Top             =   7260
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1320
      Top             =   6720
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
      Bindings        =   "frmAsiLinAdd.frx":000C
      Height          =   5895
      Left            =   120
      TabIndex        =   10
      Top             =   1260
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   10398
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Ampliación"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label lblInfInv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   7380
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
      TabIndex        =   6
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnOpciones1 
         Caption         =   "Leer desde predefinido"
         Index           =   0
      End
      Begin VB.Menu mnOpciones1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnOpciones1 
         Caption         =   "Salir"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmAsiLinAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TotalLineas As Currency


Private WithEvents frmPre As frmBasico
Attribute frmPre.VB_VarHelpID = -1

Private Modo As Byte
Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
Dim PrimeraVez As Boolean
Dim cad As String
Dim PreguntarAmplia As Boolean
Dim PasoPorAmpliacion As Boolean

Private Sub cmdAceptar_Click()
'TotalLineas llevo
    Set miRsAux = New ADODB.Recordset
    cad = "Select count(*) from tmpconext WHERE codusu = " & vUsu.Codigo & " and (timported <> 0 or timporteh <> 0)"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        cad = "X"
    Else
        If DBLet(miRsAux.Fields(0), "N") = 0 Then
            cad = "X"
        Else
            cad = "Va a insertar en el Asiento:" & vbCrLf
            cad = cad & Space(10) & "Lineas: " & miRsAux.Fields(0) & vbCrLf
            cad = cad & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) = vbYes Then
                cad = ""
            Else
                cad = "NO"
            End If
        End If
    End If
    
    miRsAux.Close
    Set miRsAux = Nothing
    
    If cad <> "" Then
        If cad = "X" Then cad = "No hay valor para ninguna de las lineas"
        
        If cad <> "NO" Then MsgBox cad, vbExclamation
        
        Exit Sub
    End If
        
    CadenaDesdeOtroForm = "OK"
    Ampliacion = Text1.Text
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
'    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not data1.Recordset.EOF And gridCargado And Modo = 4 Then
       CargaTxtAux True, True
       
       txtaux(0).SetFocus
       
       txtaux(0).SelStart = 0
       txtaux(0).SelLength = Len(Me.txtaux(0).Text)
       txtaux(0).Refresh
       
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
         PrimeraVez = False
         mnOpciones1_Click 0
    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmppal.Icon
    
     ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 4
    End With
    
    Me.Tag = 1 'NO se puede cerrar mas que de boton
    
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If vParam.autocoste Then
        Me.Width = 11890 '10315
        
    End If
    DataGrid1.Width = Me.Width - 400
    Me.cmdCancelar.Left = Me.Width - 1365
    Me.cmdAceptar.Left = Me.Width - 2565
    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
    
    BorrarDatos
    
    CargaGrid
    PrimeraVez = True
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid()


On Error GoTo ECarga

    gridCargado = False
    
    cad = "select cta,nommacta, pos ,ccost ,ampconce,timported, timporteh"
    cad = cad & " from tmpconext,cuentas  where tmpconext.cta=cuentas.codmacta AND codusu=" & vUsu.Codigo & " ORDER BY pos"

    data1.ConnectionString = Conn
    data1.RecordSource = cad
    data1.CursorType = adOpenDynamic
    data1.LockType = adLockPessimistic
    data1.Refresh
   
    PrimeraVez = False
    
    DataGrid1.Columns(0).Caption = "Cuenta"
    DataGrid1.Columns(0).Width = 1300
    
    DataGrid1.Columns(1).Caption = "Descripcion"
    If vParam.autocoste Then
        DataGrid1.Columns(1).Width = 3700
    Else
        DataGrid1.Columns(1).Width = 6000
    End If
    
    DataGrid1.Columns(2).visible = False
    
    If Not vParam.autocoste Then
        DataGrid1.Columns(3).visible = False
        DataGrid1.Columns(4).visible = False
        
    Else
        DataGrid1.Columns(3).Caption = "C.C."
        DataGrid1.Columns(3).Width = 700
        DataGrid1.Columns(4).Caption = "Nombre centro coste"
        DataGrid1.Columns(4).Width = 2300
    End If
    
    DataGrid1.Columns(5).Caption = "Debe"
    DataGrid1.Columns(5).Width = 1400
    DataGrid1.Columns(5).NumberFormat = FormatoImporte
    DataGrid1.Columns(5).Alignment = dbgRight
            
    DataGrid1.Columns(6).Caption = "Haber"
    DataGrid1.Columns(6).Width = 1400
    DataGrid1.Columns(6).NumberFormat = FormatoImporte
    DataGrid1.Columns(6).Alignment = dbgRight
            
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.RowHeight = 350 '290

    gridCargado = True
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, Limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtaux(0).top = 290
        txtaux(0).visible = visible
    
        txtaux(1).top = 290
        txtaux(1).visible = visible
    
    Else
        DeseleccionaGrid Me.DataGrid1
        If Limpiar Then 'Vaciar los textBox (Vamos a Insertar)
                txtaux(0).Text = DBLet(data1.Recordset!timported)
                txtaux(0).Locked = False
                
                txtaux(1).Text = DBLet(data1.Recordset!timporteH)
                txtaux(1).Locked = False
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.top + 220
        Else
            alto = DataGrid1.top + DataGrid1.RowTop(DataGrid1.Row) + 20
        End If
        
        'Fijamos altura y posición Top
        '-------------------------------
        txtaux(0).top = alto
        txtaux(0).Height = DataGrid1.RowHeight
        txtaux(1).top = alto
        txtaux(1).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        txtaux(0).Left = DataGrid1.Columns(5).Left + 130 'codalmac
        txtaux(0).Width = DataGrid1.Columns(5).Width - 10
        txtaux(1).Left = DataGrid1.Columns(6).Left + 130 'codalmac
        txtaux(1).Width = DataGrid1.Columns(6).Width - 10
        
        'Los ponemos Visibles o No
        '--------------------------
        txtaux(0).visible = visible
        txtaux(1).visible = visible
    End If
    PonFoco txtaux(0)
    
    If visible Then
        txtaux(0).TabIndex = 2
        txtaux(1).TabIndex = 3
    Else
        txtaux(0).TabIndex = 5
        txtaux(1).TabIndex = 6
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Me.Tag = 1 Then Cancel = 1 'o aceptar o cancelar
End Sub

Private Sub frmPre_DatoSeleccionado(CadenaSeleccion As String)
    cad = CadenaSeleccion
End Sub


Private Sub mnOpciones1_Click(Index As Integer)
    If Index = 0 Then
        PreguntarAmplia = True
        If data1.Recordset.RecordCount > 0 Then
            If MsgBox("Ya existen datos. Volver a cargarlos?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            Text1.Text = ""
            PreguntarAmplia = False
            BorrarDatos
            espera 0.5
        End If
        cad = ""
        
        
        Set frmPre = New frmBasico
        
        AyudaAsientosP frmPre
        
        
        If cad <> "" Then
        
            NumAsiPre = RecuperaValor(cad, 1)
            
            cad = "numaspre=" & RecuperaValor(cad, 1)
            
            
            If vParam.autocoste Then
                'LLEVA CENTROS DE COSTE
                cad = "left join ccoste on asipre_lineas.codccost =ccoste.codccost WHERE " & cad
            Else
                cad = " WHERE  " & cad
            End If
            cad = " FROM asipre_lineas " & cad
            
            If vParam.autocoste Then cad = ",asipre_lineas.codccost,nomccost " & cad
            
            cad = ")  select " & vUsu.Codigo & ",codmacta,0,linlapre, timported, timporteh" & cad
            If vParam.autocoste Then cad = ",ccost, ampconce" & cad
            cad = "INSERT INTO tmpconext(codusu,cta,saldo,pos,timported,timporteh " & cad
            Conn.Execute cad
            CargaGrid
            PasoPorAmpliacion = False
            BotonModificar
        End If
    
    Else
    
    
    End If
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Text1_LostFocus()
    If Trim(Text1.Text) = "" And PreguntarAmplia Then
    End If
    PasoPorAmpliacion = True
    BotonModificar

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            mnOpciones1_Click 0
    End Select

End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    txtaux(Index).SelStart = 0
    txtaux(Index).SelLength = Len(txtaux(Index).Text)
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
                End If

        Case 40 'Desplazamiento Flecha Hacia Abajo
                PasarSigReg
                Me.txtaux(Index).SelStart = 0
                Me.txtaux(Index).SelLength = Len(Me.txtaux(Index).Text)
    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)


   If KeyAscii = 13 Then 'ENTER
        If Index = 0 And ComprobarCero(txtaux(0).Text) = 0 Then
            PonFoco txtaux(1)
            Exit Sub
        End If
        
        KeyAscii = 0
        
        If Index = 0 And ComprobarCero(txtaux(1).Text) <> 0 Then
            txtaux(1).Text = ""
        End If
        If Index = 1 And ComprobarCero(txtaux(1).Text) <> 0 Then
            txtaux(0).Text = ""
        End If
        
        ModificarExistencia

        PasarSigReg

   ElseIf KeyAscii = 27 Then
        cmdCancelar_Click 'ESC
   End If
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim Importe As Currency
    If Screen.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    With txtaux(Index)
        .Text = Trim(.Text)
        If .Text = "" Then
            .Text = "0,00"
        Else
            If Not EsNumerico(.Text) Then
                MsgBox "Importes deben ser numéricos.", vbExclamation
                On Error Resume Next
                .Text = "0,00"
                PonFoco txtaux(Index)
                Exit Sub
            End If

            'Es numerico
            cad = TransformaPuntosComas(.Text)
            If CadenaCurrency(cad, Importe) Then .Text = Format(Importe, "0.00")
        End If
    End With

    If txtaux(Index).Text = "0,00" Then txtaux(Index).Text = ""
    
    If Index = 1 And ComprobarCero(txtaux(1).Text) <> 0 Then
        txtaux(0).Text = ""
    End If
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim B As Boolean
       
    Modo = Kmodo

    B = (Modo = 0) Or (Modo = 2)
    PonerBotonCabecera B
    lblIndicador.Caption = "MODIFICAR"
   

    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub BotonModificar()
    If data1.Recordset.EOF Then Exit Sub
    
    If Not PasoPorAmpliacion Then
        PasoPorAmpliacion = True
        PonFoco Text1
        Exit Sub
    End If
    PonerModo 4
    CargaTxtAux True, True
    PonFoco txtaux(0)
End Sub


Private Function DatosOK() As Boolean
'Solo se actualiza el campo de Existencia Real
    txtaux(0).Text = Trim(txtaux(0).Text)
    txtaux(1).Text = Trim(txtaux(1).Text)
    DatosOK = False
    If txtaux(0).Text <> "" Or txtaux(1).Text <> "" Then
        If EsNumerico(ComprobarCero(txtaux(0).Text)) Or EsNumerico(ComprobarCero(txtaux(1).Text)) Then DatosOK = True
    Else
        txtaux(1).Text = 0
        DatosOK = True
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
    If DataGrid1.Bookmark < data1.Recordset.RecordCount Then
        DataGrid1.Bookmark = DataGrid1.Bookmark + 1
        PonFoco Me.txtaux(0)
    ElseIf DataGrid1.Bookmark = data1.Recordset.RecordCount Then
       PonleFoco cmdAceptar
    End If

End Sub


Private Function ModificarExistencia() As Boolean
Dim NumReg As Long


    If DatosOK Then
        
        If ActualizarExistencia() Then
            
            NumReg = data1.Recordset.AbsolutePosition
            CargaGrid
            
                    
            If NumReg < data1.Recordset.RecordCount Then
                data1.Recordset.Move NumReg - 1
            Else
                data1.Recordset.MoveLast
            End If
        End If

            
            
            ModificarExistencia = True
    Else
            ModificarExistencia = False
  
    End If
End Function




Private Function ActualizarExistencia() As Boolean
'Actualiza la cantidad de stock Inventariada (Existencia Real en Almacen)
Dim SQL As String
Dim Debe As Currency
Dim Haber As Currency

    On Error GoTo EActualizar


    Debe = TransformaPuntosComas(ComprobarCero(txtaux(0).Text))
    Haber = TransformaPuntosComas(ComprobarCero(txtaux(1).Text))
    
        SQL = "UPDATE tmpconext  Set timported = " & DBSet(Debe, "N", "S")
        SQL = SQL & ", timporteh = " & DBSet(Haber, "N", "S")
        SQL = SQL & " WHERE cta = '" & data1.Recordset!Cta & "' AND "
        SQL = SQL & " pos =" & data1.Recordset!Pos & " AND codusu =" & vUsu.Codigo
        Conn.Execute SQL
        
        
        
EActualizar:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
         MuestraError Err.Number, SQL, Err.Description
         ActualizarExistencia = False
    Else
        ActualizarExistencia = True
    End If
End Function


Private Sub BorrarDatos()
    Conn.Execute "DELETE FROM tmpconext WHERE codusu = " & vUsu.Codigo
End Sub


