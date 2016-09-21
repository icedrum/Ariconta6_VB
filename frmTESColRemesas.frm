VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTESColRemesas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remesas"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   12555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   960
      Top             =   6120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11220
      TabIndex        =   2
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11220
      TabIndex        =   5
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   3
      Top             =   5895
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2550
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Ver todos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Crear soporte magnetico"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar historico"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar remesa y vencimientos"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7080
         TabIndex        =   0
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   5295
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   6000
      Width           =   6615
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnFiltro 
      Caption         =   "Filtro"
      Visible         =   0   'False
      Begin VB.Menu mnFiltro1 
         Caption         =   "Efectos"
         Checked         =   -1  'True
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Pagarés"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Talones"
         Checked         =   -1  'True
         Index           =   3
      End
   End
   Begin VB.Menu mnOrdenacion 
      Caption         =   "Ordenacion"
      Begin VB.Menu mnOrdenacion1 
         Caption         =   "Tipo, codigo, año (Desc)"
         Index           =   0
      End
      Begin VB.Menu mnOrdenacion1 
         Caption         =   "Tipo, codigo, año (Asc)"
         Index           =   1
      End
      Begin VB.Menu mnOrdenacion1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnOrdenacion1 
         Caption         =   "Año, codigo, Tipo (Desc)"
         Index           =   3
      End
      Begin VB.Menu mnOrdenacion1 
         Caption         =   "Año, codigo, Tipo (Asc)"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmTESColRemesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public Tipo As Byte
    '1:  EFECTOS
    '2:  Talones y pagares
Private CadenaConsulta As String
Dim Modo As Byte
Dim Ordenacion As Byte
'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas  INSERTAR
'   Modo 2 -> Lineas MODIFICAR
'   Modo 3 -> Lineas BUSCAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
Modo = vModo

B = (Modo = 0)

Toolbar1.Buttons(1).Enabled = B
Toolbar1.Buttons(2).Enabled = B
Toolbar1.Buttons(8).Enabled = B
Toolbar1.Buttons(7).Enabled = B
Toolbar1.Buttons(6).Enabled = B
Toolbar1.Buttons(14).Enabled = B
        
'Prueba
cmdAceptar.Visible = Not B
cmdCancelar.Visible = Not B
DataGrid2.Enabled = B

'Si es regresar
If DatosADevolverBusqueda <> "" Then
    cmdRegresar.Visible = B
End If
End Sub




Private Sub BotonAnyadir()
    
    If vUsu.Nivel > 1 Then Exit Sub
    
    frmVarios.Opcion = 4
    'Si son efectos o NO
    If Tipo = 1 Then
        frmTESVarios.SubTipo = vbTipoPagoRemesa
    Else
        frmTESVarios.SubTipo = vbTalon
    End If
    frmTESVarios.Show vbModal, Me
    
    espera 0.5
    
    CargaGrid

End Sub



Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
End Sub


'0.- Modificar recibo
'1.- Crear dislette
Private Sub BotonModificar(vOp As Byte)
Dim I As Integer


    If vUsu.Nivel > 1 Then Exit Sub

    If adodc1.Recordset.EOF Then Exit Sub
    'Si tiporemesa NO es efecto, NO genera diskett ni na
    If Val(adodc1.Recordset!Tiporem) <> 1 And vOp = 1 Then
        MsgBox "No hay soporte fisico para talones / pagarés", vbExclamation
        Exit Sub
    End If
    
    'Consideraciones previas
    '----------------------------
    'Si es modificar rcibos o para los talones vtos, modificar cuenta banco
    If vOp <= 1 Then
        CadenaDesdeOtroForm = ""
        If Val(adodc1.Recordset!Tiporem) = 1 Then
            If Asc(UCase(adodc1.Recordset!Situacion)) > Asc("B") Then CadenaDesdeOtroForm = "No se puede modificar la remesa en esta situacion"
        Else
            If Asc(UCase(adodc1.Recordset!Situacion)) <> Asc("F") Then CadenaDesdeOtroForm = "Debe estar en cancelacion cliente"
        End If
        If CadenaDesdeOtroForm <> "" Then
            MsgBox CadenaDesdeOtroForm, vbExclamation
            Exit Sub
        End If
    End If
    
    
    If BloqueoManual(True, "ModRemesas", CStr(adodc1.Recordset!Codigo & "/" & adodc1.Recordset!Anyo)) Then

        If Val(adodc1.Recordset!Tiporem) > 1 Then
            frmVarios.Opcion = 25
            frmVarios.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                
                'Hacemos el cambio de valores
                Conn.BeginTrans
                If Not HacerUpdateRemTalon Then
                    CadenaDesdeOtroForm = ""
                    Conn.RollbackTrans
                Else
                    Conn.CommitTrans
                    espera 0.2
                    CadenaDesdeOtroForm = "OK" 'para que refresque el grid
                End If
            End If
        Else
            CadenaDesdeOtroForm = adodc1.Recordset!Codigo & "|" & adodc1.Recordset!Anyo & "|" & adodc1.Recordset!Situacion & "|" & adodc1.Recordset!fecremesa & "|"
            If vOp = 0 Then
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "|"
                frmVarios.Opcion = 6  'o lo k sea
                
            Else
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & adodc1.Recordset!codmacta & "|"
                frmVarios.Opcion = 7
            End If
            
            'Indicamos tb el tipo de remesa
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & adodc1.Recordset!DescripcionT & "|" & adodc1.Recordset!Tiporem & "|"
            
            frmVarios.Show vbModal
    
        End If
    
        'Hay que poner en el formualrio de arriba valor a cadenadesdeotroform si ha modificado
        If CadenaDesdeOtroForm <> "" Then CargaGrid
                            
        
        'Desbloqueamos
        BloqueoManual False, "ModRemesas", ""
    
    Else
        MsgBox "Registro bloqueado", vbExclamation
    End If

End Sub


Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    
    'Eliminar la rmesa si esta en sitauacion A,B
    
    
    If vUsu.Nivel > 1 Then Exit Sub
    
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    
    'Boqueo, borro y sigo
    If Val(adodc1.Recordset!Tiporem) = 2 Then
        SQL = "Pagaré"
    ElseIf Val(adodc1.Recordset!Tiporem) = 3 Then
        SQL = "Talón"
    Else
        SQL = "Efectos"
    End If
    SQL = vbCrLf & "Tipo :  " & SQL
    SQL = "Seguro que desea eliminar la remesa:" & SQL
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset!Codigo
    SQL = SQL & vbCrLf & "Año: " & adodc1.Recordset!Anyo
    SQL = SQL & vbCrLf & "Banco: " & adodc1.Recordset!codmacta & " " & adodc1.Recordset!Nommacta
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        If BloqueoManual(True, "Remesas", "Remesas") Then
            'Hay que eliminar
            
            If Tipo = 1 Then
            
            
                SQL = "Delete from remesas where codigo=" & adodc1.Recordset!Codigo
                SQL = SQL & " AND anyo =" & adodc1.Recordset!Anyo
                SQL = SQL & " AND tiporem =" & adodc1.Recordset!Tiporem
                Conn.Execute SQL
            
                'Agosto2013  Ponemos a null la cuenta real de cobroctabanc2
                'Pongo A NULL todos los recibos con esos valores
                SQL = "UPDATE scobro set codrem=NULL,anyorem=NULL,siturem=NULL,tiporem=NULL"
                SQL = SQL & ",fecultco=NULL,impcobro=NULL,ctabanc2=NULL"
                SQL = SQL & " where codrem=" & adodc1.Recordset!Codigo
                SQL = SQL & " AND anyorem =" & adodc1.Recordset!Anyo
                SQL = SQL & " AND tiporem =" & adodc1.Recordset!Tiporem
                Conn.Execute SQL
            
            Else
                BorrarRemesaEnCancelacionTalonesPagares
            End If
            CargaGrid ""
            adodc1.Recordset.Cancel
            BloqueoManual False, "Remesas", ""
        
        Else
            MsgBox "Proceso bloqueado por otro usuario", vbExclamation
        End If
    End If
    Exit Sub
Error2:
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1
    If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
    
Case 3
    CargaGrid
End Select
PonerModo 0
lblIndicador.Caption = ""
DataGrid2.SetFocus
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String

If adodc1.Recordset.EOF Then
    MsgBox "Ningún registro a devolver.", vbExclamation
    Exit Sub
End If

cad = adodc1.Recordset.Fields(1) & "|"
cad = cad & adodc1.Recordset.Fields(2) & "|"
cad = cad & adodc1.Recordset.Fields(3) & "|"


RaiseEvent DatoSeleccionado(cad)
Unload Me
End Sub

Private Sub cmdRegresar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid2_DblClick()

If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
          ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 28
        .Buttons(11).Image = 16
        .Buttons(12).Image = 14
        
        .Buttons(14).Image = 26 'ELiminar rem y vtos
        
        .Buttons(16).Image = 15
    End With
    
    If Tipo = 2 Then
        Caption = Caption & "       PAGARES y TALONES"
        Label1.Caption = "Talones - Pagarés"
        Label1.Alignment = 0
    Else
        Label1.Caption = "Efectos"
        Label1.Alignment = 1
    End If
    'Para talones y pagares
    mnOrdenacion1(0).Visible = Tipo = 2
    mnOrdenacion1(1).Visible = Tipo = 2
    mnOrdenacion1(2).Visible = Tipo = 2
    Me.mnFiltro.Visible = Tipo = 2
    
    'Para efctos
    Toolbar1.Buttons(10).Visible = Tipo = 1   'Generar disquette
    Toolbar1.Buttons(14).Visible = Tipo = 1   'Elimanar remesa y efectos (REGAIXO)
    
    
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    'DespalzamientoVisible False
    PonerModo 0
    
    
    
    LeerGuardarOrdenacion True
    
    CadenaConsulta = "Select DescripcionT,codigo,anyo, fecremesa,tiporemesa.descripcion,descsituacion,remesas.codmacta,nommacta,"
    CadenaConsulta = CadenaConsulta & " Importe , remesas.descripcion, remesas.Tipo,situacion,tiporem"
    CadenaConsulta = CadenaConsulta & " from cuentas,tiporemesa2,tiposituacionrem,remesas left join tiporemesa on remesas.tipo=tiporemesa.tipo where remesas.codmacta=cuentas.codmacta"
    CadenaConsulta = CadenaConsulta & " and situacio=situacion and tiporemesa2.tipo=remesas.tiporem"
    
    Set DataGrid2.DataSource = adodc1
    CargaGrid
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    LeerGuardarOrdenacion False
End Sub

Private Sub LeerGuardarOrdenacion(Leer As Boolean)
Dim NF As Integer
    On Error GoTo ELeerGuardarOrdenacion
    
    
    CadenaConsulta = App.Path & "\OrdenRem.xdf"
      If Leer Then
            Ordenacion = 0
            If Dir(CadenaConsulta, vbArchive) <> "" Then
                'Existe el fichero
                NF = FreeFile
                Open CadenaConsulta For Input As #NF
                Line Input #NF, CadenaConsulta
                Close #NF
                If CadenaConsulta = "" Then CadenaConsulta = "0"
                If IsNumeric(CadenaConsulta) Then Ordenacion = CByte(Val(CadenaConsulta))
            End If
            
            If Ordenacion > 0 And Ordenacion < 5 Then
                Me.mnOrdenacion1(Ordenacion).Checked = True
            Else
                Me.mnOrdenacion1(Ordenacion).Checked = True
            End If
      Else
            If Ordenacion = 0 Then
                If Dir(CadenaConsulta, vbArchive) <> "" Then Kill CadenaConsulta
            Else
                NF = FreeFile
                Open CadenaConsulta For Output As #NF
                Print #NF, Ordenacion
                Close #NF
            End If
      End If
    Exit Sub
ELeerGuardarOrdenacion:
    MuestraError Err.Number, "LeerGuardarOrdenacion"
          
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnFiltro1_Click(Index As Integer)
    mnFiltro1(Index).Checked = Not mnFiltro1(Index).Checked
    CargaGrid
End Sub

Private Sub mnModificar_Click()
    BotonModificar 0
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnOrdenacion1_Click(Index As Integer)
Dim N As Integer
    For N = 0 To mnOrdenacion1.Count - 1
        'El 3 es la barra
        If N <> 2 Then mnOrdenacion1(N).Checked = False
    Next N
    mnOrdenacion1(Index).Checked = True
    If Ordenacion <> CByte(Index) Then
        Ordenacion = CByte(Index)
        CargaGrid
    End If
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

If Button.Index < 16 Then
    If DatosADevolverBusqueda <> "" Then
        MsgBox "Esta seleccionando una remesa. No puede modificar nada de ellas.", vbExclamation
        Exit Sub
    End If
End If
Select Case Button.Index
Case 1
        BotonBuscar
Case 2
        BotonVerTodos
Case 6
        BotonAnyadir
Case 7
        BotonModificar 0
Case 8
        BotonEliminar
        
        
Case 10
        If vUsu.Nivel > 1 Then Exit Sub
        BotonModificar 1

Case 11
        frmTESListado.Opcion = 6
        frmTESListado.Show vbModal
        
Case 12
        If vUsu.Nivel > 1 Then Exit Sub
        'Borraremos lo que serian las cabceceras de la remesas
        If Tipo = 1 Then
            frmTESVarios.SubTipo = vbTipoPagoRemesa
        Else
            frmTESVarios.SubTipo = vbTalon 'O pagare, daria lo mismo)
        End If
        frmTESVarios.Opcion = 17
        frmTESVarios.Show vbModal
        CargaGrid ""
        
Case 14
        'Eliminar remesa y VTO
        If Not Me.adodc1.Recordset.EOF Then BorrarRemesaVtos
Case 16
        Unload Me
        

Case Else

End Select
End Sub

Private Sub CargaGrid(Optional SQL As String)
   ' Dim J As Integer
    
    Dim I As Integer
    DataGrid2.EditActive = False
    DataGrid2.AllowUpdate = False
    adodc1.ConnectionString = Conn
    If SQL <> "" Then
        SQL = CadenaConsulta & " AND " & SQL
        Else
        SQL = CadenaConsulta
    End If
    
    SQL = SQL & PonerOrdenFiltro
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid2.AllowRowSizing = False
    DataGrid2.RowHeight = 290
    
    
    I = 0
    DataGrid2.Columns(I).Caption = "Tipo"
    DataGrid2.Columns(I).Width = 900
    DataGrid2.HeadFont.Bold = True
    
    I = 1
        DataGrid2.Columns(I).Caption = "Cod."
        DataGrid2.Columns(I).Width = 600
'        DataGrid2.Columns(i).NumberFormat = "000"
        
    
    'Leemos del vector en 2
    I = 2
        DataGrid2.Columns(I).Caption = "Año"
        DataGrid2.Columns(I).Width = 700
    
    'El importe es campo calculado
    I = 3
        DataGrid2.Columns(I).Caption = "Fecha"
        DataGrid2.Columns(I).Width = 1100
        DataGrid2.Columns(I).NumberFormat = "dd/mm/yyyy"
    
    
    DataGrid2.Columns(4).Caption = "Norma"
    DataGrid2.Columns(4).Width = 850
    
    I = 5
        DataGrid2.Columns(I).Caption = "Situación"
        DataGrid2.Columns(I).Width = 1300
        
   
    I = 6
    DataGrid2.Columns(I).Caption = "Cuenta"
    DataGrid2.Columns(I).Width = 1000
                
        
    I = 7
    DataGrid2.Columns(I).Caption = "Nombre"
    DataGrid2.Columns(I).Width = 1980
        
    I = 8
    DataGrid2.Columns(I).Caption = "Importe"
    DataGrid2.Columns(I).Width = 1100
    DataGrid2.Columns(I).Alignment = dbgRight
    DataGrid2.Columns(I).NumberFormat = FormatoImporte

    DataGrid2.Columns(9).Width = 2000
       
        
    DataGrid2.Columns(10).Visible = False
    DataGrid2.Columns(11).Visible = False
    DataGrid2.Columns(12).Visible = False
        
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
    End If
End Sub

Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub



Private Function SepuedeBorrar() As Boolean
'Dim SQL As String
    SepuedeBorrar = False
    
    If adodc1.Recordset!Situacion = "A" Or adodc1.Recordset!Situacion = "B" Then
        SepuedeBorrar = True
    Else
        
        If Tipo = 1 Then
            MsgBox "No se puede eliminar la remesa en esta situación: " & adodc1.Recordset!Situacion, vbExclamation
        Else
            'TALONES PAGARES
            If adodc1.Recordset!Situacion = "F" Then
                'En cancelacion si que dejo eliminar, ya que lo que se hace realmente es:
                '1.- QUitar la remesa de los cobros
                '2.- Quitar la remesa de la tabla remesas
                '3.- poner en scarecepdoc LlevadoBanco=0
                SepuedeBorrar = True
            Else
                MsgBox "No se puede eliminar la remesa en esta situación: " & adodc1.Recordset!Situacion, vbExclamation
            End If
        End If
    End If
End Function


Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
    
    While DataGrid2.SelBookmarks.Count > 0
        DataGrid2.SelBookmarks.Remove 0
    Wend
    Exit Sub
    
EDeseleccionaGrid:
        Err.Clear
End Sub


Private Sub KEYpress(KeyAscii As Integer)
    'Caption = KeyAscii
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            If Modo = 0 Then Unload Me
        End If
    End If
End Sub



Private Sub BorrarRemesaVtos()
Dim SQL As String

    If adodc1.Recordset.EOF Then Exit Sub
    
    If Val(adodc1.Recordset!Tiporem) > 1 Then
        MsgBox "Solo para efectos.", vbExclamation
        Exit Sub
    End If
    
    NumRegElim = 0
    SQL = "Select count(*) from scobro where codrem=" & adodc1.Recordset!Codigo
    SQL = SQL & " AND anyorem =" & adodc1.Recordset!Anyo
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    Set miRsAux = Nothing
    
    SQL = "Va a borrar la remesa y los vencimientos para: "
    SQL = SQL & vbCrLf & " --------------------------------------------------------------------"
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset!Codigo
    SQL = SQL & vbCrLf & "Año: " & adodc1.Recordset!Anyo
    SQL = SQL & vbCrLf & "Banco: " & adodc1.Recordset!codmacta & " " & adodc1.Recordset!Nommacta
    SQL = SQL & vbCrLf & "Situación: " & adodc1.Recordset!descsituacion
    SQL = SQL & vbCrLf & "Importe: " & Format(adodc1.Recordset!Importe, FormatoImporte)
    SQL = SQL & vbCrLf & "Vencimientos: " & NumRegElim
    SQL = SQL & vbCrLf & vbCrLf & "                         ¿Continuar?"
    NumRegElim = 0
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    SQL = "El proceso es irreversible"
    SQL = SQL & vbCrLf & "Desea continuar?"
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
    Screen.MousePointer = vbHourglass
    If HacerEliminacionRemesaVtos Then
        'Cargar datos
         CargaGrid ""
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub ObtenerNumeroVtosAsociados()
    NumRegElim = 0
    
End Sub
Private Function HacerEliminacionRemesaVtos() As Boolean

    On Error GoTo EHacerEliminacionRemesaVtos

    HacerEliminacionRemesaVtos = False

    'Eliminamos los vencimientos asociados
    Conn.Execute "DELETE FROM scobro where codrem=" & adodc1.Recordset!Codigo & " AND anyorem =" & adodc1.Recordset!Anyo
    
    'Eliminamos la remesa
    Conn.Execute "DELETE FROM remesas where codigo=" & adodc1.Recordset!Codigo & " AND anyo =" & adodc1.Recordset!Anyo
    
    HacerEliminacionRemesaVtos = True
    Exit Function
EHacerEliminacionRemesaVtos:
    MuestraError Err.Number, "Function: HacerEliminacionRemesaVtos"
End Function

Private Function PonerOrdenFiltro()
Dim C As String
    'Filtro
    If Tipo = 1 Then
        'REMESAS
        C = RemesaSeleccionTipoRemesa(True, False, False)
    Else
        'TALON Y PAGARE
        If Not Me.mnFiltro1(2).Checked And Not Me.mnFiltro1(3).Checked Then
             Me.mnFiltro1(2).Checked = True
              Me.mnFiltro1(3).Checked = True
        End If
        C = RemesaSeleccionTipoRemesa(False, Me.mnFiltro1(2).Checked, Me.mnFiltro1(3).Checked)
    End If
    
    If C <> "" Then C = " AND " & C
    Select Case Ordenacion
    Case 1
        PonerOrdenFiltro = "tiporem,anyo asc , codigo asc"
        'Tipo, codigo, año (Asc)   0 y 1desc
    Case 3
        PonerOrdenFiltro = "anyo desc , codigo desc,tiporem"
    Case 4
        PonerOrdenFiltro = "anyo asc , codigo asc,tiporem"
        
    Case Else
        'Por defecto
        PonerOrdenFiltro = "tiporem,anyo desc , codigo desc"
    End Select
    PonerOrdenFiltro = C & " ORDER BY " & PonerOrdenFiltro
End Function



Private Function BorrarRemesaEnCancelacionTalonesPagares() As Boolean
Dim C As String

    On Error GoTo EBorrarRemesaEnCancelacionTalonesPagares

    'En cancelacion si que dejo eliminar, ya que lo que se hace realmente es:
    '1.- QUitar la remesa de los cobros       'Estos dos puntos los hace en la otra
    '2.- Quitar la remesa de la tabla remesas
    '3.- poner en scarecepdoc LlevadoBanco=0
        
    BorrarRemesaEnCancelacionTalonesPagares = False

    'Veamos que scarecep son
    Set miRsAux = New ADODB.Recordset
    C = "select id from slirecepdoc where (numserie,numfaccl,fecfaccl,numvenci) IN ("
    C = C & "SELECT numserie,codfaccl,fecfaccl,numorden FROM scobro WHERE "
    C = C & " codrem=" & adodc1.Recordset!Codigo & " AND anyorem = " & adodc1.Recordset!Anyo & ")"
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        C = "UPDATE scarecepdoc set LlevadoBanco = 0 WHERE codigo = " & miRsAux!Id
        Conn.Execute C
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Ponemos los vencimientos sin remesa
    C = "UPDATE scobro SET codrem=NULL, anyorem=NULL,siturem=NULL where"
    C = C & " codrem=" & adodc1.Recordset!Codigo & " AND anyorem = " & adodc1.Recordset!Anyo
    Conn.Execute C
    
    'Borramos la remesa
    C = "DELETE from remesas WHERE "
    C = C & " Codigo=" & adodc1.Recordset!Codigo & " AND Anyo = " & adodc1.Recordset!Anyo
    Conn.Execute C
    BorrarRemesaEnCancelacionTalonesPagares = True
    Exit Function
EBorrarRemesaEnCancelacionTalonesPagares:
    MsgBox "Error grave. Consulte soporte técnico", vbExclamation
End Function



Private Function HacerUpdateRemTalon() As Boolean
Dim C As String
'CadenaDesdeOtroForm = fecha & "|" & cuenta banco & "|"
    On Error GoTo EHacerUpdateRemTalon
    HacerUpdateRemTalon = False
        
        
    C = RecuperaValor(CadenaDesdeOtroForm, 2)
    
    If C <> "" Then
        C = "UPDATE scobro set ctabanc2 ='" & C & "' WHERE codrem = " & adodc1.Recordset!Codigo
        C = C & " AND anyorem = " & adodc1.Recordset!Anyo & " AND tiporem =" & adodc1.Recordset!Tiporem
        Conn.Execute C
        
        
        C = RecuperaValor(CadenaDesdeOtroForm, 2)
        C = "UPDATE remesas set codmacta = '" & C & "' WHERE codigo = " & adodc1.Recordset!Codigo
        C = C & " AND anyo = " & adodc1.Recordset!Anyo & " AND tiporem =" & adodc1.Recordset!Tiporem
        Conn.Execute C
    End If
        
    'Fehca
    
    C = RecuperaValor(CadenaDesdeOtroForm, 1)
    If C <> "" Then
        C = "UPDATE remesas set fecremesa = '" & Format(C, FormatoFecha) & "' WHERE codigo = " & adodc1.Recordset!Codigo
        C = C & " AND anyo = " & adodc1.Recordset!Anyo & " AND tiporem =" & adodc1.Recordset!Tiporem
        Conn.Execute C
    End If
    'Fecha vto
    HacerUpdateRemTalon = True
    Exit Function
EHacerUpdateRemTalon:
    MuestraError Err.Number, Err.Description
End Function
