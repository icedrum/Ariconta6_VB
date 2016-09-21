VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmTESActualizar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar diario"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "frmTESActualizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameResultados 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   5115
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   4200
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Left            =   60
         TabIndex        =   10
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6165
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº Asien"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha Entrada"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Obteniendo resultados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label8 
         Caption         =   "Errores:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.Frame frame1Asiento 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin ComCtl2.Animation Animation1 
         Height          =   735
         Left            =   600
         TabIndex        =   4
         Top             =   1800
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1296
         _Version        =   327681
         FullWidth       =   241
         FullHeight      =   49
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label9 
         Caption         =   "Label2"
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
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label lblAsiento 
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Asiento :"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmTESActualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public OpcionActualizar As Byte
    '1.- Actualizar 1 asiento
    '2.- Desactualiza pero NO insertes en apuntes
    '3.- Desactualizar asiento desde hco
    
    'Si el asiento es de una factura entonces NUMSERIE tendra "FRACLI" o "FRAPRO"
    ' con lo cual habra que poner su factura asociada a NULL
    
    '4.- Si es para enviar datos a impresora
    '5.- Actualiza mas de 1 asiento
    
    '6.- Integra 1 factura
    '7.- Elimina factura integrada . DesINTEGRA   . C L I E N T E S
    '8.- Integra 1 factura PROVEEDORES
    '9.- Elimina factura integrada . Desintegra.    P R O V E E D O R E S
    
    '10 .- Integracion masiva facturas clientes
    '11 .- Integracion masiva facturas Proveedores
    
    
    '12 .- Recalcular saldos desde hlinapu

    '13 .- IMPRIMIR asientos errores
    
    
    '---------------- DE TESORERIA
    '20
    
Public NumAsiento As Long
Public FechaAsiento As Date
Public NumDiari As Integer
Public NUmSerie As String




Private Cuenta As String
Private ImporteD As Currency
Private ImporteH As Currency
Private CCost As String
'Y estas son privadas
Private Mes As Integer
Private Anyo As Integer
Dim Fecha As String  'TENDRA la fecha ya formateada en yyy-mm-dd
Dim PrimeraVez As Boolean
Dim Sql As String
Dim Rs As Recordset

Dim INC As Long

Dim NE As Integer
Dim ErroresAbiertos As Boolean
Dim NumErrores As Long

Dim ItmX As ListItem  'Para mostra errores masivos

Private Sub AñadeError(ByRef Mensaje As String)
On Error Resume Next
'Escribimos en el fichero
If Not ErroresAbiertos Then
    NE = FreeFile
    ErroresAbiertos = True
    Open App.Path & "\ErrActua.txt" For Output As NE
    If Err.Number <> 0 Then
        MsgBox " Error abriendo fichero errores", vbExclamation
        Err.Clear
    End If
End If
Print #NE, Mensaje
If Err.Number <> 0 Then
    Err.Clear
    NumErrores = -20000
Else
    NumErrores = NumErrores + 1
End If
End Sub



Private Function CadenaImporte(VaAlDebe As Boolean, ByRef Importe As Currency, ElImporteEsCero As Boolean) As String
Dim CadImporte As String

'Si va al debe, pero el importe es negativo entonces va al haber a no ser que la contabilidad admita importes negativos
    If Importe < 0 Then
        If Not vParam.abononeg Then
            VaAlDebe = Not VaAlDebe
            Importe = Abs(Importe)
        End If
    End If
    ElImporteEsCero = (Importe = 0)
    CadImporte = TransformaComasPuntos(CStr(Importe))
    If VaAlDebe Then
        CadenaImporte = CadImporte & ",NULL"
    Else
        CadenaImporte = "NULL," & CadImporte
    End If
End Function

Private Sub CargaProgres(Valor As Integer)
Me.ProgressBar1.Max = Valor
Me.ProgressBar1.Value = 0
End Sub





Private Sub IncrementaProgres(Veces As Integer)
On Error Resume Next
Me.ProgressBar1.Value = Me.ProgressBar1.Value + (Veces * INC)
If Err.Number <> 0 Then Err.Clear
Me.Refresh
End Sub




Private Sub CargaListAsiento()

NE = FreeFile
If Dir(App.Path & "\ErrActua.txt") = "" Then
    'MsgBox "Los errores han sido eliminados. Imposible ver errores. Modulo: CargaLisAsiento"
    Exit Sub
End If

Me.frameResultados.Visible = True
'Los encabezados
ListView1.ColumnHeaders.Clear
ListView1.ColumnHeaders.Add , , "Diario", 800
ListView1.ColumnHeaders.Add , , "Fecha", 1000
ListView1.ColumnHeaders.Add , , "Nº Asie.", 1000
ListView1.ColumnHeaders.Add , , "Error", 3000


Open App.Path & "\ErrActua.txt" For Input As #NE
While Not EOF(NE)
    Line Input #NE, Cuenta
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = RecuperaValor(Cuenta, 1)
    ItmX.SubItems(1) = RecuperaValor(Cuenta, 2)
    ItmX.SubItems(2) = RecuperaValor(Cuenta, 3)
    ItmX.SubItems(3) = RecuperaValor(Cuenta, 4)
Wend
Close #NE
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
Dim bol As Boolean
If PrimeraVez Then
    PrimeraVez = False
    Me.Refresh
    bol = False
    
    'TEnemos que eliminar el archivo de errores
    If OpcionActualizar = 20 Then
        EliminarArchivoErrores
        
    End If
    Select Case OpcionActualizar
    Case 2, 3
        
        bol = True
    Case 4, 13

        NUmSerie = ""
    Case 6, 8

        bol = True
    Case 7, 9

        bol = True
    Case 10, 11
    
    Case 20
        'COBROS, pagos
        lblAsiento.Caption = "Actualizando registros"
        lblAsiento.Refresh
        If ObtenerRegistrosParaActualizar Then bol = True
        
      
    End Select
    If bol Then Unload Me
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_DblClick()
    CargaListAsiento
End Sub

Private Sub Form_Load()
Dim B As Boolean
    ErroresAbiertos = False
    Limpiar Me
 
    PrimeraVez = True
    Me.frameResultados.Visible = False
    NumErrores = 0
    ListView1.ListItems.Clear
    Select Case OpcionActualizar
    Case 1, 2, 3, 20     'Pagos, cobros tambien
        Label1.Caption = "Nº Asiento"
        Me.lblAsiento.Caption = NumAsiento
        INC = 10  'Incremento para el proggress
        If OpcionActualizar = 1 Then
            Label9.Caption = "Actualizar"
        Else
            Label9.Caption = "Integracion tesoreria"
        End If
        'Tamaño
        Me.Height = 3200
        B = True
    Case 4, 5, 13
        Me.Height = 4665

        If OpcionActualizar <> 5 Then
 

        Else
            'La opcion 5: Actualizar

        End If
        B = False
    Case 6, 7, 8, 9
        '// Estamos en Facturas
        Label1.Caption = "Nº factura"
        If OpcionActualizar < 8 Then
            Label1.Caption = Label1.Caption & " Cliente"
        Else
            Label1.Caption = Label1.Caption & " Proveedor"
        End If
        Me.lblAsiento.Caption = NUmSerie & NumAsiento
        INC = 10  'Incremento para el proggress
        If OpcionActualizar = 6 Or OpcionActualizar = 8 Then
            Label9.Caption = "Integrar Factura"
        Else
            Label9.Caption = "Eliminar Factura"
        End If
        Me.Caption = "Actualizar facturas"
        'Tamaño
        Me.Height = 3315
        B = True
    Case 10, 11

    Case 12

    End Select
    Me.frame1Asiento.Visible = B
    Me.Animation1.Visible = B
End Sub







'-------------------------------------------------------
'-------------------------------------------------------
'ANALITICA
'-------------------------------------------------------
'-------------------------------------------------------


Private Sub Form_Unload(Cancel As Integer)
If NumErrores > 0 Then CerrarFichero
End Sub

Private Sub CerrarFichero()
On Error Resume Next
If NE = 0 Then Exit Sub
Close #NE
If Err.Number <> 0 Then Err.Clear
End Sub


'Esta funcion me servira para actualizar los asientos k
' se generaran desde TESORERIA.
'YA los hemos metido en tmoactualziar
Private Function ObtenerRegistrosParaActualizar() As Boolean
Dim cad As String
    Label1.Caption = "Prepara proceso."
    Label1.Refresh
    ObtenerRegistrosParaActualizar = False
    'Borramos temporal
    
    Set Rs = New ADODB.Recordset
    Rs.Open "Select count(*) from tmpActualizar WHERE codusu =" & vUsu.Codigo, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        'NINGUN REGISTTRO A ACTUALIZAR
        NumAsiento = 0
    Else
        NumAsiento = Rs.Fields(0)
    End If
    Rs.Close
    If NumAsiento = 0 Then
        MsgBox "Ningún asiento para actualizar desde tesoreria.", vbExclamation
        Exit Function
    End If
    
    'Cargamos valores
    If NumAsiento < 32000 Then
        CargaProgres CInt(NumAsiento)
        INC = 1
    End If
    
    'Ponemos en marcha la peli
    If NumAsiento > 20 Then PonerAVI 1
    
    
    
    'Ponemos el form como toca
    Label1.Caption = "Obtener registros actualización."
    lblAsiento.Caption = ""
    Me.Height = 3315
    Me.frame1Asiento.Visible = True
    Me.Refresh
    Me.Height = 3315
    Me.Refresh
    
    Rs.Open "Select * from tmpactualizar  WHERE codusu =" & vUsu.Codigo, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        IncrementaProgres 1
        'Para poder acceder a ellos desde cualquier sitio
        NumAsiento = Rs!NumAsien
        Fecha = Format(Rs!FechaEnt, FormatoFecha)
        NumDiari = Rs!NumDiari
        'No esta bloqueado
        'Comprobamos que esta cuadrado
        cad = RegistroCuadrado
        If cad <> "" Then
            InsertaError cad
            'Borramos de tmpactualizar
            cad = "delete from tmpactualizar where codusu =" & vUsu.Codigo
            cad = cad & " AND numdiari =" & Rs!NumDiari & " AND numasien =" & Rs!NumAsien
            cad = cad & " AND fechaent ='" & Format(Rs!FechaEnt, FormatoFecha) & "'"
            Conn.Execute cad
        End If
        

        'Siguiente
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    'ACtualizarRegistros
    ActualizaASientosDesdeTMP

    'Ahora si todo ha ido bien mostraremos datos de las actualizaciones
    Me.Height = 4965
    frame1Asiento.Visible = False
    
    Me.frameResultados.Visible = True
    Me.Refresh
    Screen.MousePointer = vbHourglass
    If NumErrores > 0 Then
        Close #NE
        Label7.Caption = "Se han producido errores."
        CargaListAsiento
    Else
       Label7.Caption = "NO se han producido errores."
       Me.Refresh
       ObtenerRegistrosParaActualizar = True
    End If
    
End Function


Private Sub PonerAVI(NumAVI As Integer)
On Error GoTo EPonerAVI
    If NumAVI = 1 Then
        Me.Animation1.Open App.Path & "\actua.avi"
        Me.Animation1.Play
        Me.Animation1.Visible = True
    Else
    
    End If
Exit Sub
EPonerAVI:
    MuestraError Err.Number, "Poner Video"
End Sub


Private Function RegistroCuadrado() As String
    Dim Deb As Currency
    Dim hab As Currency
    Dim RSUM As ADODB.Recordset

    'Trabajamos con RS que es global
    RegistroCuadrado = "" 'Todo bien
    
    Set RSUM = New ADODB.Recordset
    Sql = "SELECT Sum(hlinapu.timporteD) AS SumaDetimporteD, Sum(hlinapu.timporteH) AS SumaDetimporteH"
    Sql = Sql & " ,hlinapu.numdiari,hlinapu.fechaent,hlinapu.numasien"
    Sql = Sql & " From hlinapu GROUP BY hlinapu.numdiari, hlinapu.fechaent, hlinapu.numasien "
    Sql = Sql & " HAVING (((hlinapu.numdiari)=" & NumDiari
    Sql = Sql & ") AND ((hlinapu.fechaent)='" & Fecha
    Sql = Sql & "') AND ((hlinapu.numasien)=" & NumAsiento
    Sql = Sql & "));"
    
    
    
    RSUM.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RSUM.EOF Then
        Deb = DBLet(RSUM.Fields(0), "N")
        'Deb = Round(Deb, 2)
        hab = RSUM.Fields(1)
        'Hab = Round(Hab, 2)
        CCost = ""
    Else
        Deb = 0
        hab = -1
        CCost = "Asiento sin lineas"
    End If
    
    RSUM.Close
    Set RSUM = Nothing
    If Deb <> hab Then
        If CCost = "" Then CCost = "Asiento descuadrado"
        RegistroCuadrado = CCost
    End If

End Function

Private Function InsertaError(ByRef CADENA As String)
Dim vS As String
    'Insertamos en errores
    'Esta lo tratamos con error especifico
    
    On Error Resume Next


        'Insertamos error para ASIENTOS
        vS = NumDiari & "|"
        vS = vS & Fecha & "|"
        vS = vS & NumAsiento & "|"
        vS = vS & CADENA & "|"
    

    'Modificacion del 10 de marzo
    'Conn.Execute vS
    AñadeError vS
    
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error." & vbCrLf & Err.Description & vbCrLf & vS
        Err.Clear
    End If
End Function

Private Function ActualizaASientosDesdeTMP()
Dim RT As Recordset


'Para el progress
NumAsiento = ProgressBar1.Max
Me.lblAsiento.Caption = "Nº asiento:"
If NumAsiento < 3000 Then
    CargaProgres NumAsiento * 10
    Else
    CargaProgres 32000
End If
INC = 1


Sql = "Select * from tmpactualizar where codusu=" & vUsu.Codigo
Set RT = New ADODB.Recordset
RT.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
While Not RT.EOF
    NumAsiento = RT!NumAsien
    FechaAsiento = RT!FechaEnt
    NumDiari = RT!NumDiari
    'Actualiza el asiento

    'Siguiente
    RT.MoveNext
Wend
RT.Close
Set RT = Nothing
End Function


Private Sub BorrarArchivoTemporal()
On Error Resume Next
If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
If Err.Number <> 0 Then MuestraError Err.Number, "Borrar fichero temporal"
End Sub



Private Sub EliminarArchivoErrores()
On Error Resume Next
If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
If Err.Number <> 0 Then MuestraError Err.Number, "Borrar fichero temporal"
End Sub

