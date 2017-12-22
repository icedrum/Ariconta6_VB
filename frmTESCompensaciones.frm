VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESCompensaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compensaciones"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   13785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDatos 
      Height          =   7215
      Left            =   60
      TabIndex        =   6
      Top             =   30
      Width           =   13515
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "frmTESCompensaciones.frx":0000
         Top             =   480
         Width           =   7275
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1680
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   480
         Width           =   4275
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
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
         Left            =   12030
         TabIndex        =   5
         Top             =   6600
         Width           =   1215
      End
      Begin VB.TextBox Text3 
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
         Height          =   360
         Index           =   0
         Left            =   4530
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   6600
         Width           =   1365
      End
      Begin VB.TextBox Text3 
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
         Height          =   360
         Index           =   1
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text3"
         Top             =   6600
         Width           =   1365
      End
      Begin VB.TextBox Text3 
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
         Height          =   360
         Index           =   2
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   6600
         Width           =   1365
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
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
         Left            =   10710
         TabIndex        =   4
         Top             =   6600
         Width           =   1215
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   5175
         Index           =   1
         Left            =   6030
         TabIndex        =   3
         Top             =   1200
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuenta"
            Object.Width           =   4409
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2434
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Vto"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Importe"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "YaEfectuado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "numserie"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   5175
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Serie"
            Object.Width           =   1410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2434
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Vto"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Importe"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "YaEfectuado"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   13080
         TabIndex        =   14
         Top             =   120
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ayuda"
            EndProperty
         EndProperty
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   7290
         Picture         =   "frmTESCompensaciones.frx":0006
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Importes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   2100
         TabIndex        =   13
         Top             =   6600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   11
         Top             =   210
         Width           =   1095
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmTESCompensaciones.frx":0A08
         Top             =   480
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmTESCompensaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1

Private Const IdPrograma = 606



Dim SQL As String   'Cadena de uso comun
Dim Im As Currency
Dim CampoAnterior As String
Dim CadNif As String

Dim vCP As Ctipoformapago

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
Dim IT As ListItem
Dim AumentaElImporteDelVto As Boolean
Dim IndiceListView As Integer
Dim ModificarVto As Boolean  'No pone el impcobrado, pone vto el total que queda de comensar


    Dim LCob As Collection
    Dim LPag As Collection
    
    'COmprobaciones
    'Que hay seleccionado algun vencimiento
    SQL = ""
    For NumRegElim = 1 To lw1(0).ListItems.Count
        If lw1(0).ListItems(NumRegElim).Checked Then
            SQL = "1"
            Exit For
        End If
    Next
    If SQL <> "" Then
        For NumRegElim = 1 To lw1(1).ListItems.Count
            If lw1(1).ListItems(NumRegElim).Checked Then
                SQL = "1"
                'Nos salimos.
                Exit For
            End If
        Next
    End If
    If SQL = "" Then
        MsgBox "Debe seleccionar algun vencimiento(cobros y pagos)", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Vamos a dar la opcion de que el total, en vez de ser contra el banco, sea contra un vto
    'Es decir. En ese vto lo disminuire y de esa forma NO hago el apunte a la cta del banco
    'AHora vere sobre que recibo puedo hacer el cargo, para ver si voy a meter
    Im = CCur(Text3(2).Tag)
    AumentaElImporteDelVto = False
    If Im = 0 Then
        'NADA
        'Dejamos que it siga a NOTHNG
    Else
        If Im > 0 Then
            'Estoy pagando mas que cobrando
            SQL = CStr(EstableceVtoQueTotaliza(0))
            If SQL <> "0" Then Set IT = lw1(0).ListItems(CInt(SQL))
        Else
            'Estoy COBRANDO mas que pagando
            SQL = CStr(EstableceVtoQueTotaliza(1))
            If SQL <> "0" Then Set IT = lw1(1).ListItems(CInt(SQL))
        End If
        
        
        'Marzo 2009
        'Si incrementa un vto pq el importe es mayor del que habia.
        If IT Is Nothing Then
        
            'NO dejamos que el impte de un vto aumente.
            MsgBox "El importe a compensar no se puede realizar sobre un único vencimiento", vbExclamation
            If False Then
                '
                'AQUI , de momento, NO entra
                AumentaElImporteDelVto = True
                'No hay ningun vto donde compensar.
                'Seleccionare el ultimo seleccionado del listview que corresponda
                If CCur(Text3(2).Tag) > 0 Then
                    SQL = CStr(ForzarVtoQueTotaliza(0))
                    Set IT = lw1(0).ListItems(CInt(SQL))
                Else
                    'Estoy COBRANDO mas que pagando
                    SQL = CStr(ForzarVtoQueTotaliza(1))
                    Set IT = lw1(1).ListItems(CInt(SQL))
                End If
            End If
        End If
        
    End If
    
    Set vCP = New Ctipoformapago
    
    'Preparamos los campos del siguiente campo
    If vCP.Leer(vbEfectivo) = 0 Then
    
        Dim CDC As Integer, CDP As Integer  'deb cli  y debe pro
        CDC = vCP.condecli
        CDP = vCP.condepro
        
        ValoresConceptosPorDefecto True, CDC, CDP
        vCP.conhacli = CDC
        vCP.condepro = CDP
        SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(CDC))
        CadenaDesdeOtroForm = vCP.conhacli & "|" & SQL & "|"
        SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(CDP))
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & vCP.condepro & "|" & SQL & "|"
    Else
        CadenaDesdeOtroForm = "||||"
    End If
    
    'Le indico si puede realizar la compensacion sobre un vto, o no
    If IT Is Nothing Then
        '0:No
        SQL = "0|Nada|"
    Else
        '1: Si
        SQL = "1|" & IT.Index & "|"
    End If
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & SQL
    Set vCP = Nothing
    SQL = ""
    
    frmTESListado.Opcion = 22
    
    'Si puede compensar sobre algun vto en especial
    If Not IT Is Nothing Then
        
        If CCur(Text3(2).Tag) > 0 Then
            IndiceListView = 0
        Else
            IndiceListView = 1
        End If
        
        For NumRegElim = 1 To Me.lw1(IndiceListView).ListItems.Count
            If Me.lw1(IndiceListView).ListItems(NumRegElim).Checked Then
                Im = ImporteFormateado(lw1(IndiceListView).ListItems(NumRegElim).SubItems(4))
                If Im > Abs(CCur(Text3(2).Tag)) Then
                    If IndiceListView = 0 Then
                        SQL = lw1(IndiceListView).ListItems(NumRegElim).Text
                    Else
                        'pagos
                        SQL = ""
                    End If
                    
                    SQL = "Fact: " & SQL & lw1(IndiceListView).ListItems(NumRegElim).SubItems(1) & " ,vto " & lw1(IndiceListView).ListItems(NumRegElim).SubItems(3) & _
                            " de fecha " & lw1(IndiceListView).ListItems(NumRegElim).SubItems(2)
                    
                    frmTESListado.InsertaItemComboCompensaVto SQL, CInt(NumRegElim)
                End If
            End If
        Next
    End If
    
    frmTESListado.Show vbModal
    
    If CadenaDesdeOtroForm <> "" Then
        
       'Compruebo que ninguna de las dos cuentas esta bloqueda para le fecha de contabilizacion
        If CuentaBloqeada(Text1(0).Text, RecuperaValor(CadenaDesdeOtroForm, 4), True) Then Exit Sub
        'Compruebo que ninguna de las dos esta bloqueda para le fecha de contabilizacion
         
        SQL = Text4.Tag
        While SQL <> ""
             NumRegElim = InStr(1, SQL, "|")
             If NumRegElim = 0 Then
                 SQL = ""
             Else
                 If CuentaBloqeada(Mid(SQL, 1, NumRegElim - 1), RecuperaValor(CadenaDesdeOtroForm, 4), True) Then Exit Sub
                 SQL = Mid(SQL, NumRegElim + 1)
             End If
        Wend
                    
                    
        ModificarVto = RecuperaValor(CadenaDesdeOtroForm, 9) = "1"
        'Le quito el ultmo pipe para dejarlo como estaba
        CadenaDesdeOtroForm = Left(CadenaDesdeOtroForm, Len(CadenaDesdeOtroForm) - 2)     'quito el pipe  y el value
        
        'Añado las obsrvaciones
        'Le quitomel ultmo pipe
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 1, Len(CadenaDesdeOtroForm) - 1)
        'Comprueno si lleva contra un vto o NO
        NumRegElim = InStrRev(CadenaDesdeOtroForm, "|")
        SQL = Mid(CadenaDesdeOtroForm, NumRegElim + 1)
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 1, NumRegElim - 1)
        If SQL = "0" Then
            'NO ha seleccionado el vto, con lo cual pongo el IT a nothing
            Set IT = Nothing
            
        Else
            'Va a compensar contra un vto. Si el vto va a aumentar entonces le pregunto si desea continuar
          
            If IT.Index <> Val(SQL) Then
                'Ha cambiado el VTO que le ofertabamos nosotros
                Set IT = lw1(IndiceListView).ListItems(CInt(Val(SQL)))
            End If
            'Aqui NO debe de ebtrar
            If AumentaElImporteDelVto Then
                SQL = "El importe del vencimiento Factura: "
                SQL = SQL & IT.SubItems(1) & "   nº" & IT.SubItems(3) & "  de fecha " & IT.SubItems(2)
                SQL = SQL & " se va a incrementar"
                
                SQL = SQL & vbCrLf & "¿Desea continuar?"
                If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            End If
        End If
        
        'ASigno la nueva forma de pago del vto resultante (o en su defecto obvio el dato
        'Con lo cual voy a quitar el utlimi pipe que es la FP
        NumRegElim = InStrRev(CadenaDesdeOtroForm, "|")
        SQL = Mid(CadenaDesdeOtroForm, NumRegElim + 1)
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 1, NumRegElim)
        IndiceListView = -1
        If Not IT Is Nothing Then
            If SQL <> "" Then
                If IsNumeric(SQL) Then IndiceListView = Val(SQL)
            End If
        End If
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text1(0).Text & " " & Text2(0).Text & " - " & Text4.Text & "|"
    
    
        Screen.MousePointer = vbHourglass
        Set vCP = New Ctipoformapago
        If vCP.Leer(vbEfectivo) = 0 Then
            vCP.condecli = CInt(RecuperaValor(CadenaDesdeOtroForm, 1))
            vCP.condepro = CInt(RecuperaValor(CadenaDesdeOtroForm, 2))
            vCP.conhacli = vCP.condecli
            vCP.conhapro = vCP.condepro
            'Los guardo
            CDC = vCP.condecli
            CDP = vCP.condepro
            ValoresConceptosPorDefecto False, CDC, CDP
           'Las ampliaciones
           vCP.ampdecli = 0
           vCP.ampdepro = 0
           vCP.amphacli = 0
           vCP.amphapro = 0

                                                                            'IndiceListView: Si compensa cn vto y quiere cambiar la forma de pago
            If CrearColecciones(LCob, LPag, vCP, IT, AumentaElImporteDelVto, IndiceListView, ModificarVto) Then ContabilizarCompensaciones LCob, LPag, CadenaDesdeOtroForm, AumentaElImporteDelVto
           ' If CadenaDesdeOtroForm <> "" Then
                
                'LOG
                Dim SqlLog As String
                
                SqlLog = "Cliente      : " & Text1(0) & " " & Text2(0)
                SqlLog = SqlLog & vbCrLf & "Proveedores  : " & Text4
                SqlLog = SqlLog & vbCrLf & "Fras Cliente : "
                For i = 1 To Me.lw1(0).ListItems.Count
                    If lw1(0).ListItems(i).Checked Then
                        SqlLog = SqlLog & vbCrLf & lw1(0).ListItems(i).Text & " " & lw1(0).ListItems(i).SubItems(1) & " " & lw1(0).ListItems(i).SubItems(2) & " " & lw1(0).ListItems(i).SubItems(3) & " " & lw1(0).ListItems(i).SubItems(4) & " "
                    End If
                Next i
                
                SqlLog = SqlLog & vbCrLf & "Fras Proveedor : "
                For i = 1 To Me.lw1(1).ListItems.Count
                    If lw1(1).ListItems(i).Checked Then
                        SqlLog = SqlLog & vbCrLf & lw1(1).ListItems(i).Text & " " & lw1(1).ListItems(i).SubItems(6) & " " & lw1(1).ListItems(i).SubItems(1) & " " & lw1(1).ListItems(i).SubItems(2) & " " & lw1(1).ListItems(i).SubItems(3) & " " & lw1(1).ListItems(i).SubItems(4) & " "
                    End If
                Next i
                vLog.Insertar 26, vUsu, SqlLog
                
                
                CargarListView 0
                CargarListView 1
           ' End If
           

        End If
        
        
        
        
        
    End If
    Screen.MousePointer = vbDefault
    Set LCob = Nothing
    Set LPag = Nothing
    Set vCP = Nothing
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    Limpiar Me
    Text3(0).Tag = 0:    Text3(1).Tag = 0:    Text3(2).Tag = 0
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub imgCuentas_Click(Index As Integer)
    
    
    If Index = 0 Then
        'Avisar si ya han cargado datos
         Screen.MousePointer = vbHourglass
         Set frmCCtas = New frmColCtas
         SQL = ""
         CampoAnterior = Text1(Index).Text
         frmCCtas.DatosADevolverBusqueda = "0"
         frmCCtas.Show vbModal
         Set frmCCtas = Nothing
         If SQL <> "" Then
            Text1(Index).Text = RecuperaValor(SQL, 1)
            Text2(Index).Text = RecuperaValor(SQL, 2)
            If CampoAnterior <> Text1(Index).Text Then CargarListView Index
        End If
    Else
        'PROVEEDORES
        frmTESVarios.Opcion = 21
        CadenaDesdeOtroForm = Text4.Tag
        frmTESVarios.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            'Si ha cambiado algo
            If Text4.Tag <> CadenaDesdeOtroForm Then
                Text4.Tag = CadenaDesdeOtroForm
                CargarListView Index
            End If
                
                
        End If
    End If
    
End Sub




Private Sub lw1_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    
    Im = ImporteFormateado(Item.SubItems(4))
    If Not Item.Checked Then Im = -Im
    
    'Arrastro
    Text3(Index).Tag = CCur(Text3(Index).Tag) + Im
    CalculaImportes
    
    
    Set lw1(Index).SelectedItem = Item
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
    CampoAnterior = Text1(Index).Text
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYCuentas KeyAscii, 0
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYCuentas(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgCuentas_Click (Indice)
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim C As String
Dim NIF As String

        NIF = ""

        
        If Text1(Index).Text = "" Then
             Text2(Index).Text = ""

        Else
            C = Text1(Index).Text
            If Not CuentaCorrectaUltimoNivel(C, SQL) Then
                MsgBox SQL & " - " & C, vbExclamation
                SQL = ""
                C = ""
            End If
            Text1(Index).Text = C
            Text2(Index).Text = SQL
            If C = "" Then
                PonFoco Text1(Index)
            Else
                ' Añadida esta parte donde según el nif del cliente me voy a buscar las ctas de proveedor con el mismo nif
                
                
                
                    NIF = DevuelveDesdeBD("nifdatos", "cuentas", "codmacta", C, "T")
                    CadNif = ""
                    If NIF <> "" Then
                        C = ""
                        If Text4.Tag <> "" Then C = "N"
                        If C <> "" Then
                            If MsgBox("Leer datos proveedores " & NIF & "?", vbQuestion + vbYesNo) = vbYes Then C = ""
                        End If
                        If C = "" Then
                            CadNif = NIF
        
                            Text4.Tag = CuentasProveedorDelNif(NIF)
                            If Text4.Tag <> "" Then
                                CargarListView 1
                            End If
                        End If
                    End If
                
            End If
        End If
        'Cargamos el listview
        If CampoAnterior <> Text1(Index).Text Then CargarListView Index
End Sub


Private Function CuentasProveedorDelNif(NIF As String) As String
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim CadResult As String

    CuentasProveedorDelNif = ""

    SQL = "select distinct codmacta from pagos where nifprove = " & DBSet(NIF, "T")
    SQL = SQL & " and impefect - coalesce(imppagad,0) <> 0 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadResult = ""
    
    While Not Rs.EOF
        CadResult = CadResult & Rs!codmacta & "|"
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    CuentasProveedorDelNif = CadResult

End Function

Private Sub CalculaImportes()
    Text3(2).Tag = CCur(Text3(0).Tag) - CCur(Text3(1).Tag)
    Text3(0).Text = Format(Text3(0).Tag, FormatoImporte)
    Text3(1).Text = Format(Text3(1).Tag, FormatoImporte)
    Text3(2).Text = Format(Text3(2).Tag, FormatoImporte)
End Sub


Private Sub CargarListView(Indice As Integer)
Dim C As String
    Screen.MousePointer = vbHourglass
    
        lw1(Indice).ListItems.Clear
        Text3(Indice).Text = ""
        Text3(Indice).Tag = 0
        CalculaImportes
        
        If Indice = 0 Then
            If Text1(Indice).Text = "" Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    
    Set miRsAux = New ADODB.Recordset
    If Indice = 0 Then
        'CLIENTE
        CargaDatosListview Indice
    Else
        'PROVEEEDORES
        'Borramos datos anteriores
        Text4.Text = ""
        lw1(Indice).ListItems.Clear
        Text3(Indice).Text = ""
        Text3(Indice).Tag = 0
        CalculaImportes
        'Cargamos
        C = Text4.Tag
        While C <> ""
            NumRegElim = InStr(1, C, "|")
            If NumRegElim = 0 Then
                C = ""
            Else
                SQL = Mid(C, 1, NumRegElim - 1)
                C = Mid(C, NumRegElim + 1)
                CargaDatosListview Indice  'Cargamos para este cliente
            End If
        Wend
    End If
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub CargaDatosListview(Indice As Integer)
Dim IT As ListItem
Dim YaEfectuado As Currency  'Lo que ya se ha cobrado/pagado
Dim CargaEnListview As Boolean

    On Error GoTo ECargaDatosListview
    


    
    
    If Indice = 0 Then
        SQL = "select numserie,numfactu,fecfactu,numorden,impvenci,impcobro,gastos,codmacta from cobros where"
        SQL = SQL & " codrem is null and anyorem is null and coalesce(transfer,0)=0 "
        'Y que el talon pagare NO este recepcionado
        SQL = SQL & " AND recedocu = 0"
        SQL = SQL & " and  codmacta ='" & Text1(Indice).Text & "'"
    Else
        'En SQL va el codmacta
        CadNif = DevuelveDesdeBD("nifdatos", "cuentas", "codmacta", SQL, "T")
        
        SQL = " and nrodocum is null and pagos.codmacta ='" & SQL & "'"
        If CadNif <> "" Then
            SQL = SQL & " and pagos.nifprove = " & DBSet(CadNif, "T")
        End If
        
        SQL = " WHERE impefect - coalesce(imppagad,0) <> 0 " & SQL 'AND estacaja =0
        SQL = "select numfactu,fecfactu,numorden,impefect,imppagad,pagos.codmacta as codmacta,nomprove as nommacta, numserie  FROM pagos " & SQL
    End If
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not miRsAux.EOF
        'Veremos si el importe es positivo, o no
        
        If Indice = 0 Then
            Im = miRsAux!ImpVenci - DBLet(miRsAux!impcobro, "N") + DBLet(miRsAux!Gastos, "N")
            YaEfectuado = DBLet(miRsAux!impcobro, "N")
            CargaEnListview = Im > 0
            
        Else
            Im = miRsAux!ImpEfect - DBLet(miRsAux!imppagad, "N")
            YaEfectuado = DBLet(miRsAux!imppagad, "N")
            CargaEnListview = True 'Pase lo que pase
        End If
        'If Im > 0 Then
        If CargaEnListview Then
        
    
            Set IT = lw1(Indice).ListItems.Add()
            If Indice = 0 Then
                IT.Text = miRsAux!NUmSerie
                IT.SubItems(1) = miRsAux!NumFactu
                IT.SubItems(2) = miRsAux!FecFactu
                IT.SubItems(3) = miRsAux!numorden
                'Importe:
            
            Else
                
                IT.Text = Mid(miRsAux!Nommacta, 1, 20)
                'Para que aparezca en el
                If SQL = "" Then
                    If Text4.Text <> "" Then Text4.Text = Text4.Text & vbCrLf
                    Text4.Text = Text4.Text & miRsAux!codmacta & "   " & miRsAux!Nommacta
                    SQL = "D"
                End If
                IT.SubItems(1) = miRsAux!NumFactu
                IT.SubItems(2) = miRsAux!FecFactu
                IT.SubItems(3) = miRsAux!numorden
                
            End If
            IT.SubItems(4) = Format(Im, FormatoImporte)
            IT.SubItems(5) = YaEfectuado
            IT.Tag = miRsAux!codmacta
            
            If Indice = 1 Then IT.SubItems(6) = miRsAux!NUmSerie
            
            
           
    
        End If
         miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    Exit Sub
ECargaDatosListview:
    MuestraError Err.Number, Err.Description
End Sub


'Modif. Enero 2009, casi febrero
'Compensacion UNO a varios. Puede elegir el vto sobre el cual va a Imputarse la comepnsacion
'CambiaFormaPago : Si es <=0 Nada. Si >0 entonces, en el UPDATE ponemos esa forpa
Private Function CrearColecciones(ByRef CCli As Collection, ByRef CPro As Collection, ByRef FP As Ctipoformapago, ByRef ItmVto As ListItem, Va_a_AumentaElImporteDelVto As Boolean, CambiaFormaPago As Integer, CambiaIMporteVto As Boolean) As Boolean
Dim Ampliacion As String
Dim VaAlDebe As Boolean
Dim Total As Currency
Dim YaCobrado As Currency
Dim ContrapartidaPago As String   'Cual es, si la del proveedor o NULL

Dim CadenaUpdate As String
Dim CompensaSobreCobros As Byte
Dim FrasCli As String
Dim FrasPro As String

    '0: NO compensa
    '1: Cobros
    '2: Pagos

    On Error GoTo ECrearColecciones
    CrearColecciones = False

    Total = 0
    Set CCli = New Collection
    Set CPro = New Collection
            
            
    If Text4.Tag = "" Then
        ContrapartidaPago = Trim(RecuperaValor(CadenaDesdeOtroForm, 5)) 'Banco
        If ContrapartidaPago = "" Then ContrapartidaPago = "NULL"
    Else
        'cta1
        CompensaSobreCobros = InStr(1, Text4.Tag, "|")
        'Vemos si tiene mas de uno
        CompensaSobreCobros = InStr(CompensaSobreCobros + 1, Text4.Tag, "|")
        If CompensaSobreCobros = 0 Then
            'Solo1
            ContrapartidaPago = "'" & RecuperaValor(Text4.Tag, 1) & "'"
        Else
            ContrapartidaPago = "NULL"
        End If
    End If
    
    FrasCli = ""
    FrasPro = ""
    For NumRegElim = 1 To lw1(0).ListItems.Count
        If lw1(0).ListItems(NumRegElim).Checked Then
            With lw1(0).ListItems(NumRegElim)
                CampoAnterior = .Text & Format(.SubItems(1), "00000")
                FrasCli = FrasCli & "," & CampoAnterior
            End With
        End If
    Next NumRegElim
    FrasCli = Mid(FrasCli, 2)
    
    FrasPro = ""
    For NumRegElim = 1 To lw1(1).ListItems.Count
        If lw1(1).ListItems(NumRegElim).Checked Then
            With lw1(1).ListItems(NumRegElim)
                CampoAnterior = .SubItems(1)
                FrasPro = FrasPro & "," & CampoAnterior
            End With
        End If
    Next NumRegElim
    FrasPro = Mid(FrasPro, 2)
    
    
    
    CampoAnterior = ""
    CompensaSobreCobros = 0
    If Not ItmVto Is Nothing Then
        'Puede compensar contra un vencimiento. Pero SI NO quiere no habra marcado el check
        
        If CCur(Text3(2).Tag) >= 0 Then
            CompensaSobreCobros = 1
        Else
            CompensaSobreCobros = 2
        End If
    End If
    
        'Montaremos esta linea que sera la que hagamos INSERT
        'codconce numdocum, ampconce , codmacta, timporteD,timporteH, ctacontr) "
        
        
        'Para los cobros
        


        
        'Descripcion concepto
        CampoAnterior = DevuelveDesdeBD("nomconce", "conceptos", "codconce", FP.conhacli)
        SQL = ""
        
        '----------------------------------------------------------------------
        '----------------------------------------------------------------------
        'CLIENTES
        For NumRegElim = 1 To lw1(0).ListItems.Count
            
            If lw1(0).ListItems(NumRegElim).Checked Then
                With lw1(0).ListItems(NumRegElim)
                    
                    'Monto el tocito para el sql
                    Ampliacion = CampoAnterior & " "

                    If FrasCli = "" Then

                            If FP.ampdecli = 3 Then
                                'NUEVA forma de ampliacion
                                'No hacemos nada pq amp11 ya lleva lo solicitado
                                
                            Else
                                If FP.ampdecli = 4 Then
                                    'COntrapartida
                                    Ampliacion = Ampliacion & RecuperaValor(CadenaDesdeOtroForm, 5)
                                               
                                Else
                                    If FP.ampdecli = 2 Then
                                       Ampliacion = Ampliacion & Format(.SubItems(2), "dd/mm/yyyy")
                                    Else
                                       If FP.ampdecli = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
                                       'Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!codfaccl
                                       Ampliacion = Ampliacion & .Text & "/" & .SubItems(1)
                                       
                                    End If
                                End If
                            End If
                    Else
                        Ampliacion = Ampliacion & FrasCli
                    End If
                    
                    Im = ImporteFormateado(.SubItems(4))
                    CadenaUpdate = ""
                    'Si compensa sobre un vto de cobro
                    If CompensaSobreCobros = 1 Then
                        'Hace la comep
                        If lw1(0).ListItems(NumRegElim).Index = ItmVto.Index Then
                        
                            'Nuevo Marzo 2009
                            If Va_a_AumentaElImporteDelVto Then
                                'Es decir, habia un importe y va a haber otro (mayor)
                                'Con lo cual. Gastos CER=, ultco CERO Y pondre num vto 99
                                'Impvenci el nuevo importe. Y fecha venci la fecha de contablizacion
                                
                                Im = CCur(Text3(2).Tag)
                                CadenaUpdate = RecuperaValor(CadenaDesdeOtroForm, 4) 'Fecha contabilizacion
                                CadenaUpdate = "UPDATE cobros set impvenci= " & TransformaComasPuntos(CStr(Im)) & ",fecvenci = '" & Format(CadenaUpdate, FormatoFecha) & "'"
                                CadenaUpdate = CadenaUpdate & ",impcobro= NULL,fecultco = NULL,gastos=NULL,Referencia='Compen. " & Format(Now, "dd/mm/yyyy hh:mm") & "'"
                                CadenaUpdate = CadenaUpdate & ",numorden=99"
                                If CambiaFormaPago > 0 Then CadenaUpdate = CadenaUpdate & " , codforpa = " & CambiaFormaPago
                                CadenaUpdate = CadenaUpdate & " " 'Por si queremos añadir mas camos a updatear
                                
                                Im = ImporteFormateado(.SubItems(4))
                                
                            
                            Else
                                
                                'ES SOBRE ESTE VTO sobre el que comepenso
                                YaCobrado = CCur(.SubItems(5))
                                Im = CCur(.SubItems(4)) - CCur(Text3(2).Tag)
                                YaCobrado = YaCobrado + Im
                                
                                CadenaUpdate = "UPDATE cobros set "
                                If CambiaIMporteVto Then
                                    'Cambia el importe vto y pone a NULL el cobrado
                                    CadenaUpdate = CadenaUpdate & " impvenci=  " & TransformaComasPuntos(CStr(CCur(Text3(2).Tag))) & ",gastos =NULL, impcobro=NULL"
                                    CadenaUpdate = CadenaUpdate & " ,observa=trim(concat(if(observa is null, """",observa),""    "",""Compen. " & Format(Now, "dd/mm/yyyy") & " Vto: " & CStr(YaCobrado) & """))"
                                Else
                                    'No cambial el importe del vecnimiento, lo deja como estaba y lo pone sobre impcobro
     
                                    CadenaUpdate = CadenaUpdate & " impcobro=  " & TransformaComasPuntos(CStr(YaCobrado))
                                    
                                End If
                                
                                CadenaUpdate = CadenaUpdate & ",fecultco = '" & Format(RecuperaValor(CadenaDesdeOtroForm, 4), FormatoFecha) & "'"
                                'Si cambia la Forpa
                                If CambiaFormaPago > 0 Then CadenaUpdate = CadenaUpdate & " , codforpa = " & CambiaFormaPago
                                'Por cuanto ira el apunte
                                Im = CCur(Text3(2).Tag)
                                Im = ImporteFormateado(.SubItems(4)) - Im
                            End If
                        End If
                    End If
                    
                    Total = Total + Im
                    VaAlDebe = False
                    
                    If Im < 0 Then
                        If Not vParam.abononeg Then
                               VaAlDebe = True
                               Im = -Im
                        End If
                    End If
                    'codconce numdocum, ampconce , codmacta, timporteD,timporteH, ctacontr
                    SQL = FP.condecli & ",'" & .Text & Format(.SubItems(1), "000000") & "','"
        
                    SQL = SQL & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "','" & Text1(0).Text & "',"

                    'Importe
                    If VaAlDebe Then
                        SQL = SQL & TransformaComasPuntos(CStr(Im)) & ",NULL"
                    Else
                        SQL = SQL & "NULL," & TransformaComasPuntos(CStr(Im))
                    End If
                    
                    'Contrapartida. esta guaddad en ContrapartidaPago
                    SQL = SQL & "," & ContrapartidaPago & ","
                    
                    'Habran dos pipes.
                    '   1.- lo que tengo que insertar en hlinapu
                    '   2.- El select prparado para eliminar el cobro / pago
                    '       Si compensa, habra una C al principio
                    '   3.- Para buscar la factura
                    Ampliacion = "|" & CadenaUpdate & " WHERE `numserie`='" & .Text & "' and numfactu=" & .SubItems(1)
                    Ampliacion = Ampliacion & " and `fecfactu`='" & Format(.SubItems(2), FormatoFecha) & "' and `numorden`=" & .SubItems(3) & "|"
                    '
                    Ampliacion = Ampliacion & " WHERE `numserie`='" & .Text & "' and numfactu=" & .SubItems(1)
                    Ampliacion = Ampliacion & " and `fecfactu`='" & Format(.SubItems(2), FormatoFecha) & "' and `numorden`=" & .SubItems(3) & "|"
            
                    CCli.Add SQL & Ampliacion
                End With
            End If
    Next NumRegElim
    
    '----------------------------------------------------------------------
    '----------------------------------------------------------------------
    'PROVEEDORES
    CampoAnterior = DevuelveDesdeBD("nomconce", "conceptos", "codconce", FP.condepro)
    For NumRegElim = 1 To lw1(1).ListItems.Count
        
           
        
        
            If lw1(1).ListItems(NumRegElim).Checked Then
                With lw1(1).ListItems(NumRegElim)
                    
                    'Monto el tocito para el sql
                    Ampliacion = CampoAnterior & " "
                    
                    
                    If FrasCli = "" Then
                            Select Case FP.amphapro
                            Case 0, 1
                               If FP.amphapro = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
                               Ampliacion = Ampliacion & .SubItems(1)
                            
                            Case 2
                               'Fecha vto
                               Ampliacion = Ampliacion & .SubItems(1)
                            
                            
                            Case 4
                                'COntrapartida
                                Ampliacion = Ampliacion & RecuperaValor(CadenaDesdeOtroForm, 5)
                                
                            End Select
                    Else
                        Ampliacion = Ampliacion & FrasCli
                    End If
                    
                    
                    Im = ImporteFormateado(.SubItems(4))
                    CadenaUpdate = ""
                    'Si compensa sobre un vto de pago
                    If CompensaSobreCobros = 2 Then
                        
                        If lw1(1).ListItems(NumRegElim).Index = ItmVto.Index Then
                        
                           
                            'Nuevo Marzo 2009
                            If Va_a_AumentaElImporteDelVto Then
                                'Es decir, habia un importe y va a haber otro (mayor)
                                'Con lo cual. Gastos CER=, ultco CERO Y pondre num vto 99
                                'Impvenci el nuevo importe. Y fecha venci la fecha de contablizacion
                                
                                
                                Im = Abs(CCur(Text3(2).Tag))
                                CadenaUpdate = RecuperaValor(CadenaDesdeOtroForm, 4) 'Fecha contabilizacion
                                CadenaUpdate = "UPDATE pagos set impefect= " & TransformaComasPuntos(CStr(Im)) & ",fecefect = '" & Format(CadenaUpdate, FormatoFecha) & "'"
                                CadenaUpdate = CadenaUpdate & ",imppagad= NULL,fecultpa = NULL,Referencia='Compen. " & Format(Now, "dd/mm/yyyy hh:mm") & "'"
                                CadenaUpdate = CadenaUpdate & ",numorden=99"
                                If CambiaFormaPago > 0 Then CadenaUpdate = CadenaUpdate & " , codforpa = " & CambiaFormaPago
                                CadenaUpdate = CadenaUpdate & " " 'Por si queremos añadir mas camos a updatear
                                Im = ImporteFormateado(.SubItems(4))
                                
                        
                            Else
                                'ES SOBRE ESTE VTO sobre el que comepenso
                                'El importe estara en negativo
                                YaCobrado = CCur(.SubItems(5))
                                Im = CCur(.SubItems(4)) - Abs(CCur(Text3(2).Tag))
                                YaCobrado = YaCobrado + Im
                                
                                CadenaUpdate = "UPDATE pagos set "
                                If CambiaIMporteVto Then
                                    CadenaUpdate = CadenaUpdate & "impefect=  " & TransformaComasPuntos(CStr(Abs(CCur(Text3(2).Tag))))
                                    CadenaUpdate = CadenaUpdate & ",imppagad= NULL"
                                    
                                    
                                    CadenaUpdate = CadenaUpdate & " ,observa=trim(concat(if(observa is null, """",observa),""    "",""Compen. " & Format(Now, "dd/mm/yyyy") & " Vto: " & CStr(YaCobrado) & """))"
                                    
                                Else
                                    CadenaUpdate = CadenaUpdate & "imppagad= " & TransformaComasPuntos(CStr(YaCobrado))
                                    
                                End If
                                '
                                CadenaUpdate = CadenaUpdate & ",fecultpa = '" & Format(RecuperaValor(CadenaDesdeOtroForm, 4), FormatoFecha) & "'"
                                If CambiaFormaPago > 0 Then CadenaUpdate = CadenaUpdate & " , codforpa = " & CambiaFormaPago
                                'Por cuanto ira el apunte
                                Im = CCur(Text3(2).Tag)
                                Im = ImporteFormateado(.SubItems(4)) + Im   'Pq im sera negativo
                            End If
                        End If
                    End If
                    
                    
                    
                    
                    
                    
                    VaAlDebe = True
                    Total = Total - Im
                    If Im < 0 Then
                        If Not vParam.abononeg Then
                               VaAlDebe = False
                               Im = -Im
                        End If
                    End If
                    'numdocum, ampconce , codmacta, timporteD,timporteH, ctacontr
                    SQL = FP.condepro & ",'" & DevNombreSQL(.SubItems(1)) & "','"
        
                    SQL = SQL & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "','" & .Tag & "',"

                    'Importe
                    If VaAlDebe Then
                        SQL = SQL & TransformaComasPuntos(CStr(Im)) & ",NULL"
                    Else
                        SQL = SQL & "NULL," & TransformaComasPuntos(CStr(Im))
                    End If
                    
                    'Contrapartida
                    Ampliacion = ""
                    If Text1(0).Text = "" Then
                        If FP.ctrhapro = 1 Then Ampliacion = Trim(RecuperaValor(CadenaDesdeOtroForm, 5))
                    Else
                        Ampliacion = Text1(0).Text
                    End If

                    
                    If Ampliacion <> "" Then
                        SQL = SQL & ",'" & Ampliacion & "',"
                    Else
                        SQL = SQL & ",NULL,"
                    End If
                    
                    
                    'Habran dos pipes.
                    '   1.- lo que tengo que insertar en hlinapu
                    '   2.- El select prparado para eliminar el cobro / pago
                    '   3.- el where para buscar la factura
                    Ampliacion = "|" & CadenaUpdate & " WHERE `codmacta`='" & .Tag & "' and `numfactu`='" & DevNombreSQL(.SubItems(1))
                    Ampliacion = Ampliacion & "' and `fecfactu`='" & Format(.SubItems(2), FormatoFecha) & "' and `numorden`=" & .SubItems(3) & "|"
                    '
                    Ampliacion = Ampliacion & " WHERE `codmacta`='" & .Tag & "' and `numfactu`='" & DevNombreSQL(.SubItems(1))
                    Ampliacion = Ampliacion & "' and `fecfactu`='" & Format(.SubItems(2), FormatoFecha) & "' and `numorden`=" & .SubItems(3) & "|"
                    CPro.Add SQL & Ampliacion
                    
                    
  
                End With
            End If
    Next NumRegElim

    'El ajuste de la linea del banco
     If SQL <> "" And (ItmVto Is Nothing) Then
     
        'Una pequeña comprobacion
        'Valor calculado ahora: Total
        '    "    "      antes: text3(2).text
     
         Im = ImporteFormateado(Text3(2).Text)
         If Im <> Total Then
            CampoAnterior = "ERROR importe calculado"
            SQL = ""
        Else
            If Im <> 0 Then
                'Meteremos, o bien en la lista de cobro, o bien en la de pagos, en funcion del importe
                SQL = ""
                NumRegElim = 0
                Ampliacion = "Compensa:" & Text1(0).Text & " // "
                Do
                    NumRegElim = NumRegElim + 1
                    SQL = RecuperaValor(Text4.Tag, CInt(NumRegElim))
                    If SQL <> "" Then Ampliacion = Ampliacion & " " & SQL
                Loop Until SQL = ""
               
                
                Ampliacion = Mid(Ampliacion, 1, 30)
                                                
                VaAlDebe = True
                SQL = FP.condepro
                If Im < 0 Then
                    SQL = FP.condecli
                    VaAlDebe = False
                    Im = -Im
                End If
                
                'coconce numdocum, ampconce , codmacta, timporteD,timporteH, ctacontr
                SQL = SQL & ",'COMPENSA.','" & DevNombreSQL(Ampliacion) & "','" & RecuperaValor(CadenaDesdeOtroForm, 5) & "',"
                If VaAlDebe Then
                    SQL = SQL & TransformaComasPuntos(CStr(Im)) & ",NULL"
                Else
                    SQL = SQL & "NULL," & TransformaComasPuntos(CStr(Im))
                End If
                SQL = SQL & ",NULL,||" 'No elimna cobro/pago
                CPro.Add SQL
                                
        
            End If
            CampoAnterior = ""
        End If
    End If
    Set FP = Nothing
        
        
    If SQL <> "" Then
    


        SQL = "Los efectos serán modificados despues de contabilizar la compensación." & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then CrearColecciones = True

    Else
        If CampoAnterior = "" Then CampoAnterior = "No se ha seleccionado ningún vencimiento."
        MsgBox CampoAnterior, vbExclamation
    End If
    Exit Function
ECrearColecciones:
    MuestraError Err.Number
End Function



Private Function EstableceVtoQueTotaliza(Indice As Integer) As Integer

    EstableceVtoQueTotaliza = 0

    'Vamos a buscar el vencimiento
    'Recorremos desde el final
    'Y el primero que le quepa la diferencia.... ese lo devuelvo
    For NumRegElim = Me.lw1(Indice).ListItems.Count To 1 Step -1
        If lw1(Indice).ListItems(NumRegElim).Checked Then
            Im = ImporteFormateado(lw1(Indice).ListItems(NumRegElim).SubItems(4))
            If Im > Abs(CCur(Text3(2).Tag)) Then
                    EstableceVtoQueTotaliza = lw1(Indice).ListItems(NumRegElim).Index
                    Exit Function
            End If
        End If
    Next






End Function

Private Function ForzarVtoQueTotaliza(Indice As Integer) As Integer
    ForzarVtoQueTotaliza = 0


    'Vamos a forzar el vencimiento
    'Recorremos desde el final
    'Y el primero que le quepa la diferencia.... ese lo devuelvo
    For NumRegElim = Me.lw1(Indice).ListItems.Count To 1 Step -1
        If lw1(Indice).ListItems(NumRegElim).Checked Then
            ForzarVtoQueTotaliza = lw1(Indice).ListItems(NumRegElim).Index
            Exit Function
        End If
    Next






End Function


Private Sub ValoresConceptosPorDefecto(Leer As Boolean, ByRef CDe As Integer, ByRef CHa As Integer)
Dim NF As Integer
Dim C As String
On Error GoTo EValoresConceptosPorDefecto

    NF = FreeFile
    If Leer Then
        
        Open App.Path & "\Concomp.dat" For Input As #NF
        'Debe
        C = ""
        If Not EOF(NF) Then Line Input #NF, C
        C = Trim(C)
        If C <> "" Then
            If Not IsNumeric(C) Then C = CDe
        Else
            C = CDe
        End If
        CDe = CInt(C)
        
        'Haber
        C = ""
        If Not EOF(NF) Then Line Input #NF, C
        If C <> "" Then
            If Not IsNumeric(C) Then C = CHa
        Else
            C = CHa
        End If
        CHa = CInt(C)
        Close #NF
    Else
        Open App.Path & "\Concomp.dat" For Output As #NF
        Print #NF, CDe
        Print #NF, CHa
        Close #NF

    End If

    Exit Sub
EValoresConceptosPorDefecto:
    Err.Clear
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        KEYCuentas KeyAscii, 1
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub
