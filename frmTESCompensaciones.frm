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
      Begin VB.CheckBox chkNo_x_NIF 
         Alignment       =   1  'Right Justify
         Caption         =   "No vincular pagos por nif"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10560
         TabIndex        =   15
         ToolTipText     =   "Solo tendr� en cuenta el codigo cuenta contable"
         Top             =   240
         Width           =   2295
      End
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
         Top             =   6720
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
         Top             =   6720
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
         Top             =   6720
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
         Height          =   5172
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
         NumItems        =   8
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
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ParaHCO"
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
         NumItems        =   7
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
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ParaHCO"
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
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   6000
         Picture         =   "frmTESCompensaciones.frx":0006
         ToolTipText     =   "quitar seleccion"
         Top             =   6375
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   6360
         Picture         =   "frmTESCompensaciones.frx":0150
         Top             =   6375
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "frmTESCompensaciones.frx":029A
         ToolTipText     =   "quitar seleccion"
         Top             =   6375
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   480
         Picture         =   "frmTESCompensaciones.frx":03E4
         Top             =   6375
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   7290
         Picture         =   "frmTESCompensaciones.frx":052E
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
         Left            =   2640
         TabIndex        =   13
         Top             =   6720
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
         Picture         =   "frmTESCompensaciones.frx":0F30
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



Dim Sql As String   'Cadena de uso comun
Dim Im As Currency
Dim CampoAnterior As String
Dim CadNif As String

Dim vCP As Ctipoformapago
Dim CadeCompenHco As String


Private Sub chkNo_x_NIF_Click()
    CargarListView 1
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
Dim IT As ListItem
Dim AumentaElImporteDelVto As Boolean
Dim IndiceListView As Integer
Dim ModificarVto As Boolean  'No pone el impcobrado, pone vto el total que queda de comensar
Dim B As Boolean

    Dim LCob As Collection
    Dim LPag As Collection
    
    'COmprobaciones
    'Que hay seleccionado algun vencimiento
    Sql = ""
    For NumRegElim = 1 To lw1(0).ListItems.Count
        If lw1(0).ListItems(NumRegElim).Checked Then
            Sql = "1"
            Exit For
        End If
    Next
    If Sql <> "" Then
        Sql = ""
        For NumRegElim = 1 To lw1(1).ListItems.Count
            If lw1(1).ListItems(NumRegElim).Checked Then
                Sql = "1"
                'Nos salimos.
                Exit For
            End If
        Next
    End If
    If Sql = "" Then
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
            Sql = CStr(EstableceVtoQueTotaliza(0))
            If Sql <> "0" Then Set IT = lw1(0).ListItems(CInt(Sql))
        Else
            'Estoy COBRANDO mas que pagando
            Sql = CStr(EstableceVtoQueTotaliza(1))
            If Sql <> "0" Then Set IT = lw1(1).ListItems(CInt(Sql))
        End If
        
        
        'Marzo 2009
        'Si incrementa un vto pq el importe es mayor del que habia.
        If IT Is Nothing Then
        
            'NO dejamos que el impte de un vto aumente.
            MsgBox "El importe a compensar no se puede realizar sobre un �nico vencimiento", vbExclamation
            If False Then
                '
                'AQUI , de momento, NO entra
                AumentaElImporteDelVto = True
                'No hay ningun vto donde compensar.
                'Seleccionare el ultimo seleccionado del listview que corresponda
                If CCur(Text3(2).Tag) > 0 Then
                    Sql = CStr(ForzarVtoQueTotaliza(0))
                    Set IT = lw1(0).ListItems(CInt(Sql))
                Else
                    'Estoy COBRANDO mas que pagando
                    Sql = CStr(ForzarVtoQueTotaliza(1))
                    Set IT = lw1(1).ListItems(CInt(Sql))
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
        Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(CDC))
        CadenaDesdeOtroForm = vCP.conhacli & "|" & Sql & "|"
        Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(CDP))
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & vCP.condepro & "|" & Sql & "|"
    Else
        CadenaDesdeOtroForm = "||||"
    End If
    
    'Le indico si puede realizar la compensacion sobre un vto, o no
    If IT Is Nothing Then
        '0:No
        Sql = "0|Nada|"
    Else
        '1: Si
        Sql = "1|" & IT.Index & "|"
    End If
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Sql
    Set vCP = Nothing
    Sql = ""
    
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
                        Sql = lw1(IndiceListView).ListItems(NumRegElim).Text
                    Else
                        'pagos
                        Sql = ""
                    End If
                    
                    Sql = "Fact: " & Sql & lw1(IndiceListView).ListItems(NumRegElim).SubItems(1) & " ,vto " & lw1(IndiceListView).ListItems(NumRegElim).SubItems(3) & _
                            " de fecha " & lw1(IndiceListView).ListItems(NumRegElim).SubItems(2)
                    
                    frmTESListado.InsertaItemComboCompensaVto Sql, CInt(NumRegElim)
                End If
            End If
        Next
    End If
    
    frmTESListado.Show vbModal
    
    If CadenaDesdeOtroForm <> "" Then
        
       'Compruebo que ninguna de las dos cuentas esta bloqueda para le fecha de contabilizacion
        If CuentaBloqeada(Text1(0).Text, RecuperaValor(CadenaDesdeOtroForm, 4), True) Then Exit Sub
        'Compruebo que ninguna de las dos esta bloqueda para le fecha de contabilizacion
         
        Sql = Text4.Tag
        While Sql <> ""
             NumRegElim = InStr(1, Sql, "|")
             If NumRegElim = 0 Then
                 Sql = ""
             Else
                 If CuentaBloqeada(Mid(Sql, 1, NumRegElim - 1), RecuperaValor(CadenaDesdeOtroForm, 4), True) Then Exit Sub
                 Sql = Mid(Sql, NumRegElim + 1)
             End If
        Wend
                    
                    
        ModificarVto = RecuperaValor(CadenaDesdeOtroForm, 9) = "1"
        'Le quito el ultmo pipe para dejarlo como estaba
        CadenaDesdeOtroForm = Left(CadenaDesdeOtroForm, Len(CadenaDesdeOtroForm) - 2)     'quito el pipe  y el value
        
        'A�ado las obsrvaciones
        'Le quitomel ultmo pipe
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 1, Len(CadenaDesdeOtroForm) - 1)
        'Comprueno si lleva contra un vto o NO
        NumRegElim = InStrRev(CadenaDesdeOtroForm, "|")
        Sql = Mid(CadenaDesdeOtroForm, NumRegElim + 1)
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 1, NumRegElim - 1)
        If Sql = "0" Then
            'NO ha seleccionado el vto, con lo cual pongo el IT a nothing
            Set IT = Nothing
            
        Else
            'Va a compensar contra un vto. Si el vto va a aumentar entonces le pregunto si desea continuar
          
            If IT.Index <> Val(Sql) Then
                'Ha cambiado el VTO que le ofertabamos nosotros
                Set IT = lw1(IndiceListView).ListItems(CInt(Val(Sql)))
            End If
            'Aqui NO debe de ebtrar
            If AumentaElImporteDelVto Then
                Sql = "El importe del vencimiento Factura: "
                Sql = Sql & IT.SubItems(1) & "   n�" & IT.SubItems(3) & "  de fecha " & IT.SubItems(2)
                Sql = Sql & " se va a incrementar"
                
                Sql = Sql & vbCrLf & "�Desea continuar?"
                If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            End If
        End If
        
        'ASigno la nueva forma de pago del vto resultante (o en su defecto obvio el dato
        'Con lo cual voy a quitar el utlimi pipe que es la FP
        NumRegElim = InStrRev(CadenaDesdeOtroForm, "|")
        Sql = Mid(CadenaDesdeOtroForm, NumRegElim + 1)
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 1, NumRegElim)
        IndiceListView = -1
        If Not IT Is Nothing Then
            If Sql <> "" Then
                If IsNumeric(Sql) Then IndiceListView = Val(Sql)
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
            
            CadeCompenHco = ""
            B = False
                                                                            'IndiceListView: Si compensa cn vto y quiere cambiar la forma de pago
            If CrearColecciones(LCob, LPag, vCP, IT, AumentaElImporteDelVto, IndiceListView, ModificarVto) Then
                B = ContabilizarCompensaciones(LCob, LPag, CadenaDesdeOtroForm, AumentaElImporteDelVto)
            End If
                
                
            If B Then
                'LOG
                Dim SqlLog As String
                '09/abri/20   NOhay log. Hay hco
                'YA NO HAY LOG
                
                
                'Abril2020
                'Tabla compensaciones
                SqlLog = DevuelveDesdeBD("max(codigo)", "compensaclipro", "1", "1")
                If SqlLog = "" Then SqlLog = 0
                NumRegElim = Val(SqlLog) + 1
                i = 0
                
                'insert into compensaclipro_facturas(codigo,linea,EsCobro,codmacta,numserie,numfactu,fecfactu, numorden,importe,gastos,impcobro,fechavto ,destino,compensado )
                SqlLog = "INSERT INTO compensaclipro(codigo,autom,fecha,login,PC,codmacta,nommacta,resultado,fechahora) VALUES (" & NumRegElim & ",0,"
                SqlLog = SqlLog & DBSet(RecuperaValor(CadenaDesdeOtroForm, 4), "F") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(vUsu.PC, "T") & ","
                SqlLog = SqlLog & DBSet(Me.Text1(0).Text, "T") & "," & DBSet(Me.Text2(0).Text, "T") & "," & DBSet(Me.Text3(2).Text, "N") & "," & DBSet(Now, "FH") & ")"
                If Ejecuta(SqlLog, False) Then
                    CadeCompenHco = Replace(CadeCompenHco, "###codcomep###", CStr(NumRegElim))
                    CadeCompenHco = Mid(CadeCompenHco, 2) 'quito la primera coma
                    SqlLog = "insert into compensaclipro_facturas(codigo,linea,EsCobro,codmacta,numserie,numfactu,fecfactu, numorden,importe,gastos,impcobro,fechavto ,destino,compensado ) VALUES " & CadeCompenHco
                    If Ejecuta(SqlLog, False) Then i = 1
                End If
                
                
                If i = 0 Then
                    'Ha habiado un error. Meto log, que avisen soporte
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
                
                    MsgBoxA "Error insertando en tabla historico compensaciones. Avise soporte t�cnico", vbExclamation
                
                End If
           End If
           CadenaDesdeOtroForm = ""
           CargarListView 0
           CargarListView 1
           
           

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
    Sql = CadenaSeleccion
End Sub

Private Sub imgCheck_Click(Index As Integer)
    
    NumRegElim = IIf(Index < 2, 0, 1)
    J = IIf((Index Mod 2) = 0, 0, 1)
    Im = 0
    If lw1(NumRegElim).ListItems.Count = 0 Then Exit Sub
    For i = 1 To lw1(NumRegElim).ListItems.Count
        lw1(NumRegElim).ListItems(i).Checked = J = 1
        If J = 1 Then Im = Im + ImporteFormateado(lw1(NumRegElim).ListItems(i).SubItems(4))
    Next
    'Arrastro
    Text3(NumRegElim).Tag = Im
    CalculaImportes
    
    
End Sub

Private Sub imgCuentas_Click(Index As Integer)
    
    
    If Index = 0 Then
        'Avisar si ya han cargado datos
         Screen.MousePointer = vbHourglass
         Set frmCCtas = New frmColCtas
         Sql = ""
         CampoAnterior = Text1(Index).Text
         frmCCtas.DatosADevolverBusqueda = "0"
         frmCCtas.Show vbModal
         Set frmCCtas = Nothing
         If Sql <> "" Then
            Text1(Index).Text = RecuperaValor(Sql, 1)
            Text2(Index).Text = RecuperaValor(Sql, 2)
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
            If Not CuentaCorrectaUltimoNivel(C, Sql) Then
                MsgBox Sql & " - " & C, vbExclamation
                Sql = ""
                C = ""
            End If
            Text1(Index).Text = C
            Text2(Index).Text = Sql
            If C = "" Then
                PonFoco Text1(Index)
            Else
                ' A�adida esta parte donde seg�n el nif del cliente me voy a buscar las ctas de proveedor con el mismo nif
                
                
                
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim CadResult As String

    CuentasProveedorDelNif = ""

    Sql = "select distinct codmacta from pagos where nifprove = " & DBSet(NIF, "T")
    Sql = Sql & " and impefect - coalesce(imppagad,0) <> 0 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadResult = ""
    
    While Not Rs.EOF
    
        If Rs!codmacta <> "" Then
    
            CadResult = CadResult & Rs!codmacta & "|"
        Else
            MsgBox "Cuenta en PAGOS vacia para NIF: " & NIF, vbExclamation
        End If
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
                Sql = Mid(C, 1, NumRegElim - 1)
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
Dim Aux As String

    On Error GoTo ECargaDatosListview
    


    
    
    If Indice = 0 Then
        Sql = "select numserie,numfactu,fecfactu,numorden,impvenci,impcobro,gastos,codmacta,fecvenci from cobros where"
        Sql = Sql & " codrem is null and anyorem is null and coalesce(transfer,0)=0 "
        'Y que el talon pagare NO este recepcionado
        Sql = Sql & " AND recedocu = 0"
        Sql = Sql & " and  codmacta ='" & Text1(Indice).Text & "'"
    Else
        'En SQL va el codmacta
        CadNif = DevuelveDesdeBD("nifdatos", "cuentas", "codmacta", Sql, "T")
        
        Sql = " and nrodocum is null and pagos.codmacta ='" & Sql & "'"
        If chkNo_x_NIF.Value = 0 Then
            If CadNif <> "" Then
                Sql = Sql & " and pagos.nifprove = " & DBSet(CadNif, "T")
            End If
        End If
        Sql = " WHERE impefect - coalesce(imppagad,0) <> 0 " & Sql 'AND estacaja =0
        Sql = "select numfactu,fecfactu,numorden,impefect,imppagad,pagos.codmacta as codmacta,nomprove as nommacta, numserie,fecefect  FROM pagos " & Sql
    End If
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not miRsAux.EOF
        'Veremos si el importe es positivo, o no
        
        If Indice = 0 Then
            Im = miRsAux!ImpVenci - DBLet(miRsAux!impcobro, "N") + DBLet(miRsAux!Gastos, "N")
            YaEfectuado = DBLet(miRsAux!impcobro, "N")
            'If vParamT.PideFechaImpresionTalonPagare Then
                CargaEnListview = Im <> 0  'nollar de moemnto
            'Else
            '    CargaEnListview = Im > 0
            'End If
        Else
            Im = miRsAux!ImpEfect - DBLet(miRsAux!imppagad, "N")
            YaEfectuado = DBLet(miRsAux!imppagad, "N")
            CargaEnListview = True 'Pase lo que pase
        End If
        'If Im > 0 Then
        'Para el hco
        'codmacta,numserie,numfactu,fecfactu, numorden,importe,gastos,impcobro,fechavto ,compensado,destino
        If CargaEnListview Then
        
    
            Set IT = lw1(Indice).ListItems.Add()
            If Indice = 0 Then
                IT.Text = miRsAux!NUmSerie
                IT.SubItems(1) = miRsAux!numfactu
                IT.SubItems(2) = miRsAux!FecFactu
                IT.SubItems(3) = miRsAux!numorden
                'Importe:
                
            Else
                Aux = DBLet(miRsAux!Nommacta, "T")
                J = 0
                If Aux = "" Then
                    J = 1
                    Aux = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", miRsAux!codmacta, "T")
                    IT.ToolTipText = "Falta datos fiscales en el pago"
                End If
                IT.Text = Mid(Aux, 1, 20)
                'Para que aparezca en el
                If Sql = "" Then
                    If Text4.Text <> "" Then Text4.Text = Text4.Text & vbCrLf
                    Text4.Text = Text4.Text & miRsAux!codmacta & "   " & miRsAux!Nommacta
                    Sql = "D"
                End If
                IT.SubItems(1) = miRsAux!numfactu
                IT.SubItems(2) = miRsAux!FecFactu
                IT.SubItems(3) = miRsAux!numorden
                
            End If
            IT.SubItems(4) = Format(Im, FormatoImporte)
            IT.SubItems(5) = YaEfectuado
            IT.Tag = miRsAux!codmacta
            
            'codmacta,numserie,numfactu,fecfactu, numorden,importe,gastos,impcobro,fechavto ,compensado,destino
            If Indice = 1 Then
                IT.SubItems(6) = miRsAux!NUmSerie
                        
                'Para la hco compensacion
                IT.SubItems(7) = "0," & DBSet(miRsAux!codmacta, "T") & "," & DBSet(miRsAux!NUmSerie, "T") & "," & DBSet(miRsAux!numfactu, "T") & "," & DBSet(miRsAux!FecFactu, "F")
                IT.SubItems(7) = IT.SubItems(7) & "," & DBSet(miRsAux!numorden, "N") & "," & DBSet(miRsAux!ImpEfect, "N") & ",null"
                IT.SubItems(7) = IT.SubItems(7) & "," & DBSet(miRsAux!imppagad, "N", "S") & "," & DBSet(miRsAux!fecefect, "F")
            Else
                IT.SubItems(6) = "1," & DBSet(miRsAux!codmacta, "T") & "," & DBSet(miRsAux!NUmSerie, "T") & "," & Format(miRsAux!numfactu, "000000") & "," & DBSet(miRsAux!FecFactu, "F")
                IT.SubItems(6) = IT.SubItems(6) & "," & DBSet(miRsAux!numorden, "T") & "," & DBSet(miRsAux!ImpVenci, "N") & "," & DBSet(miRsAux!Gastos, "N", "S")
                IT.SubItems(6) = IT.SubItems(6) & "," & DBSet(miRsAux!impcobro, "N", "S") & "," & DBSet(miRsAux!FecVenci, "F")
            End If
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
Dim LineaHcoCompensa As Integer
Dim FrasCli As String
Dim FrasPro As String
Dim FraGastos As String
Dim Gastos As Currency



    '0: NO compensa
    '1: Cobros
    '2: Pagos

    On Error GoTo ECrearColecciones
    CrearColecciones = False

    Total = 0
    Set CCli = New Collection
    Set CPro = New Collection
            
    'Para el hco de compensaciones
    'compensaclipro_facturas( codigo,linea,EsCobro,codmacta,numserie,numfactu,fecfactu,numorden,importe,gastos,impcobro,fechavto ,compensado,destino
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
    
    
    LineaHcoCompensa = 0
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
        Sql = ""
        
        '----------------------------------------------------------------------
        '----------------------------------------------------------------------
        'CLIENTES
        For NumRegElim = 1 To lw1(0).ListItems.Count
            
            If lw1(0).ListItems(NumRegElim).Checked Then
                With lw1(0).ListItems(NumRegElim)
                    
                    'Priemro el trocito para las compensaciones
                    LineaHcoCompensa = LineaHcoCompensa + 1
                    CadeCompenHco = CadeCompenHco & ", (###codcomep###," & LineaHcoCompensa & "," & .SubItems(6) & ",##destino##,##compensado##)"
                    
                    
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
                    
                                        
                    
                    
                    J = 0 'no es el vto destino
                    Im = ImporteFormateado(.SubItems(4))
                    
                    Gastos = 0    'Para ver si a�ade a cuena de gastos
                    FraGastos = ""
                    CadenaUpdate = " `numserie`='" & .Text & "' and numfactu=" & .SubItems(1)
                    CadenaUpdate = CadenaUpdate & " and `fecfactu`='" & Format(.SubItems(2), FormatoFecha) & "' and `numorden`"
                    CadenaUpdate = DevuelveDesdeBD("gastos", "cobros", CadenaUpdate, .SubItems(3))
                    If CadenaUpdate <> "" Then
                        Gastos = CCur(CadenaUpdate)
                        If Gastos > 0 Then
                            If Gastos > Im Then Err.Raise 513, , "Mas gastos que importe vencimiento"
                            Im = Im - Gastos
                        
                            FraGastos = DevuelveDesdeBD("ctabenbanc", "paramtesor", "1", "1")
                            If FraGastos = "" Then Err.Raise 513, , "Falta configurar cta beneficios bancarios"
                        End If
                    End If
                    
                    
                    
                    CadenaUpdate = ""
                    'Si compensa sobre un vto de cobro
                    If CompensaSobreCobros = 1 Then
                        'Hace la comep
                        If lw1(0).ListItems(NumRegElim).Index = ItmVto.Index Then
                            
                            J = 1 'si que es el vto destino
                            
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
                                CadenaUpdate = CadenaUpdate & " " 'Por si queremos a�adir mas camos a updatear
                                
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
                    
                    
                    'Ajustamos el importe que se ha compensado, si a lugar
                    CadeCompenHco = Replace(CadeCompenHco, "##destino##", J)
                    CadeCompenHco = Replace(CadeCompenHco, "##compensado##", DBSet(Im, "N"))
                    
                    Total = Total + Im
                    VaAlDebe = False
                    
                    If Im < 0 Then
                        If Not vParam.abononeg Then
                               VaAlDebe = True
                               Im = -Im
                        End If
                    End If
                    'codconce numdocum, ampconce , codmacta, timporteD,timporteH, ctacontr
                    Sql = FP.condecli & ",'" & .Text & Format(.SubItems(1), "000000") & "','"
                    
        
        
                    Sql = Sql & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "','" & Text1(0).Text & "',"
                    
                    'Por si llevara gastos
                    If Gastos > 0 Then
                        FraGastos = DevNombreSQL(Mid(Ampliacion, 1, 30)) & "','" & FraGastos & "',"
                        FraGastos = FP.condecli & ",'" & .Text & Format(.SubItems(1), "000000") & "','" & FraGastos
                    End If
                    
                    'Importe
                    If VaAlDebe Then
                        Sql = Sql & TransformaComasPuntos(CStr(Im)) & ",NULL"
                    Else
                        Sql = Sql & "NULL," & TransformaComasPuntos(CStr(Im))
                    End If
                    FraGastos = FraGastos & "NULL," & TransformaComasPuntos(CStr(Gastos))
                    
                    
                    
                    
                    'Contrapartida. esta guaddad en ContrapartidaPago
                    Sql = Sql & "," & ContrapartidaPago & ","
                    FraGastos = FraGastos & "," & ContrapartidaPago & ","
                    
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
            
                    CCli.Add Sql & Ampliacion
                    
                    If Gastos > 0 Then
                        CCli.Add FraGastos & "|" & "|"
                    
                    End If
                   
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
                    
                    
                    'Priemro el trocito para las compensaciones
                    LineaHcoCompensa = LineaHcoCompensa + 1
                    CadeCompenHco = CadeCompenHco & ", (###codcomep###," & LineaHcoCompensa & "," & .SubItems(7) & ",##destino##,##compensado##)"
                    
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
                    
                    J = 0 'si es vto destino
                    Im = ImporteFormateado(.SubItems(4))
                    CadenaUpdate = ""
                    'Si compensa sobre un vto de pago
                    If CompensaSobreCobros = 2 Then
                        
                        If lw1(1).ListItems(NumRegElim).Index = ItmVto.Index Then
                            J = 1
                           
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
                                CadenaUpdate = CadenaUpdate & " " 'Por si queremos a�adir mas camos a updatear
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
                    
                    'Ajustamos el importe que se ha compensado, si a lugar
                    CadeCompenHco = Replace(CadeCompenHco, "##destino##", J)
                    CadeCompenHco = Replace(CadeCompenHco, "##compensado##", DBSet(Im, "N"))
                    
                    
                    
                    
                    VaAlDebe = True
                    Total = Total - Im
                    If Im < 0 Then
                        If Not vParam.abononeg Then
                               VaAlDebe = False
                               Im = -Im
                        End If
                    End If
                    'numdocum, ampconce , codmacta, timporteD,timporteH, ctacontr
                    Sql = FP.condepro & ",'" & DevNombreSQL(.SubItems(1)) & "','"
        
                    Sql = Sql & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "','" & .Tag & "',"

                    'Importe
                    If VaAlDebe Then
                        Sql = Sql & TransformaComasPuntos(CStr(Im)) & ",NULL"
                    Else
                        Sql = Sql & "NULL," & TransformaComasPuntos(CStr(Im))
                    End If
                    
                    'Contrapartida
                    Ampliacion = ""
                    If Text1(0).Text = "" Then
                        If FP.ctrhapro = 1 Then Ampliacion = Trim(RecuperaValor(CadenaDesdeOtroForm, 5))
                    Else
                        Ampliacion = Text1(0).Text
                    End If

                    
                    If Ampliacion <> "" Then
                        Sql = Sql & ",'" & Ampliacion & "',"
                    Else
                        Sql = Sql & ",NULL,"
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
                    CPro.Add Sql & Ampliacion
                    
                    
  
                End With
            End If
    Next NumRegElim

    'El ajuste de la linea del banco
     If Sql <> "" And (ItmVto Is Nothing) Then
     
        'Una peque�a comprobacion
        'Valor calculado ahora: Total
        '    "    "      antes: text3(2).text
     
         Im = ImporteFormateado(Text3(2).Text)
         If Im <> Total Then
            CampoAnterior = "ERROR importe calculado"
            Sql = ""
        Else
            If Im <> 0 Then
                'Meteremos, o bien en la lista de cobro, o bien en la de pagos, en funcion del importe
                Sql = ""
                NumRegElim = 0
                Ampliacion = "Compensa:" & Text1(0).Text & " // "
                Do
                    NumRegElim = NumRegElim + 1
                    Sql = RecuperaValor(Text4.Tag, CInt(NumRegElim))
                    If Sql <> "" Then Ampliacion = Ampliacion & " " & Sql
                Loop Until Sql = ""
               
                
                Ampliacion = Mid(Ampliacion, 1, 30)
                                                
                VaAlDebe = True
                Sql = FP.condepro
                If Im < 0 Then
                    Sql = FP.condecli
                    VaAlDebe = False
                    Im = -Im
                End If
                
                'coconce numdocum, ampconce , codmacta, timporteD,timporteH, ctacontr
                Sql = Sql & ",'COMPENSA.','" & DevNombreSQL(Ampliacion) & "','" & RecuperaValor(CadenaDesdeOtroForm, 5) & "',"
                If VaAlDebe Then
                    Sql = Sql & TransformaComasPuntos(CStr(Im)) & ",NULL"
                Else
                    Sql = Sql & "NULL," & TransformaComasPuntos(CStr(Im))
                End If
                Sql = Sql & ",NULL,||" 'No elimna cobro/pago
                CPro.Add Sql
                                
        
            End If
            CampoAnterior = ""
        End If
    End If
    Set FP = Nothing
        
        
    If Sql <> "" Then
    


        Sql = "Los efectos ser�n modificados despues de contabilizar la compensaci�n." & vbCrLf & "�Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then CrearColecciones = True

    Else
        If CampoAnterior = "" Then CampoAnterior = "No se ha seleccionado ning�n vencimiento."
        MsgBox CampoAnterior, vbExclamation
    End If
    Exit Function
ECrearColecciones:
    MuestraError Err.Number, , Err.Description
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
