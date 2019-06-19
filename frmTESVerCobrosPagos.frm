VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESVerCobrosPagos 
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14640
   Icon            =   "frmTESVerCobrosPagos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   14640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frame 
      Height          =   1095
      Left            =   90
      TabIndex        =   4
      Top             =   30
      Width           =   14445
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "Regresar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12480
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Regresar"
         Top             =   360
         Width           =   1365
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
         Left            =   1440
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkReme 
         Caption         =   "Mostrar riesgo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3510
         TabIndex        =   2
         Top             =   450
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1140
         Top             =   420
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   420
         Width           =   840
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   750
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTESVerCobrosPagos.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTESVerCobrosPagos.frx":686E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTESVerCobrosPagos.frx":6B88
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   6270
      Width           =   14415
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   2370
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   60
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   5400
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   60
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   8580
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   60
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   12060
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   60
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Riesgo Talón/Pagaré"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   3
         Left            =   60
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   2520
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Riesgo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   4380
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   " PENDIENTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Index           =   1
         Left            =   10650
         TabIndex        =   9
         Top             =   120
         Width           =   1290
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Vencido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   7530
         TabIndex        =   7
         Top             =   120
         Width           =   990
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5025
      Left            =   90
      TabIndex        =   0
      Top             =   1200
      Width           =   14475
      _ExtentX        =   25532
      _ExtentY        =   8864
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Menu mnContextual 
      Caption         =   "Contextual"
      Visible         =   0   'False
      Begin VB.Menu mnNumero 
         Caption         =   "Poner numero Talón/Pagaré"
      End
      Begin VB.Menu mnbarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSelectAll 
         Caption         =   "Seleccionar todos"
      End
      Begin VB.Menu mnQUitarSel 
         Caption         =   "Quitar selección"
      End
   End
End
Attribute VB_Name = "frmTESVerCobrosPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Situacion As Byte
Public vSql As String
Public Cobros As Boolean
Public OrdenarEfecto As Boolean
Public Regresar As Boolean
Public vTextos As String  'Dependera de donde venga
Public Tipo As Byte
Public SegundoParametro As String
Public ContabTransfer As Boolean
Public OrdenacionEfectos As Byte


    'Diversas utilidades
    '-------------------------------------------------------------------------------
    'Para las transferencias me dice que transferencia esta siendo creada/modificada
    '
    'Para mostrar un check con los efectos k se van a generar en remesa y/o pagar
 
 
 ' 13 Mayo 08
    ' Cuando contabilice el los cobros por tarjeta entonces
    ' si lleva gastos los añadire
Public ImporteGastosTarjeta_ As Currency   'Para cuando viene de recepciondocumentos pondre el importe que le falta
                                          ' y asi ofertarlo al divisonvencimiento
     '-ABRIL 2014.  Navarres. Llevara el % interes
 
 
 
 
'Agosto 2009
'Desde recepcion de talones.
'Tendra la posibilidad de desdoblar un vencimiento
Public DesdeRecepcionTalones As Boolean
 
'Febrero 2010
'Para el pago de talones y pagareses ;)
'Enviara el nº de talon/pagare
Public NumeroTalonPagere As String


'Marzo 2013
'Cuando cobro/pago un mismo clie/prov aparecera un icono para poder añadir
'cualquier cobro /pago del mismo. Se contabilizaran con los datos pendientes
Public CodmactaUnica As String



Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim Cad As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Fecha As Date
Dim Importe As Currency
Dim Vencido As Currency
Dim impo As Currency
Dim riesgo As Currency

Dim ImpSeleccionado As Currency
Dim i As Integer
Private PrimeraVez As Boolean
Private SeVeRiesgo As Boolean
Dim RiesTalPag As Currency
Private SeVeRiesgoTalPag As Boolean
Private FechaAsiento As Date
Private vp As Ctipoformapago
Private SubItemVto As Integer

Private DescripcionTransferencia As String
Private GastosTransferencia As Currency



Dim CampoOrden As String
Dim Orden As Boolean
Dim Campo2 As Integer


Dim OrdenEnFicheroXDF As String


Private Sub chkReme_Click()
    SeVeRiesgo = False
    If Not OrdenarEfecto Then
        'Ver cobros pagos
        If Cobros And (Me.chkReme.Value = 1) Then SeVeRiesgo = True
    End If
    Label2(2).visible = SeVeRiesgo
    Text2(2).visible = SeVeRiesgo
    Label2(3).visible = SeVeRiesgo And Cobros
    Text2(3).visible = SeVeRiesgo And Cobros
    CargaList
End Sub





Private Sub cmdRegresar_Click()
    If Not (ListView1.SelectedItem Is Nothing) Then
        If Cobros Then
            CadenaDesdeOtroForm = ListView1.SelectedItem.Text & "|" & ListView1.SelectedItem.SubItems(1) & "|"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & ListView1.SelectedItem.SubItems(2) & "|" & ListView1.SelectedItem.SubItems(4) & "|"
        Else
            'Pagos proveedores
            CadenaDesdeOtroForm = ListView1.SelectedItem.Text & "|"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & ListView1.SelectedItem.SubItems(1) & "|" & ListView1.SelectedItem.SubItems(2) & "|"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & ListView1.SelectedItem.SubItems(4) & "|" & ListView1.SelectedItem.Tag & "|"
        End If
    Else
        CadenaDesdeOtroForm = ""
    End If
    Unload Me
End Sub

Private Sub Refrescar()
    Screen.MousePointer = vbHourglass
    CargaList
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        SeVeRiesgo = True
        Me.Refresh
        espera 0.1
        'Cargamos el LIST
        CargaList
        
        'PonerFocoLw Me.ListView1
        
    End If
    Screen.MousePointer = vbDefault
End Sub
 
Private Sub Form_Load()

    PrimeraVez = True
    Limpiar Me
    Me.Icon = frmppal.Icon
    For i = 1 To imgFecha.Count - 1
        Me.imgFecha(i).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next i
    
    
    CargaIconoListview Me.ListView1
    ListView1.Checkboxes = OrdenarEfecto
    Text1.Enabled = Not OrdenarEfecto
    Me.chkReme.visible = False
    
'    imgFecha(2).Visible = False 'Para cambiar la fecha de contabilizacion de los pagos

    'Cobros y pagos pendientes
    CampoOrden = ""
    LeerFiltroOrdenacion True
    
    If Cobros Then
        
    
        Caption = "Cobros pendientes"
        chkReme.Value = 1
        chkReme.visible = True
    Else
        Caption = "Pagos pendientes"
        
        
    End If
    
    
    i = 0
    If Cobros And (Tipo = 2 Or Tipo = 3) Then i = 1
    Me.mnBarra1.visible = i = 1
    Me.mnNumero.visible = i = 1
    'Efectuar cobros
    Me.cmdRegresar.visible = Regresar
    ListView1.SmallIcons = Me.ImageList1
    Text1.Text = Format(Now, "dd/mm/yyyy")
    Text1.Tag = "'" & Format(Now, FormatoFecha) & "'"
    CargaColumnas
    
    
    'Octubre 2014
    'Norma 57 pagos ventanilla
    'Si en el select , en el SQL, viene un
    If Cobros And Tipo = 0 Then
'--
'        If InStr(1, vSQL, "from tmpconext  WHERE codusu") > 0 Then chkPorFechaVenci.Value = 1
    End If
End Sub

Private Sub Form_Resize()
Dim i As Integer
Dim H As Integer
    If Me.WindowState = 1 Then Exit Sub  'Minimizar
    If Me.Height < 2700 Then Me.Height = 2700
    If Me.Width < 2700 Then Me.Width = 2700
    
    'Situamos el frame y demas
    Me.frame.Width = Me.Width - 120
    Me.Frame1.Left = Me.Width - 120 - Me.Frame1.Width
    Me.Frame1.top = Me.Height - Frame1.Height - 540 '360
    
    Me.ListView1.top = Me.frame.Height + 60
    Me.ListView1.Height = Me.Frame1.top - Me.ListView1.top - 60
    Me.ListView1.Width = Me.frame.Width
    
    'Las columnas
    H = ListView1.Tag
    ListView1.Tag = ListView1.Width - ListView1.Tag - 320 'Del margen
    For i = 1 To Me.ListView1.ColumnHeaders.Count
        If InStr(1, ListView1.ColumnHeaders(i).Tag, "%") Then
            Cad = (Val(ListView1.ColumnHeaders(i).Tag) * (Val(ListView1.Tag)) / 100)
        Else
            'Si no es de % es valor fijo
            Cad = Val(ListView1.ColumnHeaders(i).Tag)
        End If
        Me.ListView1.ColumnHeaders(i).Width = Val(Cad)
    Next i
    ListView1.Tag = H
End Sub


Private Sub CargaColumnas()
Dim ColX As ColumnHeader
Dim Columnas As String
Dim Ancho As String
Dim ALIGN As String
Dim NCols As Integer
Dim i As Integer

    ListView1.ColumnHeaders.Clear
   If Cobros Then
        NCols = 11
        Columnas = "Serie|Factura|F.Factura|F. VTO|Nº|CLIENTE|Tipo|Importe|Gasto|Cobrado|Pendiente|"
        Ancho = "800|10%|12%|12%|520|23%|840|12%|8%|11%|12%|"
        ALIGN = "LLLLLLLDDDD"
        
        
        ListView1.Tag = 2200  'La suma de los valores fijos. Para k ajuste los campos k pueden crecer
        
        If Tipo = 2 Or Tipo = 3 Then
            ''Si es un talon o pagare entonces añadire un campo mas
            NCols = NCols + 1
            Columnas = Columnas & "Nº Documento|"
            Ancho = Ancho & "2500|"
            ALIGN = ALIGN & "L"
        End If
   Else
        NCols = 10
        Columnas = "Serie|Nº Factura|F. Fact|F. VTO|Nº|PROVEEDOR|Tipo|Importe|Pagado|Pendiente|"
        Ancho = "800|12%|11%|11%|400|25%|800|12%|11%|12%|"
        ALIGN = "LLLLLLLDDD"
        ListView1.Tag = 1600  'La suma de los valores fijos. Para k ajuste los campos k pueden crecer
    End If
        
   For i = 1 To NCols
        Cad = RecuperaValor(Columnas, i)
        If Cad <> "" Then
            Set ColX = ListView1.ColumnHeaders.Add()
            ColX.Text = Cad
            'ANCHO
            Cad = RecuperaValor(Ancho, i)
            ColX.Tag = Cad
            'align
            Cad = Mid(ALIGN, i, 1)
            If Cad = "L" Then
                'NADA. Es valor x defecto
            Else
                If Cad = "D" Then
                    ColX.Alignment = lvwColumnRight
                Else
                    'CENTER
                    ColX.Alignment = lvwColumnCenter
                End If
            End If
        End If
    Next i

End Sub


Private Sub CargaList()
On Error GoTo ECargando

    Me.MousePointer = vbHourglass
    Screen.MousePointer = vbHourglass
    SeVeRiesgo = (chkReme.Value = 1)
    SeVeRiesgoTalPag = False
    If Not OrdenarEfecto Then
        'Ver cobros pagos
'        If Cobros And (Me.chkReme.Value = 1) Then SeVeRiesgo = True
    End If
    Label2(2).visible = SeVeRiesgo
    Text2(2).visible = SeVeRiesgo
    Label2(3).visible = SeVeRiesgo And Cobros
    Text2(3).visible = SeVeRiesgo And Cobros
    
    
    Set Rs = New ADODB.Recordset
    Fecha = CDate(Text1.Text)
    ListView1.ListItems.Clear
    Importe = 0
    Vencido = 0
    riesgo = 0
    ImpSeleccionado = 0
    Screen.MousePointer = vbHourglass
    If Cobros Then
        CargaCobros
    Else
        CargaPagos
    End If
    If OrdenarEfecto Then
        Text2(2).Text = "0,00"
        Label2(2).Caption = "Selec."
        Label2(2).visible = True
        Text2(2).visible = True
        Label2(3).visible = True And Cobros
        Text2(3).visible = True And Cobros
    End If
    
ECargando:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        If Cobros Then
            CampoOrden = "cobros.fecvenci"
        Else
            CampoOrden = "pagos.fecefect"
        End If
    End If
    Text2(0).Text = Format(Importe, FormatoImporte)
    Text2(1).Text = Format(Vencido, FormatoImporte)
    
    Text2(2).Text = Format(riesgo, FormatoImporte)
    Text2(3).Text = Format(RiesTalPag, FormatoImporte)
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    Set Rs = Nothing
End Sub

Private Sub CargaCobros()
Dim Inserta As Boolean

    RiesTalPag = 0
    Cad = DevSQL
    
    'ORDENACION
    If CampoOrden = "" Then CampoOrden = "cobros.fecvenci"
    Cad = Cad & " ORDER BY " & CampoOrden
    'If Orden Then Cad = Cad & " DESC"
    If CampoOrden <> "cobros.fecvenci" Then Cad = Cad & ", cobros.fecvenci"
    
    
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Inserta = True
        '[Monica]16/08/2016: solo en el caso de pendientes de cobro no lo veo todo,  situacion = 0
        '                    nuevo parametro de situacion
        If Situacion = 0 Then
            If Rs!tipoformapago = vbTipoPagoRemesa Then
                If Not OrdenarEfecto Then
                 
                    If Not SeVeRiesgo Then
                    ' por lo de mc añado la condicion And DBLet(RS!siturem, "T") > "B"
                        If DBLet(Rs!Codrem, "N") > 0 And DBLet(Rs!siturem, "T") > "B" Then
                            Inserta = False
                            
                        End If
                    ' añadido lo que pide Mc de que se vean las remesas que tengan situacion B
                    Else
                        If (DBLet(Rs!Codrem, "N") > 0 And DBLet(Rs!siturem, "T") > "B") Then Inserta = False
                    End If
                End If
                
            ElseIf Rs!tipoformapago = vbTalon Or Rs!tipoformapago = vbPagare Then
                If Not OrdenarEfecto And Not SeVeRiesgoTalPag Then
                    If Rs!recedocu = 1 Then Inserta = False
                End If
            End If
        End If
        
        If Inserta Then
    
            InsertaItemCobro
            
            
        End If  'de insertar
        Rs.MoveNext
    Wend
    Rs.Close
End Sub


Private Sub InsertaItemCobro()
Dim vImporte As Currency
Dim DiasDif As Long
Dim ImpAux As Currency

    Set ItmX = ListView1.ListItems.Add()
    
    ItmX.Text = Rs!NUmSerie
    ItmX.SubItems(1) = Rs!NumFactu
    ItmX.SubItems(2) = Format(Rs!FecFactu, "dd/mm/yyyy")
    ItmX.SubItems(3) = Format(Rs!FecVenci, "dd/mm/yyyy")
    ItmX.SubItems(4) = Rs!numorden
    ItmX.SubItems(5) = DBLet(Rs!Nommacta, "T")
    ItmX.SubItems(6) = DBLet(Rs!siglas, "T")
    
    ItmX.SubItems(7) = Format(Rs!ImpVenci, FormatoImporte)
    vImporte = DBLet(Rs!Gastos, "N")
    
    'Gastos
    ItmX.SubItems(8) = Format(vImporte, FormatoImporte)
    vImporte = vImporte + Rs!ImpVenci
    
    If Not IsNull(Rs!impcobro) Then
        ItmX.SubItems(9) = Format(Rs!impcobro, FormatoImporte)
        impo = vImporte - Rs!impcobro
        ItmX.SubItems(10) = Format(impo, FormatoImporte)
    Else
        impo = vImporte
        ItmX.SubItems(9) = "0.00"
        ItmX.SubItems(10) = Format(vImporte, FormatoImporte)
    End If
    If Rs!tipoformapago = vbTipoPagoRemesa Then
        '81--->
        'asc("Q") =81 or asc("B") = 66
        If Asc(Right(" " & DBLet(Rs!siturem, "T"), 1)) = 81 Or Asc(Right(" " & DBLet(Rs!siturem, "T"), 1)) = 66 Then
            riesgo = riesgo + vImporte
        Else
           
        End If
    
    ElseIf Rs!tipoformapago = vbTalon Or Rs!tipoformapago = vbPagare Then
        If OrdenarEfecto Then
            'If RS!ImpVenci > 0 Then ItmX.SubItems(11) = DBLet(RS!reftalonpag, "T")
        End If
        If SeVeRiesgoTalPag Then
            If Rs!recedocu = 1 Then RiesTalPag = RiesTalPag + DBLet(Rs!impcobro, "N")
        End If
    End If
    
    If Rs!tipoformapago = vbTarjeta Then
        'Si tiene el parametro y le ha puesto valor
        If vParamT.IntereseCobrosTarjeta2 > 0 And ImporteGastosTarjeta_ > 0 Then
            DiasDif = 0
            If Rs!FecVenci < Fecha Then DiasDif = DateDiff("d", Rs!FecVenci, Fecha)
            If DiasDif > 0 Then
                'Si ya tenia gastos.
                If DBLet(Rs!Gastos, "N") > 0 Then
                    ItmX.ListSubItems(8).Bold = True
                    ItmX.ListSubItems(8).ForeColor = vbRed
                End If
                
                ImpAux = ((ImporteGastosTarjeta_ / 365) * DiasDif) / 100
                ImpAux = Round(ImpAux * impo, 2)
                
                impo = impo + ImpAux
                ItmX.SubItems(10) = Format(impo, FormatoImporte)
                'La de gastos
                ImpAux = DBLet(Rs!Gastos, "N") + ImpAux
                ItmX.SubItems(8) = Format(ImpAux, FormatoImporte)
            End If
            
        End If
    End If
    If Rs!FecVenci < Fecha Then
        'LO DEBE
        If impo <> 0 Then ItmX.SmallIcon = 1
        Vencido = Vencido + impo
    Else
'        ItmX.SmallIcon = 2
    End If
    Importe = Importe + impo
    
    ItmX.Tag = Rs!codmacta
    
    If Tipo = 1 And SegundoParametro <> "" Then
        If Not IsNull(Rs!transfer) Then
            ItmX.Checked = True
            ImpSeleccionado = ImpSeleccionado + impo
        End If
    End If

End Sub



Private Function DevSQL() As String
Dim Cad As String

    If Not Cobros Then
        Cad = "SELECT pagos.*, pagos.nomprove nommacta, tipofpago.siglas,pagos.codmacta,ImpEfect-coalesce(imppagad,0) as imppdte  FROM"
        Cad = Cad & " pagos, formapago, tipofpago"
        Cad = Cad & " Where formapago.tipforpa = tipofpago.tipoformapago"
        Cad = Cad & " AND pagos.codforpa = formapago.codforpa"
        If vSql <> "" Then Cad = Cad & " AND " & vSql
    
    Else
        'cobros
        Cad = "SELECT cobros.*, formapago.nomforpa, tipofpago.descformapago, tipofpago.siglas, "
        Cad = Cad & " cobros.nomclien nommacta,cobros.codmacta,tipofpago.tipoformapago, "
        Cad = Cad & " coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0) imppdte "
        Cad = Cad & " FROM (cobros INNER JOIN formapago ON cobros.codforpa = formapago.codforpa) INNER JOIN tipofpago ON formapago.tipforpa = tipofpago.tipoformapago "
        If vSql <> "" Then Cad = Cad & " WHERE " & vSql
    End If
    'SQL pedido
    DevSQL = Cad
End Function


Private Sub CargaPagos()

    Cad = DevSQL
    
    'ORDENACION
    If CampoOrden = "" Then CampoOrden = "pagos.fecefect"
    Cad = Cad & " ORDER BY " & CampoOrden
    If Orden Then Cad = Cad & " DESC"
    If CampoOrden <> "pagos.fecefect" Then Cad = Cad & ", pagos.fecefect"


    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        InsertaItemPago
        Rs.MoveNext
    Wend
    Rs.Close

End Sub


Private Sub InsertaItemPago()
Dim J As Byte
    
    Set ItmX = ListView1.ListItems.Add()
    
    ItmX.Text = Rs!NUmSerie
    ItmX.SubItems(1) = Rs!NumFactu
    ItmX.SubItems(2) = Format(Rs!FecFactu, "dd/mm/yyyy")
    ItmX.SubItems(3) = Format(Rs!fecefect, "dd/mm/yyyy")
    ItmX.SubItems(4) = Rs!numorden
    ItmX.SubItems(5) = DBLet(Rs!Nommacta, "T")
    ItmX.SubItems(6) = DBLet(Rs!siglas, "T")
    ItmX.SubItems(7) = Format(Rs!ImpEfect, FormatoImporte)
    If Not IsNull(Rs!imppagad) Then
        ItmX.SubItems(8) = Format(Rs!imppagad, FormatoImporte)
        impo = Rs!ImpEfect - Rs!imppagad
        ItmX.SubItems(9) = Format(impo, FormatoImporte)
    Else
        impo = Rs!ImpEfect
        ItmX.SubItems(8) = "0.00"
        ItmX.SubItems(9) = ItmX.SubItems(7)
    End If
    If Rs!fecefect < Fecha Then
        'LO DEBE
        If impo <> 0 Then
            ItmX.SmallIcon = 1
            ItmX.ToolTipText = "Pendiente"
        End If
        Vencido = Vencido + impo
        
    Else
'        ItmX.SmallIcon = 2
    End If
    
    If Tipo = 1 Then
        If Not IsNull(Rs!nrodocum) Then
            ItmX.Checked = True
            ImpSeleccionado = ImpSeleccionado + impo
        End If
    End If
    'El tag lo utilizo para la cta proveedor
    ItmX.Tag = Rs!codmacta
    
    Importe = Importe + impo
    
    'Si el documento estaba emitido ya
    If Val(Rs!emitdocum) = 1 Then
        'Tiene marcado DOCUMENTO EMITIDO
        ItmX.ForeColor = vbRed
        For J = 1 To ListView1.ColumnHeaders.Count - 1
            ItmX.ListSubItems(J).ForeColor = vbRed
        Next J
        'If DBLet(Rs!Referencia, "T") = "" Then ItmX.ListSubItems(4).ForeColor = vbMagenta
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Para dejar las variables bien
    ContabTransfer = False
    DesdeRecepcionTalones = False
    'Por si acaso
    NumeroTalonPagere = ""
    CodmactaUnica = ""
    
    
    
    LeerFiltroOrdenacion False
    
End Sub


Private Sub frmC_Selec(vFecha As Date)
    Cad = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub imgFecha_Click(Index As Integer)
    Fecha = Now
    Select Case Index
    Case 1
        If Text1.Text <> "" Then
            If IsDate(Text1.Text) Then Fecha = CDate(Text1.Text)
        End If
    End Select
    Cad = ""
    Set frmC = New frmCal
    frmC.Fecha = Fecha
    frmC.Show vbModal
    Set frmC = Nothing
    If Cad <> "" Then
        Select Case Index
        Case 1
            Text1.Text = Cad
        End Select
    End If
End Sub





Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Campo2 As Integer

    

    Orden = Not Orden
    If Cobros Then
'        Columnas = "Serie|Nº Factura|F.Factura|F. VTO|Nº|CLIENTE|Tipo|Importe|Gasto|Cobrado|Pendiente|"
        Select Case ColumnHeader
            Case "Serie"
                CampoOrden = "cobros.numserie " & IIf(Orden, "DESC", "") & ",cobros.numfactu"
            Case "Factura"
                CampoOrden = "cobros.numfactu " & IIf(Orden, "DESC", "") & ",cobros.fecfactu"
            Case "F.Factura"
                CampoOrden = "cobros.fecfactu " & IIf(Orden, "DESC", "") & ",cobros.numserie,cobros.numfactu"
            Case "F. VTO"
                CampoOrden = "cobros.fecvenci " & IIf(Orden, "DESC", "") & ",cobros.fecfactu,cobros.numfactu"
            Case "Nº"
                CampoOrden = "cobros.numorden " & IIf(Orden, "DESC", "") & ""
            Case "CLIENTE"
                CampoOrden = "nommacta " & IIf(Orden, "DESC", "") & ",cobros.fecfactu,cobros.numfactu"
            Case "Tipo"
                CampoOrden = "siglas " & IIf(Orden, "DESC", "") & ",cobros.fecfactu,cobros.numfactu"
            Case "Importe"
                CampoOrden = "cobros.impvenci " & IIf(Orden, "DESC", "") & " ,cobros.fecfactu,cobros.numfactu"
            Case "Gasto"
                CampoOrden = "cobros.gastos " & IIf(Orden, "DESC", "") & " ,cobros.fecfactu,cobros.numfactu"
            Case "Cobrado"
                CampoOrden = "cobros.impcobro " & IIf(Orden, "DESC", "") & " ,cobros.fecfactu,cobros.numfactu"
            Case "Pendiente"
                CampoOrden = "imppdte " & IIf(Orden, "DESC", "") & " ,cobros.fecfactu,cobros.numfactu"
        End Select
        CargaList
    Else
        Select Case ColumnHeader
            Case "Serie"
                CampoOrden = "pagos.numserie"
            Case "Nº Factura"
                CampoOrden = "pagos.numfactu"
            Case "F. Fact"
                CampoOrden = "pagos.fecfactu"
            Case "F. VTO"
                CampoOrden = "pagos.fecefect"
            Case "Nº"
                CampoOrden = "pagos.numorden"
            Case "PROVEEDOR"
                CampoOrden = "pagos.nomprove"
            Case "Tipo"
                CampoOrden = "siglas"
            Case "Importe"
                CampoOrden = "pagos.impefect"
            Case "Pagado"
                CampoOrden = "pagos.imppagad"
            Case "Pendiente"
                CampoOrden = "imppdte"
        End Select
        CargaList
    
    End If
    
End Sub

Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    If Regresar Then
        cmdRegresar_Click
    Else
    
    End If
    
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    i = ColD(0)
    impo = ImporteFormateado(Item.SubItems(i))
    
    If Item.Checked Then
        Set ListView1.SelectedItem = Item
        i = 1
    Else
        i = -1
    End If
    ImpSeleccionado = ImpSeleccionado + (i * impo)
    Text2(2).Text = Format(ImpSeleccionado, FormatoImporte)
    
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu Me.mnContextual
    End If
End Sub

Private Sub SeleccionarTodos(Seleccionar As Boolean)
Dim J As Integer
    J = ColD(0)
    ImpSeleccionado = 0
    For i = 1 To Me.ListView1.ListItems.Count
        ListView1.ListItems(i).Checked = Seleccionar
        impo = ImporteFormateado(ListView1.ListItems(i).SubItems(J))
        ImpSeleccionado = ImpSeleccionado + impo
    Next i
    If Not Seleccionar Then ImpSeleccionado = 0
    Text2(2).Text = Format(ImpSeleccionado, FormatoImporte)
End Sub

Private Sub mnNumero_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
        
End Sub

Private Sub mnQUitarSel_Click()
    SeleccionarTodos False
End Sub

Private Sub mnSelectAll_Click()
    SeleccionarTodos True
End Sub

Private Sub Text1_GotFocus()
    ConseguirFoco Text1, 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus()
    If Not EsFechaOK(Text1) Then
        MsgBox "Fecha incorrecta", vbExclamation
        Text1.Text = ""
        Text1.SetFocus
    Else
        Screen.MousePointer = vbHourglass
        CargaList
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



' A partir de un numero de columna nos dira k columna es
' en el LISTVIEW
'
Private Function ColD(Colu As Integer) As Integer
    Select Case Colu
    Case 0
            'IMporte pendiente
            ColD = 10
    Case 1
    
    End Select
    If Not Cobros Then ColD = ColD - 2
End Function




Private Sub LeerFiltroOrdenacion(Leer As Boolean)
Dim NF As Integer
Dim C As String
    On Error GoTo eLeerFiltroOrdenacion
    
    
    
    
    NF = FreeFile
    If Cobros Then
        
        
        
        
        C = App.Path & "\OrdenCob.xdf"
        If Leer Then
            CampoOrden = "cobros.fecvenci"
        
            If Dir(C, vbArchive) <> "" Then
                Open C For Input As #NF
                Line Input #NF, C
                Close #NF
                CampoOrden = C
               
            End If
            OrdenEnFicheroXDF = CampoOrden
        Else
            
            If OrdenEnFicheroXDF <> CampoOrden Then
                If CampoOrden = "cobros.fecvenci" Then
                    If Dir(C, vbArchive) <> "" Then Kill C
                Else
                    Open C For Output As #NF
                    Print #NF, CampoOrden
                    Close #NF
                End If
            
            End If
        End If
        
    Else
        C = App.Path & "\Ordenpag.xdf"
     
       
        If Leer Then
            CampoOrden = "pagos.fecefect"
        
            If Dir(C, vbArchive) <> "" Then
                Open C For Input As #NF
                Line Input #NF, C
                Close #NF
                CampoOrden = C
                
            End If
            
        Else
            
            If OrdenEnFicheroXDF <> CampoOrden Then
                If CampoOrden = "pagos.fecefect" Then
                    If Dir(C, vbArchive) <> "" Then Kill C
                Else
                    Open C For Output As #NF
                    Print #NF, CampoOrden
                    Close #NF
                End If
            
            End If
        End If
    
    End If
    
    OrdenEnFicheroXDF = CampoOrden
    
    Exit Sub
eLeerFiltroOrdenacion:
    Err.Clear
End Sub
