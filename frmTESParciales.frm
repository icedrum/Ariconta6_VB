VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESParciales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anticipo vto."
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkImprimir 
      Caption         =   "Imprimir Recibo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   19
      Top             =   7020
      Width           =   2685
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   7140
      TabIndex        =   6
      Top             =   7140
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
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
      Index           =   0
      Left            =   5820
      TabIndex        =   5
      Top             =   7140
      Width           =   1095
   End
   Begin VB.Frame FrCobro 
      Height          =   6855
      Left            =   60
      TabIndex        =   7
      Top             =   90
      Width           =   8175
      Begin VB.TextBox Text2 
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
         Left            =   6000
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   4290
         Width           =   1755
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmTESParciales.frx":0000
         Left            =   1590
         List            =   "frmTESParciales.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Tipo de pago|N|N|||formapago|tipforpa|||"
         Top             =   4260
         Width           =   2475
      End
      Begin VB.TextBox txtCta 
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
         Left            =   1680
         TabIndex        =   0
         Text            =   "Text2"
         Top             =   1470
         Width           =   1215
      End
      Begin VB.TextBox txtDescCta 
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
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   1470
         Width           =   4785
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Index           =   5
         Left            =   6030
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   2940
         Width           =   1755
      End
      Begin VB.TextBox Text2 
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
         Left            =   6000
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   3855
         Width           =   1755
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Index           =   4
         Left            =   6030
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2490
         Width           =   1755
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Index           =   3
         Left            =   6030
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2010
         Width           =   1755
      End
      Begin VB.TextBox Text3 
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
         Left            =   1590
         TabIndex        =   1
         Top             =   3825
         Width           =   1305
      End
      Begin MSComctlLib.ListView ListView8 
         Height          =   1455
         Left            =   240
         TabIndex        =   27
         Top             =   5130
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   2566
         View            =   3
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "C�digo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gasto Bancario"
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
         Height          =   240
         Index           =   10
         Left            =   4350
         TabIndex        =   30
         Top             =   4335
         Width           =   1485
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   9
         Left            =   240
         TabIndex        =   29
         Top             =   4290
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cobros realizados: "
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
         Height          =   240
         Index           =   8
         Left            =   240
         TabIndex        =   28
         Top             =   4830
         Width           =   1920
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   56
         Left            =   270
         TabIndex        =   21
         Top             =   360
         Width           =   6150
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
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
         Height          =   240
         Index           =   57
         Left            =   270
         TabIndex        =   20
         Top             =   720
         Width           =   6270
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   7860
         Y1              =   4710
         Y2              =   4710
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
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
         Left            =   240
         TabIndex        =   18
         Top             =   3330
         Width           =   6195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cta banco"
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
         Height          =   240
         Index           =   7
         Left            =   270
         TabIndex        =   17
         Top             =   1470
         Width           =   1050
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   1380
         Picture         =   "frmTESParciales.frx":0004
         Top             =   1530
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
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
         Height          =   240
         Index           =   6
         Left            =   4380
         TabIndex        =   15
         Top             =   3900
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pagado"
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
         Height          =   240
         Index           =   5
         Left            =   5100
         TabIndex        =   14
         Top             =   2940
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   7860
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gastos"
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
         Height          =   240
         Index           =   4
         Left            =   5160
         TabIndex        =   11
         Top             =   2550
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe TOTAL"
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
         Height          =   240
         Index           =   2
         Left            =   4380
         TabIndex        =   9
         Top             =   2100
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   3840
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   1230
         Picture         =   "frmTESParciales.frx":6856
         Top             =   3870
         Width           =   240
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
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
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento"
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
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   26
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
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
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   25
      Top             =   1080
      Width           =   675
   End
End
Attribute VB_Name = "frmTESParciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Public Cobro As Boolean
Public Vto As String  'Llevara empipado las claves
Public Cta As String
Public Importes As String 'Empipado los importes
Public FormaPago As Integer

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1

Dim impo As Currency
Dim cad As String
Dim PrimeraVez As Boolean
Dim TipForpa As Integer

Dim LineaCobro As Long

Private Sub PonFoco(ByRef T1 As TextBox)
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub ImprimirRecibo()

    If ImporteFormateado(Text2(0).Text) <= 0 Then
        MsgBox "No se pueden emitir recibos por importes menores o iguales a cero", vbExclamation
        Exit Sub
    End If
    
    CargarTemporal
    
    frmTESImpRecibo.Show vbModal
    
End Sub

Private Sub CargarTemporal()
Dim SQL As String

    SQL = "delete from tmppendientes where codusu = " & vUsu.Codigo
    Conn.Execute SQL

    ' en tmppendientes metemos la clave primaria de cobros_recibidos y el importe en letra
                                                      'importe=nro factura,   codforpa=linea de cobros_realizados
    SQL = "insert into tmppendientes (codusu,serie_cta,importe,fecha,numorden,codforpa, observa) values ("
    SQL = SQL & vUsu.Codigo & "," & DBSet(RecuperaValor(Vto, 1), "T") & "," 'numserie
    SQL = SQL & DBSet(RecuperaValor(Vto, 2), "N") & "," 'numfactu
    SQL = SQL & DBSet(RecuperaValor(Vto, 3), "F") & "," 'fecfactu
    SQL = SQL & DBSet(RecuperaValor(Vto, 4), "N") & "," 'numorden
    SQL = SQL & DBSet(LineaCobro, "N") & "," 'numlinea
    SQL = SQL & DBSet(EscribeImporteLetra(ImporteFormateado(Text2(0).Text)), "T") & ") "
    
    Conn.Execute SQL

End Sub


Private Sub Command1_Click(Index As Integer)
Dim B As Boolean
    
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        'Comprobamos importes. Y fecha de contabilizacioon
        If Not DatosOK Then Exit Sub
        
        If Cobro Then
            CadenaDesdeOtroForm = "cobro"
        Else
            CadenaDesdeOtroForm = "pago"
        End If
        CadenaDesdeOtroForm = "Desea generar el " & CadenaDesdeOtroForm & "?"
        B = True
        If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbNo Then B = False
        CadenaDesdeOtroForm = ""
        If Not B Then Exit Sub
        
        'UPDATEAMOS EL Vencimiento y CONTABILIZAMOS EL COBRO/PAGO
        Screen.MousePointer = vbHourglass
        B = RealizarAnticipo
        Screen.MousePointer = vbDefault
        If Not B Then Exit Sub
        CadenaDesdeOtroForm = "OK" 'Para que refresque los datos en el form
        
        If chkImprimir.Value = 1 Then ImprimirRecibo
    End If
    
    Unload Me
End Sub

Private Sub CargarListView()
Dim RS As ADODB.Recordset
Dim IT As ListItem
    
    On Error GoTo ECargarlistview
    
    
    ListView8.ColumnHeaders.Clear
    ListView8.ListItems.Clear
    
    
    ListView8.ColumnHeaders.Add , , "Fecha", 1400.2522
    ListView8.ColumnHeaders.Add , , "Usuario", 2000.2522
    ListView8.ColumnHeaders.Add , , "Tipo", 900.2522
    ListView8.ColumnHeaders.Add , , "Importe", 1700.2522, 1
    
    Set RS = New ADODB.Recordset
    
    If Cobro Then
        cad = "select hlinapu.fechaent, hcabapu.usucreacion, tipofpago.siglas,  coalesce(timporteh,0) - coalesce(timported,0) impcobro "
        cad = cad & " from (hlinapu inner join tipofpago on hlinapu.tipforpa = tipofpago.tipoformapago) "
        cad = cad & " inner join hcabapu on hlinapu.numdiari = hcabapu.numdiari and hlinapu.fechaent = hcabapu.fechaent and hlinapu.numasien = hcabapu.numasien "
        cad = cad & " where numserie = " & DBSet(RecuperaValor(Vto, 1), "T")
        cad = cad & " and numfaccl = " & DBSet(RecuperaValor(Vto, 2), "N")
        cad = cad & " and fecfactu = " & DBSet(RecuperaValor(Vto, 3), "F")
        cad = cad & " and numorden = " & DBSet(RecuperaValor(Vto, 4), "N")
        cad = cad & " order by fechaent  "
    Else
        cad = "select hlinapu.fechaent, hcabapu.usucreacion, tipofpago.siglas, coalesce(timporteh,0) - coalesce(timported,0) imppago "
        cad = cad & " from (hlinapu inner join tipofpago on hlinapu.tipforpa = tipofpago.tipoformapago) "
        cad = cad & " inner join hcabapu on hlinapu.numdiari = hcabapu.numdiari and hlinapu.fechaent = hcabapu.fechaent and hlinapu.numasien = hcabapu.numasien "
        cad = cad & " where numserie = " & DBSet(RecuperaValor(Vto, 1), "T")
        cad = cad & " and numfacpr = " & DBSet(RecuperaValor(Vto, 2), "T")
        cad = cad & " and fecfactu = " & DBSet(RecuperaValor(Vto, 3), "F")
        cad = cad & " and numorden = " & DBSet(RecuperaValor(Vto, 4), "N")
        cad = cad & " and codmacta = " & DBSet(RecuperaValor(Cta, 1), "T")
        cad = cad & " order by fechaent "
    End If
    
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not RS.EOF
                    
        Set IT = ListView8.ListItems.Add
        
        IT.Text = DBLet(RS.Fields(0))
        IT.SubItems(1) = DBLet(RS.Fields(1))
        IT.SubItems(2) = DBLet(RS.Fields(2))
        IT.SubItems(3) = Format(DBLet(RS.Fields(3)), "###,###,##0.00")
        
        'Siguiente
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ECargarlistview:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
            
        CargarListView
    
        PosicionarCombo Combo1, FormaPago
    
    End If
        
End Sub

Private Sub Form_Load()
        
    Me.Icon = frmPpal.Icon
        
    PrimeraVez = True
        
    If Cobro Then
        Caption = "Cobro"
        Text1(0).Text = RecuperaValor(Vto, 1) & "/" & RecuperaValor(Vto, 2) & "   Fecha: " & RecuperaValor(Vto, 3) & "   Vto. num: " & RecuperaValor(Vto, 4)
        Text1(1).Text = RecuperaValor(Cta, 1)
        Text1(2).Text = RecuperaValor(Cta, 2)
        'Dos
        txtCta(1).Text = RecuperaValor(Cta, 3)
        Me.txtDescCta(1).Text = RecuperaValor(Cta, 4)
        
        'Importes
        Text1(3).Text = RecuperaValor(Importes, 1)
        Text1(4).Text = RecuperaValor(Importes, 2)
        Text1(5).Text = RecuperaValor(Importes, 3)
        Text3(0).Text = Format(Now, "dd/mm/yyyy")
        Label4(4).Caption = "Gastos"
        Label4(1).Caption = "Cliente"
                
        Label4(57).Caption = Text1(0).Text
        Label4(56).Caption = Text1(2)
        
        Label4(8).Caption = "Cobros Realizados:"
    Else
        'PAGO
        Label4(1).Caption = "Proveedor"
        Caption = "Pago"
        
        'Cobro parcial de vencimientos
        Text1(0).Text = RecuperaValor(Vto, 1) & "/" & RecuperaValor(Vto, 2) & "   Fecha: " & RecuperaValor(Vto, 3) & "   Vto. num: " & RecuperaValor(Vto, 4)
        Text1(1).Text = RecuperaValor(Cta, 1)
        Text1(2).Text = RecuperaValor(Cta, 2)
        'Dos
        txtCta(1).Text = RecuperaValor(Cta, 3)
        Me.txtDescCta(1).Text = RecuperaValor(Cta, 4)
        
        Text1(3).Text = RecuperaValor(Importes, 1)
        Text1(4).Text = ""
        Text1(5).Text = RecuperaValor(Importes, 2)  'Esto es lo pagado ya
        Text3(0).Text = Format(Now, "dd/mm/yyyy")
        
    
        Label4(57).Caption = Text1(0).Text
        Label4(56).Caption = Text1(2)
    
        chkImprimir.Enabled = False
        chkImprimir.Visible = False
        chkImprimir.Value = 0
    
        Label4(8).Caption = "Pagos Realizados:"
    End If
    
    
    'IMPORTE Restante
    
    impo = ImporteFormateado(Text1(3).Text) 'Vto
    If Cobro Then
        'Gastos
        If Text1(4).Text <> "" Then impo = impo + ImporteFormateado(Text1(4).Text)
            
        'Ya cobrado
        If Text1(5).Text <> "" Then impo = impo - ImporteFormateado(Text1(5).Text)
        
    Else
        'Gastos
        If Text1(4).Text <> "" Then impo = impo + ImporteFormateado(Text1(4).Text)
        
        'Ya cobrado
        If Text1(5).Text <> "" Then impo = impo - ImporteFormateado(Text1(5).Text)
            
    End If
    Label1.Caption = "Pendiente: " & Format(impo, FormatoImporte)
    
    CargaCombo
    
    Label4(4).Visible = Cobro
    Text1(4).Visible = Cobro
    Me.Height = Me.FrCobro.Height + 1200 '240 + Me.Command1(0).Height + 240
    Text2(0).Text = Format(impo, FormatoImporte)
    Text2(1).Text = "0,00"
    
    Caption = Caption & " de factura"
End Sub

Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtCta(CInt(imgCuentas(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescCta(CInt(imgCuentas(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(Text3(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image2_Click(Index As Integer)
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text3(Index).Text <> "" Then frmC.Fecha = CDate(Text3(Index).Text)
    Text3(0).Tag = Index
    frmC.Show vbModal
    Set frmC = Nothing
End Sub



Private Sub imgCuentas_Click(Index As Integer)
    imgCuentas(1).Tag = Index
    Set frmBa = New frmBanco
    frmBa.DatosADevolverBusqueda = "OK"
    frmBa.Show vbModal
    Set frmBa = Nothing
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    PonFoco Text2(Index)
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus(Index As Integer)
Dim Valor

    If Text2(Index).Text = "" Then Exit Sub
    If Not IsNumeric(Text2(Index).Text) Then
        MsgBox "Importe debe ser num�rico", vbExclamation
        Text2(Index).Text = ""
        PonFoco Text2(Index)
    Else
        If InStr(1, Text2(Index).Text, ",") > 0 Then
            Valor = ImporteFormateado(Text2(Index).Text)
        Else
            Valor = CCur(TransformaPuntosComas(Text2(Index).Text))
        End If
        Text2(Index).Text = Format(Valor, FormatoImporte)
    End If
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    PonFoco Text3(Index)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index))
    If Text3(Index) = "" Then Exit Sub
    If Not EsFechaOK(Text3(Index)) Then
        MsgBox "Fecha incorrecta: " & Text3(Index), vbExclamation
        Text3(Index).Text = ""
        Text3(Index).SetFocus
    End If
End Sub


Private Function DatosOK() As Boolean
Dim Im As Currency
Dim CtaBancoGastos As String


    On Error GoTo EDa
    DatosOK = False
    
    
    cad = ""
    If Text2(0).Text = "" Then cad = "importe"
    If Text3(0).Text = "" Then cad = cad & " fecha"
    If cad <> "" Then
        MsgBox "Falta: " & cad, vbExclamation
        Exit Function
    End If
    
    '----------------------------------
    'Junio 2011
    'YA dejamos cobros negativos
    Im = ImporteFormateado(Text2(0).Text)
    'If Im < 0 Then
    If Im = 0 Then
        MsgBox "importes CERO", vbExclamation
        Exit Function
    End If
    
    
    If txtCta(1).Text = "" Then
        MsgBox "Falta cuenta banco", vbExclamation
        Exit Function
    End If
        
    'Fecha dentro ejercicios
    If FechaCorrecta2(CDate(Text3(0).Text), True) > 1 Then Exit Function
    
    If ComprobarCero(Text2(1).Text) <> 0 Then
        CtaBancoGastos = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", txtCta(1), "T")
        If CtaBancoGastos = "" Then
            CtaBancoGastos = DevuelveDesdeBD("ctabenbanc", "paramtesor", "codigo", "1", "N")
        End If
        If CtaBancoGastos = "" Then
            MsgBox "Falta configurar la cuenta de gastos bancarios. Revise.", vbExclamation
            Exit Function
        End If
    End If
    
    If Cobro Then
        impo = ImporteFormateado(Text1(3).Text) 'Vto
        'Gastos
        If Text1(4).Text <> "" Then
            Im = ImporteFormateado(Text1(4).Text)
            impo = impo + Im
        End If
        
        'Ya cobrado
        If Text1(5).Text <> "" Then
            Im = ImporteFormateado(Text1(5).Text)
            impo = impo - Im
        End If
    Else
        impo = ImporteFormateado(Text1(3).Text) 'Vto

        'Ya cobrado
        If Text1(5).Text <> "" Then
            Im = ImporteFormateado(Text1(5).Text)
            impo = impo - Im
        End If
    End If
    
    Im = ImporteFormateado(Text2(0).Text) 'Lo que voy a pagar
    cad = ""
    If impo < 0 Then
        'Importes negativos
        If Im >= 0 Then
            cad = "negativo"
        Else
            If Im < impo Then cad = "X"
        End If
    Else
        If Im <= 0 Then
            cad = "positivo"
        Else
            If Im > impo Then cad = "X"
        End If
    End If
        
    If cad <> "" Then
        
        If cad = "X" Then
            cad = "Importe a pagar mayor que el importe restante.(" & Format(Im, FormatoImporte) & " : " & Format(impo, FormatoImporte) & ")"
        Else
            cad = "El importe debe ser " & cad
        End If
        MsgBox cad, vbExclamation
        Exit Function
    End If
        
    'Comprobaremos un par de cosillas
    If CuentaBloqeada(RecuperaValor(Cta, 1), CDate(Text3(0).Text), True) Then Exit Function
        
    DatosOK = True
    Exit Function
EDa:
    MuestraError Err.Number, "Datos Ok"
End Function


Private Function RealizarAnticipo() As Boolean

    Conn.BeginTrans
    If Contabilizar Then
        Conn.CommitTrans
        RealizarAnticipo = True
    Else
        'Conn.RollbackTrans
        TirarAtrasTransaccion
        RealizarAnticipo = False
    End If

End Function


Private Function Contabilizar() As Boolean
Dim Mc As Contadores
Dim FP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim Numdocum As String
Dim Conce As Integer
Dim LlevaContr As Boolean
Dim Im As Currency
Dim Debe As Boolean
Dim ElConcepto As Integer
Dim vNumDiari As Integer
Dim Situacion As Integer

Dim Gastos As Currency
Dim Importe As Currency
Dim CtaBancoGastos As String
Dim DescuentaImporteDevolucion As Boolean
Dim Sql5 As String


    On Error GoTo ECon
    Contabilizar = False
    
    
    
    Set Mc = New Contadores
    If Mc.ConseguirContador("0", CDate(Text3(0).Text) <= vParam.fechafin, True) = 1 Then Exit Function

    Set FP = New Ctipoformapago
    If FP.Leer(Combo1.ListIndex) Then  ' antes forma de pago
        Set Mc = Nothing
        Set FP = Nothing
    End If
    
    
    'importe
    impo = ImporteFormateado(Text2(0).Text)
    
    'Inserto cabecera de apunte
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
    If Cobro Then
        SQL = SQL & FP.diaricli
        vNumDiari = FP.diaricli
    Else
        SQL = SQL & FP.diaripro
        vNumDiari = FP.diaripro
    End If
    SQL = SQL & ",'" & Format(Text3(0).Text, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ",'"
    SQL = SQL & "Generado desde Tesorer�a el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & DevNombreSQL(vUsu.Nombre)
    If impo < 0 Then SQL = SQL & "  (ABONO)"
    SQL = SQL & "',"
    If Cobro Then
        SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizar Cobros')"
    Else
        SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizar Pagos')"
    End If
    
    Conn.Execute SQL
        
        
    'Inserto en las lineas de apuntes
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada,"
    
    'campos a�adidos en hlinapu
    If Cobro Then
        SQL = SQL & "numserie,numfaccl,fecfactu,numorden,tipforpa,reftalonpag,bancotalonpag) VALUES ("
    Else
        SQL = SQL & "numserie,numfacpr,fecfactu,numorden,tipforpa,reftalonpag,bancotalonpag) VALUES ("
    End If
    
    If Cobro Then
        SQL = SQL & FP.diaricli
    Else
        SQL = SQL & FP.diaripro
    End If
    SQL = SQL & ",'" & Format(Text3(0).Text, FormatoFecha) & "'," & Mc.Contador & ","
    
    
    'numdocum
    Numdocum = DevNombreSQL(RecuperaValor(Vto, 2))
    If Cobro Then
        Numdocum = RecuperaValor(Vto, 1) & Format(Numdocum, "0000000")
    End If
    
    
    'Concepto y ampliacion del apunte
    Ampliacion = ""
    If Cobro Then
        'CLIENTES
        Debe = False
        If impo < 0 Then
            If Not vParam.abononeg Then Debe = True
        End If
        If Debe Then
            Conce = FP.ampdecli
            LlevaContr = FP.ctrdecli = 1
            ElConcepto = FP.condecli
        Else
            ElConcepto = FP.conhacli
            Conce = FP.amphacli
            LlevaContr = FP.ctrhacli = 1
        End If
    Else
        'PAGOS
        Debe = True
        If impo < 0 Then
            If Not vParam.abononeg Then Debe = False
        End If
        If Debe Then
            Conce = FP.ampdepro
            LlevaContr = FP.ctrdepro = 1
            ElConcepto = FP.condepro
        Else
            ElConcepto = FP.conhapro
            Conce = FP.amphapro
            LlevaContr = FP.ctrhapro = 1
        End If

    End If
           
    'Si el importe es negativo y no permite abonos negativos
    'como ya lo ha cambiado de lado (dbe <-> haber)
    If impo < 0 Then
        If Not vParam.abononeg Then impo = Abs(impo)
    End If
       
           
    If Conce = 2 Then
       Ampliacion = Ampliacion & RecuperaValor(Vto, 3)  'Fecha vto
    ElseIf Conce = 4 Then
        'Contra partida
        Ampliacion = DevNombreSQL(txtDescCta(1).Text)
    Else
        
       If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
       If Cobro Then
            Ampliacion = Ampliacion & RecuperaValor(Vto, 1) & Format(RecuperaValor(Vto, 2), "0000000") '& "/" & Mid(RecuperaValor(Vto, 2), 1, 9)
       Else
            Ampliacion = Ampliacion & Mid(RecuperaValor(Vto, 2), 1, 9)
       End If
    End If
    
    'Fijo en concepto el codconce
    Conce = ElConcepto
    cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
    Ampliacion = cad & " " & Ampliacion
    Ampliacion = Mid(Ampliacion, 1, 35)
    
    
    
    
    'Ahora ponemos linliapu codmacta numdocum codconce ampconce timported timporte codccost ctacontr idcontab punteada
    'Cuenta Cliente/proveedor
    cad = "1,'" & Text1(1).Text & "','" & Numdocum & "'," & Conce & ",'" & DevNombreSQL(Ampliacion) & "',"
    'Importe cobro-pago
    ' nos lo dire "debe"
    If Not Debe Then
        cad = cad & "NULL," & TransformaComasPuntos(CStr(impo))
    Else
        cad = cad & TransformaComasPuntos(CStr(impo)) & ",NULL"
    End If
    'Codccost
    cad = cad & ",NULL,"
    If LlevaContr Then
        cad = cad & "'" & txtCta(1).Text & "'"
    Else
        cad = cad & "NULL"
    End If
    If Cobro Then
        cad = cad & ",'COBROS',0,"
        cad = cad & DBSet(RecuperaValor(Vto, 1), "T") & "," '& RecuperaValor(Vto, 2) & ","
    Else
        cad = cad & ",'PAGOS',0,"
        cad = cad & DBSet(RecuperaValor(Vto, 1), "T") & ","
    End If
    
    cad = cad & DBSet(RecuperaValor(Vto, 2), "T") & "," & DBSet(RecuperaValor(Vto, 3), "F") & ","
    cad = cad & DBSet(RecuperaValor(Vto, 4), "N") & "," & DBSet(Combo1.ItemData(Combo1.ListIndex), "N") & "," & ValorNulo & "," & ValorNulo & ")"
    
    cad = SQL & cad
    Conn.Execute cad
    
       
    'El banco    *******************************************************************************
    '---------------------------------------------------------------------------------------------
    
    'Vuelvo a fijar los valores
     'Concepto y ampliacion del apunte
    Ampliacion = ""
    If Cobro Then
       'CLIENTES
        'Si el apunte va al debe, el contrapunte va al haber
        If Not Debe Then
            Conce = FP.ampdecli
            LlevaContr = FP.ctrdecli = 1
            ElConcepto = FP.condecli
        Else
            ElConcepto = FP.conhacli
            Conce = FP.amphacli
            LlevaContr = FP.ctrhacli = 1
        End If
    Else
        'PAGOS
        'Si el apunte va al debe, el contrapunte va al haber
        If Not Debe Then
            Conce = FP.ampdepro
            LlevaContr = FP.ctrdepro = 1
            ElConcepto = FP.condepro
        Else
            ElConcepto = FP.conhapro
            Conce = FP.amphapro
            LlevaContr = FP.ctrhapro = 1
        End If
    End If
           
           
    If Conce = 2 Then
       Ampliacion = Ampliacion & RecuperaValor(Vto, 3)  'Fecha vto
    ElseIf Conce = 4 Then
        'Contra partida
        Ampliacion = DevNombreSQL(Text1(2).Text)
    Else
        If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
        If Cobro Then
             Ampliacion = Ampliacion & RecuperaValor(Vto, 1) & Format(RecuperaValor(Vto, 2), "0000000") ' "/" & Mid(RecuperaValor(Vto, 2), 1, 9)
        Else
             Ampliacion = Ampliacion & Mid(RecuperaValor(Vto, 2), 1, 9)
        End If
    End If
    
    
    Conce = ElConcepto
    cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
    Ampliacion = cad & " " & Ampliacion
    Ampliacion = Mid(Ampliacion, 1, 35)
    
    
    Gastos = 0
    If ComprobarCero(Text2(1).Text) <> 0 Then
        Gastos = ImporteFormateado(Text2(1).Text)
    End If
    
    DescuentaImporteDevolucion = False
    If Gastos > 0 Then
        Sql5 = txtCta(1)
        If Cobro Then
            Sql5 = DevuelveDesdeBD("GastRemDescontad", "bancos", "codmacta", Sql5, "T")
        Else
            Sql5 = DevuelveDesdeBD("GasttransDescontad", "bancos", "codmacta", Sql5, "T")
        End If
        If Sql5 = "1" Then DescuentaImporteDevolucion = True
    End If
    Importe = impo
    If DescuentaImporteDevolucion Then
        Importe = impo - Gastos
    End If
    
    cad = "2,'" & txtCta(1).Text & "','" & Numdocum & "'," & Conce & ",'" & Ampliacion & "',"
    'Importe cliente
    'Si el cobro/pago va al debe el contrapunte ira al haber
    If Not Debe Then
        'al debe
        cad = cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
    Else
        'al haber
        cad = cad & "NULL," & TransformaComasPuntos(CStr(Importe))
    End If
    
    'Codccost
    cad = cad & ",NULL,"
    
    If LlevaContr Then
        cad = cad & "'" & Text1(1).Text & "'"
    Else
        cad = cad & "NULL"
    End If
    If Cobro Then
        cad = cad & ",'COBROS',0," ' idcontab
    Else
        cad = cad & ",'PAGOS',0," ' idcontab
    End If
    
    ' todo valores a null ????
    cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
    
    cad = SQL & cad
    Conn.Execute cad
    
        
    '++
    'Gasto.
    ' Si tiene y no agrupa
    '-------------------------------------------------------
    If Gastos > 0 Then
        If CtaBancoGastos = "" Then CtaBancoGastos = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", txtCta(1), "T")
        If CtaBancoGastos = "" Then
            CtaBancoGastos = DevuelveDesdeBD("ctabenbanc", "paramtesor", "codigo", "1", "N")
        End If

        cad = "3,'"

        cad = cad & CtaBancoGastos & "','" & Numdocum & "'," & Conce
        cad = cad & ",'Gastos vto.'"

        'Importe al debe
        cad = cad & "," & TransformaComasPuntos(CStr(Gastos)) & ",NULL,"

        'Codccost
        cad = cad & "NULL,"

        If LlevaContr Then
            If Not DescuentaImporteDevolucion Then
                cad = cad & "'" & txtCta(1).Text & "'"
            Else
                cad = cad & "'" & Text1(1).Text & "'"
            End If
        Else
            cad = cad & "NULL"
        End If

        If Cobro Then
            cad = cad & ",'COBROS',0,"
        Else
            cad = cad & ",'PAGOS',0,"
        End If
        ' todo valores a null ????
        cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
        
        cad = SQL & cad
        Conn.Execute cad
        
        
        If Not DescuentaImporteDevolucion Then
            cad = "4,'"
    
            cad = cad & txtCta(1).Text & "','" & Numdocum & "'," & Conce
            cad = cad & ",'Gastos vto.'"
    
            'Importe al debe
            cad = cad & ",NULL, " & TransformaComasPuntos(CStr(Gastos)) & ","
    
            'Codccost
            cad = cad & "NULL,"
    
            If LlevaContr Then
                cad = cad & "'" & CtaBancoGastos & "'"
            Else
                cad = cad & "NULL"
            End If
    
            If Cobro Then
                cad = cad & ",'COBROS',0,"
            Else
                cad = cad & ",'PAGOS',0,"
            End If
            ' todo valores a null ????
            cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
            
            
            cad = SQL & cad
            Conn.Execute cad
        
        End If
        
    End If
    '++
    
    'Insertamos en la temporal para que lo ac
    If Cobro Then
        SQL = FP.diaricli
    Else
        SQL = FP.diaripro
    End If
    
    'Actualizamos VTO
    ' o lo eliminamos. Segun sea el importe que falte
    'Tomomos prestada LlevaContr
    
    Im = ImporteFormateado(Text2(0).Text)  'lo que voy a anticipar
    
    impo = ImporteFormateado(Text1(3).Text)  'lo que me falta
    If Cobro Then
        If Text1(4).Text <> "" Then impo = impo + ImporteFormateado(Text1(4).Text)
        If Text1(5).Text <> "" Then impo = impo - ImporteFormateado(Text1(5).Text)
    Else
        If Text1(5).Text <> "" Then impo = impo - ImporteFormateado(Text1(5).Text)
    End If
    If impo - Im = 0 Then
        LlevaContr = True  'ELIMINAR VTO ya que esta totalmente pagado
    Else
        LlevaContr = False
    End If
    
    
    impo = ImporteFormateado(Text2(0).Text)
    If Cobro Then
        SQL = "cobros"
        Ampliacion = "fecultco"
        Numdocum = "impcobro"
        'El importe es el total. Lo que ya llevaba mas lo de ahora
        If Text1(5).Text <> "" Then impo = impo + ImporteFormateado(Text1(5).Text)
    Else
        
        SQL = "pagos"
        Ampliacion = "fecultpa"
        Numdocum = "imppagad"
        'El importe es el total. Lo que ya llevaba mas lo de ahora
        If Text1(5).Text <> "" Then impo = impo + ImporteFormateado(Text1(5).Text)
    End If
    
    
    '++monica
    Dim NumLin As Long
    
    If Cobro Then
        SQL = "update cobros set impcobro = coalesce(impcobro,0) + " & DBSet(Text2(0).Text, "N")
        SQL = SQL & ", fecultco = " & DBSet(Text3(0).Text, "F")
        SQL = SQL & " where numserie = " & DBSet(RecuperaValor(Vto, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(Vto, 2), "N")
        SQL = SQL & " and fecfactu = " & DBSet(RecuperaValor(Vto, 3), "F") & " and numorden = " & DBSet(RecuperaValor(Vto, 4), "N")
    
        Conn.Execute SQL
        
        SQL = "select impvenci + coalesce(gastos,0) - coalesce(impcobro,0) from cobros where numserie = " & DBSet(RecuperaValor(Vto, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(Vto, 2), "N")
        SQL = SQL & " and fecfactu = " & DBSet(RecuperaValor(Vto, 3), "F") & " and numorden = " & DBSet(RecuperaValor(Vto, 4), "N")
     
        'ahora es cuando ponemos la situacion
        Situacion = 0
        If DevuelveValor(SQL) = 0 Then
            Situacion = 1
        End If
    
        SQL = "update cobros set "
        SQL = SQL & " situacion = " & DBSet(Situacion, "N")
        SQL = SQL & " where numserie = " & DBSet(RecuperaValor(Vto, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(Vto, 2), "N")
        SQL = SQL & " and fecfactu = " & DBSet(RecuperaValor(Vto, 3), "F") & " and numorden = " & DBSet(RecuperaValor(Vto, 4), "N")
    
        Conn.Execute SQL
    
    Else
        
        SQL = "update pagos set imppagad = coalesce(imppagad,0) + " & DBSet(Text2(0).Text, "N")
        SQL = SQL & ", fecultpa = " & DBSet(Text3(0).Text, "F")
        SQL = SQL & " where numserie = " & DBSet(RecuperaValor(Vto, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(Vto, 2), "T")
        SQL = SQL & " and fecfactu = " & DBSet(RecuperaValor(Vto, 3), "F") & " and numorden = " & DBSet(RecuperaValor(Vto, 4), "N")
        SQL = SQL & " and codmacta = " & DBSet(RecuperaValor(Cta, 1), "T")
    
        Conn.Execute SQL
        
        SQL = "select impefect  - coalesce(imppagad,0) from pagos where numserie = " & DBSet(RecuperaValor(Vto, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(Vto, 2), "T")
        SQL = SQL & " and fecfactu = " & DBSet(RecuperaValor(Vto, 3), "F") & " and numorden = " & DBSet(RecuperaValor(Vto, 4), "N") & " and codmacta = " & DBSet(RecuperaValor(Cta, 1), "T")
     
        'ahora es cuando ponemos la situacion
        Situacion = 0
        If DevuelveValor(SQL) = 0 Then
            Situacion = 1
        End If
    
        SQL = "update pagos set "
        SQL = SQL & " situacion = " & DBSet(Situacion, "N")
        SQL = SQL & " where numserie = " & DBSet(RecuperaValor(Vto, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(Vto, 2), "T")
        SQL = SQL & " and fecfactu = " & DBSet(RecuperaValor(Vto, 3), "F") & " and numorden = " & DBSet(RecuperaValor(Vto, 4), "N")
        SQL = SQL & " and codmacta = " & DBSet(RecuperaValor(Cta, 1), "T")
    
        Conn.Execute SQL
    
    End If
    
    Contabilizar = True

    Set Mc = Nothing
    Set FP = Nothing

    Exit Function
ECon:
    MuestraError Err.Number, "Contabilizar anticipo"
    Set Mc = Nothing
    Set FP = Nothing
End Function
    
Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(1)
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)

        txtCta(Index).Text = Trim(txtCta(Index).Text)
        cad = txtCta(Index).Text
        impo = 0
        If cad <> "" Then
            If CuentaCorrectaUltimoNivel(cad, CadenaDesdeOtroForm) Then
                cad = DevuelveDesdeBD("codmacta", "bancos", "codmacta", cad, "T")
                If cad = "" Then
                    CadenaDesdeOtroForm = ""
                    MsgBox "La cuenta contable no esta asociada a ninguna cuenta bancaria", vbExclamation
                End If
            Else
                MsgBox CadenaDesdeOtroForm, vbExclamation
                cad = ""
                CadenaDesdeOtroForm = ""
            End If
            impo = 1
        Else
            CadenaDesdeOtroForm = ""
        End If
        
        
        txtCta(Index).Text = cad
        txtDescCta(Index).Text = CadenaDesdeOtroForm
        If cad = "" And impo <> 0 Then
            PonFoco txtCta(Index)
        End If
        CadenaDesdeOtroForm = ""
End Sub


Private Sub CargaCombo()
    Combo1.Clear
    'Conceptos
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from tipofpago order by tipoformapago", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!descformapago
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!tipoformapago
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub


