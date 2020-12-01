VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFVARContabFact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contabilización de Facturas "
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7620
   Icon            =   "frmFVARContabFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   7515
      Left            =   90
      TabIndex        =   10
      Top             =   120
      Width           =   7410
      Begin VB.TextBox txtNombre 
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
         Index           =   5
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   2535
         Width           =   5205
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   4
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   2160
         Width           =   5205
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   8
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   1350
         Width           =   1350
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   8
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   1350
         Width           =   4305
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   7
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   570
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2520
         Width           =   540
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   2
         Top             =   2145
         Width           =   540
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   3
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3735
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3345
         Width           =   1350
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   6015
         TabIndex        =   9
         Top             =   6735
         Width           =   1065
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
         Left            =   4830
         TabIndex        =   8
         Top             =   6735
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   6
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   4500
         Width           =   1030
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   7
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   4890
         Width           =   1030
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   330
         TabIndex        =   20
         Top             =   5580
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1125
         MouseIcon       =   "frmFVARContabFact.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar contador"
         Top             =   2535
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1125
         MouseIcon       =   "frmFVARContabFact.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar contador"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   330
         TabIndex        =   25
         Top             =   6330
         Width           =   6825
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   5970
         Width           =   6795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Banco"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   5
         Left            =   360
         TabIndex        =   23
         Top             =   1005
         Width           =   1500
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1110
         MouseIcon       =   "frmFVARContabFact.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cta.banco"
         Top             =   1350
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1110
         Picture         =   "frmFVARContabFact.frx":0402
         ToolTipText     =   "Buscar fecha"
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Vencimiento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   22
         Top             =   255
         Width           =   2445
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Letra de Serie"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   19
         Top             =   1815
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
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
         Index           =   1
         Left            =   450
         TabIndex        =   18
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
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
         Index           =   0
         Left            =   450
         TabIndex        =   17
         Top             =   2145
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   16
         Left            =   330
         TabIndex        =   16
         Top             =   3015
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
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
         Index           =   15
         Left            =   450
         TabIndex        =   15
         Top             =   3345
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
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
         Index           =   14
         Left            =   450
         TabIndex        =   14
         Top             =   3735
         Width           =   645
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1110
         Picture         =   "frmFVARContabFact.frx":048D
         ToolTipText     =   "Buscar fecha"
         Top             =   3345
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1110
         Picture         =   "frmFVARContabFact.frx":0518
         ToolTipText     =   "Buscar fecha"
         Top             =   3735
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
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
         Index           =   13
         Left            =   450
         TabIndex        =   13
         Top             =   4500
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
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
         Index           =   12
         Left            =   450
         TabIndex        =   12
         Top             =   4905
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   11
         Left            =   360
         TabIndex        =   11
         Top             =   4170
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmFVARContabFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmBan As frmBasico2 'Cuentas contables
Attribute frmBan.VB_VarHelpID = -1
Private WithEvents frmC As frmCal  'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico ' contadores
Attribute frmConta.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim IndCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Dim BdConta As Integer

Dim cContaFra As cContabilizarFacturas



Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim cadMen As String
Dim I As Byte
Dim Sql As String
Dim Tipo As Byte
Dim Nregs As Long
Dim NumError As Long

    If Not DatosOK Then Exit Sub
    
    cadselect = tabla & ".intconta=0 "
    
    
    
    If Not PonerDesdeHasta("fvarfactura.fecfactu", "F", Me.txtCodigo(2), Me.txtCodigo(2), Me.txtCodigo(3), Me.txtCodigo(3), "pDHFecha=""") Then Exit Sub
    If Not PonerDesdeHasta("fvarfactura.numserie", "SER", Me.txtCodigo(4), Me.txtNombre(4), Me.txtCodigo(5), Me.txtNombre(5), "pDHSerie=""") Then Exit Sub
    If Not PonerDesdeHasta("fvarfactura.numfactu", "FRA", Me.txtCodigo(0), Me.txtCodigo(0), Me.txtCodigo(1), Me.txtCodigo(1), "pDHNumfactu=""") Then Exit Sub
    
    
    If Not HayRegParaInforme(tabla, cadselect) Then Exit Sub
    
    ContabilizarFacturas tabla, cadselect
     'Eliminar la tabla TMP
    BorrarTMPFacturas
    'Desbloqueamos ya no estamos contabilizando facturas
    BloqueoManual False, "", "VENCON" 'VENtas CONtabilizar
        
    
eError:
    If Err.Number <> 0 Or NumError <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización. Llame a soporte."
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        ValoresPorDefecto
        PonFoco txtCodigo(7)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    Limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmppal.imgIcoForms.ListImages(1).Picture
     Me.imgBuscar(8).Picture = frmppal.imgIcoForms.ListImages(1).Picture
     
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    tabla = "fvarfactura"
    
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
    
    txtCodigo(2).Text = Format(Now, "dd/mm/yyyy")
    txtCodigo(3).Text = Format(Now, "dd/mm/yyyy")
    
End Sub


Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
' cta de banco
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nommacta
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(IndCodigo).Text = Format(vFecha, "dd/MM/yyyy")
    PonFoco txtCodigo(IndCodigo)
End Sub

Private Sub frmConta_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
        txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub imgFec_Click(Index As Integer)
    
    Set frmC = New frmCal

    IndCodigo = Index
    If txtCodigo(IndCodigo).Text <> "" Then frmC.wndDatePicker.Select (CDate(txtCodigo(IndCodigo).Text))
    frmC.Show vbModal
    
    Set frmC = Nothing
    
    PonFoco txtCodigo(IndCodigo)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim Sql As String

   Select Case Index
        Case 0, 1 ' contadores
            IndCodigo = Index + 4
        
            Set frmConta = New frmBasico
            AyudaContadores frmConta, txtCodigo(Index), "tiporegi REGEXP '^[0-9]+$' = 0"
            Set frmConta = Nothing
    
            PonFoco Me.txtCodigo(Index)
        
   
        Case 8 ' Cta Contable de Banco
            IndCodigo = 8
            Set frmBan = New frmBasico2
            AyudaBanco frmBan
            Set frmBan = Nothing
            PonFoco Me.txtCodigo(Index)
    End Select
    PonFoco txtCodigo(IndCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Me.Caption = "Facturas por Cliente"
        Case 1
            Me.Caption = "Facturas por Tarjeta"
        Case 2
            Me.Caption = "Facturas por Cliente y por Tarjeta"
    End Select
    
End Sub

Private Sub txtcodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtcodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtcodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes estaba esto
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 8: KEYBusqueda KeyAscii, 8 'cta banco
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 4: KEYBusqueda KeyAscii, 0 'contador desde
            Case 5: KEYBusqueda KeyAscii, 1 'contador hasta
            Case 7: KEYFecha KeyAscii, 7 'fecha de vencimiento
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub txtcodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim Sql As String, Sql2 As String

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 8 ' Cuenta de Banco
'                    txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(8), 1, , BdConta, True) 'DevuelveDesdeBDNewFac("cuentas", "nommacta", "codmacta", txtCodigo(Index), "N")
            
            If txtCodigo(Index).Text = "" Then Exit Sub
            
            Sql = txtCodigo(Index).Text
            If CuentaCorrectaUltimoNivel(Sql, Sql2) Then
                Sql = DevuelveDesdeBD("codmacta", "bancos", "codmacta", Sql, "T")
                If Sql = "" Then
                    MsgBox "La cuenta NO pertenece a ningúna cta. bancaria", vbExclamation
                    Sql2 = ""
                Else
                    'CORRECTO
                End If
            Else
                Sql = ""
                MsgBox Sql2, vbExclamation
                Sql2 = ""
            End If
            txtCodigo(Index).Text = Sql
            txtNombre(Index).Text = Sql2
            If Sql = "" Then PonFoco txtCodigo(Index)
            
        Case 2, 3, 7  'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        
        Case 0, 1 ' NUMERO DE FACTURA
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
        
        Case 4, 5 ' LETRA DE SERIE
            txtCodigo(Index).Text = UCase(Trim(txtCodigo(Index).Text))
            
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "contadores", "nomregis", "tiporegi", "T")
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 7515
        Me.FrameCobros.Width = 7410
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub

Private Sub ValoresPorDefecto()
    txtCodigo(7).Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
End Sub

'Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
''IN: codD,codH --> codigo Desde/Hasta
''    nomD,nomH --> Descripcion Desde/Hasta
''Añade a cadFormula y cadSelect la cadena de seleccion:
''       "(codigo>=codD AND codigo<=codH)"
'' y añade a cadParam la cadena para mostrar en la cabecera informe:
''       "codigo: Desde codD-nomd Hasta: codH-nomH"
'Dim Devuelve As String
'Dim devuelve2 As String
'
'    PonerDesdeHasta = False
'    Devuelve = CadenaDesdeHasta(codD, codH, Codigo, TipCod)
'    If Devuelve = "Error" Then Exit Function
'    If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Function
'    If TipCod <> "F" Then 'Fecha
'        If Not AnyadirAFormula(cadselect, Devuelve) Then Exit Function
'    Else
'        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
'        If devuelve2 = "Error" Then Exit Function
'        If Not AnyadirAFormula(cadselect, devuelve2) Then Exit Function
'    End If
'    If Devuelve <> "" Then
'        If param <> "" Then
'            'Parametro Desde/Hasta
'            cadParam = cadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
'            numParam = numParam + 1
'        End If
'        PonerDesdeHasta = True
'    End If
'End Function


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim cad As String

    DatosOK = False


    If txtCodigo(7).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una Fecha de Vencimiento.", vbExclamation
        PonFoco txtCodigo(7)
        Exit Function
    End If
    
    If txtCodigo(8).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una Cta.Banco para realizar el cobro.", vbExclamation
        PonFoco txtCodigo(8)
        Exit Function
    Else
        txtNombre(8).Text = DevuelveDesdeBD("descripcion", "bancos", "codmacta", txtCodigo(8), "T")
        If txtNombre(8).Text = "" Then
            PonFoco txtCodigo(8)
            Exit Function
        End If
        Orden1 = ""
        Orden1 = vParam.fechaini

        Orden2 = ""
        Orden2 = vParam.fechafin
        'comprobar que se han rellenado los dos campos de fecha
        'sino rellenar con fechaini o fechafin del ejercicio
        'que guardamos en vbles Orden1,Orden2
        If txtCodigo(2).Text = "" Then
            txtCodigo(2).Text = Orden1 'fechaini del ejercicio de la conta
        End If
    
        If txtCodigo(3).Text = "" Then
            txtCodigo(3).Text = DateAdd("yyyy", 1, CDate(Orden2))  'fecha fin del ejercicio de la conta
        End If
        If Not ComprobarFechasConta(2) Then Exit Function
        If Not ComprobarFechasConta(3) Then Exit Function
                
    End If
    
    DatosOK = True
End Function

' copiado del ariges
Private Sub ContabilizarFacturas(cadTABLA As String, CadWhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim Sql As String
Dim B As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cad As String

    Sql = "VENCON" 'contabilizar facturas de venta

    'Bloquear para que nadie mas pueda contabilizar
    BloqueoManual False, "", Sql
    If Not BloqueoManual(True, "", Sql) Then
        MsgBox "No se pueden Contabilizar Facturas. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2
     If txtCodigo(2).Text = "" Then
        txtCodigo(2).Text = Orden1 'vEmpresa.FechaIni  'fechaini del ejercicio de la conta
     End If

     If txtCodigo(3).Text = "" Then
        txtCodigo(3).Text = Orden2 'vEmpresa.FechaFin  'fecha fin del ejercicio de la conta
     End If

     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los parámetros
     If Not ComprobarFechasConta(3) Then Exit Sub
     
     

    'La comprobacion solo lo hago para facturas nuestras, ya que mas adelante
    'el programa hara cdate(text1(31) cuando contabilice las facturas y dara error de tipos
    If Me.txtCodigo(2).Text = "" Then
        MsgBox "Fecha inicio incorrecta", vbExclamation
        Exit Sub
    End If



    'comprobar si existen en Ariagroutil facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtCodigo(2).Text <> "" Then
        Sql = "SELECT COUNT(*) FROM " & cadTABLA
        Sql = Sql & " WHERE fecfactu <"
        Sql = Sql & DBSet(txtCodigo(2), "F") & " AND intconta=0 "
        If TotalRegistros(Sql) > 0 Then
            '[Monica]11/10/2011: indico si es de esta seccion o de otra seccion
            Sql = "select count(*) from " & cadTABLA
            Sql = Sql & " WHERE fecfactu <"
            Sql = Sql & DBSet(txtCodigo(2), "F") & " AND intconta=0 "
            If TotalRegistros(Sql) > 0 Then
                cad = "Hay Facturas anteriores sin contabilizar." & vbCrLf
            End If
            cad = cad & "            ¿ Desea continuar ? "
            If MsgBox(cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
'    Me.Pb1.Top = 3350
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100
        
    BorrarTMPFacturas
    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    B = CrearTMPFacturas(cadTABLA, CadWhere, True)
    If Not B Then Exit Sub
            
    BorrarTMPErrComprob
    ' nuevo
    B = CrearTMPErrComprob()
    If Not B Then Exit Sub
    
    'comprobar que no haya Nº FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Nº Facturas en contabilidad ..."
    Sql = "anofactu>=" & Year(txtCodigo(2).Text) & " AND anofactu<= " & Year(txtCodigo(3).Text)
    B = ComprobarNumFacturasFacContaNueva(Sql)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not B Then
        frmFVARInformes.OpcionListado = 2
        frmFVARInformes.Show vbModal
        Exit Sub
    End If
    
    
    'comprobar que todas las CUENTAS de retencion de las distintas facturas que vamos a
    'contabilizar existen en la Conta: cabfact.cuereten IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    B = ComprobarCtaContableFac(8, cadselect)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not B Then
        frmFVARInformes.OpcionListado = 2
        frmFVARInformes.Show vbModal
        Exit Sub
    End If
    
    
    
    'comprobar que todas las CUENTAS de venta de la familia de los articulos que vamos a
    'contabilizar existen en la Conta: sfamia.ctaventa IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    B = ComprobarCtaContableFac(2, cadselect)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not B Then
        frmFVARInformes.OpcionListado = 2
        frmFVARInformes.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas las CUENTAS de venta de los conceptos que vamos a
    'contabilizar son de grupo de ventas: empiezan por conta.parametros.grupovtas
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    B = ComprobarCtaContableFac(3)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not B Then
        frmFVARInformes.OpcionListado = 2
        frmFVARInformes.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas la CUENTA del banco propio donde contabilizar el cobro
    'que existen en la Conta: sbanpr.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables del Banco en contabilidad ..."
    
    B = ComprobarCtaContableFac(4, CStr(txtCodigo(8).Text))
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not B Then
        frmFVARInformes.OpcionListado = 2
        frmFVARInformes.Show vbModal
        Exit Sub
    End If
    
    
    'comprobar que todos las TIPO IVA de las distintas facturas que vamos a
    'contabilizar existen en la Conta: schfac.codigiv1,codigiv2,codigiv3 IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    B = ComprobarTiposIVA
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not B Then
        frmFVARInformes.OpcionListado = 2 '3
        frmFVARInformes.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todos los CENTRO DE COSTE de las distintas facturas que vamos a
    'contabilizar existen en la Conta: codccost in conta.cabccost
    '--------------------------------------------------------------------------
    If vParam.autocoste Then
        Me.lblProgres(1).Caption = "Comprobando Centros Coste en contabilidad ..."
        B = ComprobarCCoste()
        IncrementarProgres Me.Pb1, 10
        Me.Refresh
        If Not B Then
            frmFVARInformes.OpcionListado = 2 '0
            frmFVARInformes.Show vbModal
            Exit Sub
        End If
    Else
        IncrementarProgres Me.Pb1, 10
        Me.Refresh
    End If
    
    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Facturas: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad..."
       
    
    'Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTABLA)
    
    
    B = PasarFacturasAContab(cadTABLA, txtCodigo(7).Text, txtCodigo(8).Text, CCoste)
    
    If Not B Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmFVARInformes.OpcionListado = 2 '10
            frmFVARInformes.Show vbModal
        Else
            MsgBoxA "No pueden mostrarse los errores.", vbExclamation
        End If
    Else
        MsgBoxA "El proceso ha finalizado correctamente.", vbInformation
    End If
    
    'Eliminar tabla TEMP de Errores
    BorrarTMPErrFact
    BorrarTMPErrComprob
End Sub


Private Function PasarFacturasAContab(cadTABLA As String, FecVenci As String, Banpr As String, CCoste As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim B As Boolean
Dim I As Integer
Dim numfactu As Long
Dim codigo1 As String

    On Error GoTo EPasarFac

    PasarFacturasAContab = False
    
    'Total de Facturas a Insertar en la contabilidad
    Sql = "SELECT count(*) "
    Sql = Sql & " FROM " & cadTABLA & " INNER JOIN tmpfactu "
    codigo1 = "numserie"
    Sql = Sql & " ON " & cadTABLA & "." & codigo1 & "=tmpfactu.numserie"
    Sql = Sql & " AND " & cadTABLA & ".numfactu=tmpfactu.numfactu AND " & cadTABLA & ".fecfactu=tmpfactu.fecfactu "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        numfactu = Rs.Fields(0)
    Else
        numfactu = 0
    End If
    Rs.Close
    Set Rs = Nothing

    If numfactu > 0 Then
        CargarProgres Me.Pb1, numfactu
        
        Set cContaFra = New cContabilizarFacturas
        
        If Not cContaFra.EstablecerValoresInciales(Conn) Then
            'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
            ' obviamente, no va a contabilizar las FRAS
            Sql = "Si continua, las facturas se insertaran en el registro, pero no serán contabilizadas" & vbCrLf
            Sql = Sql & "en este momento. Deberán ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
            Sql = Sql & Space(50) & "¿Continuar?"
            If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
        End If
        
        
        
        Sql = "SELECT * "
        Sql = Sql & " FROM tmpfactu "
            
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenStatic, adLockPessimistic, adCmdText
        I = 1

        B = True
        'contabilizar cada una de las facturas seleccionadas
        While Not Rs.EOF
            Sql = ""
            Sql = Sql & cadTABLA & "." & codigo1 & "=" & DBSet(Rs.Fields(0), "T") & " and numfactu=" & DBLet(Rs!numfactu, "N")
            Sql = Sql & " and fecfactu=" & DBSet(Rs!FecFactu, "F")
            If PasarFacturaFac(Sql, FecVenci, Banpr, CCoste, cContaFra) = False And B Then B = False
            
            IncrementarProgres Me.Pb1, 1
            Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad...   (" & I & " de " & numfactu & ")"
            Me.Refresh
            I = I + 1
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
    End If
    
EPasarFac:
    If Err.Number <> 0 Then B = False
    
    If B Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function


Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim fechaini As String, fechafin As String
Dim cad As String
Dim Rs As ADODB.Recordset
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    If txtCodigo(ind).Text <> "" Then
        fechaini = "Select fechaini,fechafin From parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open fechaini, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs.EOF Then
            fechaini = DBLet(Rs!fechaini, "F")
            fechafin = DateAdd("yyyy", 1, CDate(DBLet(Rs!fechafin, "F"))) ' + 365
            'nos guardamos los valores
            Orden1 = fechaini
            Orden2 = fechafin
        
            If Not EntreFechas(fechaini, txtCodigo(ind).Text, fechafin) Then
                 cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 cad = cad & "    Desde: " & fechaini & vbCrLf
                 cad = cad & "    Hasta: " & fechafin
                 MsgBox cad, vbExclamation
                 txtCodigo(ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        ComprobarFechasConta = True
    End If
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function



