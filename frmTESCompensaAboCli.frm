VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESCompensaAboCli 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compensación abonos cliente"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13050
   Icon            =   "frmTESCompensaAboCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   13050
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCompensaAbonosCliente 
      BorderStyle     =   0  'None
      Height          =   6885
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12735
      Begin VB.Frame FrameBotonGnral 
         Height          =   705
         Left            =   180
         TabIndex        =   13
         Top             =   150
         Width           =   2325
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   330
            Left            =   240
            TabIndex        =   14
            Top             =   180
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   5
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Establecer Vencimiento"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Ver todos"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Cta Anterior"
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Cta Siguiente"
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtimpNoEdit 
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
         Index           =   2
         Left            =   10380
         TabIndex        =   11
         Top             =   5790
         Width           =   1845
      End
      Begin VB.TextBox txtimpNoEdit 
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
         Index           =   1
         Left            =   10260
         TabIndex        =   10
         Top             =   1170
         Width           =   2025
      End
      Begin VB.TextBox txtimpNoEdit 
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
         Index           =   0
         Left            =   8250
         TabIndex        =   7
         Top             =   1170
         Width           =   1965
      End
      Begin MSComctlLib.ListView lwCompenCli 
         Height          =   3675
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   12315
         _ExtentX        =   21722
         _ExtentY        =   6482
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
            Text            =   "Tipo"
            Object.Width           =   1410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Factura"
            Object.Width           =   2116
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Vto"
            Object.Width           =   1234
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha Vto"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Forma pago"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Cobro"
            Object.Width           =   3590
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Abonos"
            Object.Width           =   3590
         EndProperty
      End
      Begin VB.TextBox DtxtCta 
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
         Index           =   17
         Left            =   1560
         TabIndex        =   4
         Text            =   "Text5"
         Top             =   1170
         Width           =   4905
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
         Index           =   17
         Left            =   240
         TabIndex        =   3
         Top             =   1170
         Width           =   1305
      End
      Begin VB.CommandButton cmdCompensar 
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
         Left            =   10110
         TabIndex        =   2
         Top             =   6300
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   0
         Left            =   11430
         TabIndex        =   1
         Top             =   6300
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   7950
         Top             =   150
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   11820
         Picture         =   "frmTESCompensaAboCli.frx":000C
         ToolTipText     =   "Quitar al Debe"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   12180
         Picture         =   "frmTESCompensaAboCli.frx":0156
         ToolTipText     =   "Puntear al Debe"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   72
         Left            =   9300
         TabIndex        =   12
         Top             =   5835
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Rectifca./Abono"
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
         Index           =   71
         Left            =   10740
         TabIndex        =   9
         Top             =   900
         Width           =   1590
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cobro"
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
         Index           =   70
         Left            =   9300
         TabIndex        =   8
         Top             =   900
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
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
         Index           =   69
         Left            =   240
         TabIndex        =   5
         Top             =   900
         Width           =   1440
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   17
         Left            =   1770
         Top             =   900
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6885
      Left            =   90
      TabIndex        =   15
      Top             =   -30
      Visible         =   0   'False
      Width           =   12735
      Begin VB.CommandButton cmdCancelar 
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
         Left            =   11550
         TabIndex        =   18
         Top             =   6300
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   240
         TabIndex        =   16
         Top             =   180
         Width           =   3585
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   180
            TabIndex        =   19
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   10
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Nuevo"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
                  Object.Tag             =   "2"
                  Object.Width           =   1e-4
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Buscar"
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Ver Todos"
                  Object.Tag             =   "0"
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Imprimir"
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Salir"
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   5085
         Left            =   240
         TabIndex        =   17
         Top             =   990
         Width           =   12315
         _ExtentX        =   21722
         _ExtentY        =   8969
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cuenta"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre"
            Object.Width           =   10054
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   12060
         TabIndex        =   20
         Top             =   210
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
   End
End
Attribute VB_Name = "frmTESCompensaAboCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SaltoLinea = """ + chr(13) + """

Private Const IdPrograma = 607


    
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private frmMens As frmMensajes

Dim SQL As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim Cad As String
Dim CONT As Long
Dim i As Integer
Dim TotalRegistros As Long

Dim Importe As Currency
Dim MostrarFrame As Boolean
Dim Fecha As Date

Dim DevfrmCCtas As String

Dim CampoOrden As String
Dim Orden As Boolean
Dim Modo As Byte

Dim Txt33Csb As String
Dim Txt41Csb As String

Dim VerTodos As Boolean
Dim Cuentas As String
Dim Observaciones As String


Private Sub PonFoco(ByRef T1 As TextBox)
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
End Sub

Private Function ComprobarObjeto(ByRef T As TextBox) As Boolean
    Set miTag = New CTag
    ComprobarObjeto = False
    If miTag.Cargar(T) Then
        If miTag.Cargado Then
            If miTag.Comprobar(T) Then ComprobarObjeto = True
        End If
    End If

    Set miTag = Nothing
End Function



Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 0 Then
        Frame1.visible = True
        Frame1.Enabled = True
        Me.FrameCompensaAbonosCliente.visible = False
        Me.FrameCompensaAbonosCliente.Enabled = False
        CargaList
    Else
        Unload Me
    End If
    
    If Index = 0 Then BotonVerTodos True
End Sub


Private Sub cmdCompensar_Click()
    
    Cad = DevuelveDesdeBD("informe", "scryst", "codigo", IdPrograma) 'Orden de pago a bancos
    If Cad = "" Then
        MsgBox "No esta configurada la aplicación. Falta el informe", vbCritical
        Exit Sub
    End If
    Me.Tag = Cad
    
    Cad = ""
    RC = ""
    CONT = 0
    TotalRegistros = 0
    NumRegElim = 0
    For i = 1 To Me.lwCompenCli.ListItems.Count
        If Me.lwCompenCli.ListItems(i).Checked Then
            If Trim(lwCompenCli.ListItems(i).SubItems(6)) = "" Then
                'Es un abono
                TotalRegistros = TotalRegistros + 1
            Else
                NumRegElim = NumRegElim + 1
            End If
        End If
        If Me.lwCompenCli.ListItems(i).Bold Then
            Cad = Cad & "A"
            If CONT = 0 Then CONT = i
        End If
    Next

    If TotalRegistros + NumRegElim = 1 Then
        MsgBox "Debe seleccionar mas de un vencimiento", vbExclamation
        Exit Sub
    End If
    If TotalRegistros = 0 Or NumRegElim = 0 Then
        If TotalRegistros = 0 And NumRegElim = 0 Then
            MsgBox "No ha marcado ningun venciminto", vbExclamation
            Exit Sub
        End If
        SQL = "-No va a realizar compensaciones. Va a agrupar los cobros. " & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    End If

    i = 0
    SQL = ""
    If Len(Cad) <> 1 Then
        'Ha seleccionado o cero o mas de uno
        If txtimpNoEdit(0).Text <> txtimpNoEdit(1).Text Then
            'importes distintos. Solo puede seleccionar UNO
            SQL = "Debe selecionar uno(y solo uno) como vencimiento destino"
        End If
    Else
        'Comprobaremos si el selecionado esta tb checked
        If Not lwCompenCli.ListItems(CONT).Checked Then
            SQL = "El vencimiento seleccionado no esta marcado"
        
        Else
            'Si el importe Cobro es mayor que abono, deberia estar
            Importe = CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag)  'txtimpNoEdit(1).Tag es negativo
            If Importe <> 0 Then
                If Importe > 0 Then
                    'Es un abono
                    If Trim(lwCompenCli.ListItems(CONT).SubItems(6)) = "" Then SQL = "cobro"
                Else
                    If Trim(lwCompenCli.ListItems(CONT).SubItems(6)) <> "" Then SQL = "abono"
                End If
                If SQL <> "" Then SQL = "Debe marcar un " & SQL & " como destino"
            End If
            
        End If
    End If
    
    
    'Nuevo. Mao 2017
    'Si hay uno y uno, y no esta el vencimiento destino establecido, lo establezco yo
    If TotalRegistros = 1 And NumRegElim = 1 And CONT = 0 Then
        Importe = CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag)  'txtimpNoEdit(1).Tag es negativo
        SQL = ""
        For i = 1 To Me.lwCompenCli.ListItems.Count
            If lwCompenCli.ListItems(i).Checked Then
                If Importe > 0 Then
                    'Este es el nodo SELECCIONADO
                    If Trim(lwCompenCli.ListItems(i).SubItems(6)) <> "" Then SQL = "OK"
                Else
                    If Trim(lwCompenCli.ListItems(i).SubItems(6)) = "" Then SQL = "OK"
                End If
            End If
            If SQL <> "" Then
                'ESTE ES EL NODO A MARCAR
                
                Me.lwCompenCli.ListItems(i).Bold = True
                Me.lwCompenCli.ListItems(i).ForeColor = vbRed
                For CONT = 1 To Me.lwCompenCli.ColumnHeaders.Count - 1
                    Me.lwCompenCli.ListItems(i).ListSubItems(CONT).ForeColor = vbRed
                    Me.lwCompenCli.ListItems(i).ListSubItems(CONT).Bold = True
                Next
                lwCompenCli.Refresh
                lwCompenCli.ListItems(i).EnsureVisible
                CONT = i  'establezco destino
                Cad = "1" 'Para que la funcion de comprbacion no de que debe seleccionar un vencimiento
                SQL = ""
                Exit For
            End If
        Next
        
        
      
    
    End If
    
    
    
    
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    
    
    
    
    
    
    
    
    If TotalRegistros = 0 Or NumRegElim = 0 Then SQL = "Debe selecionar cobro(s) y abono(s)" & vbCrLf & SQL
        
    Observaciones = ""

        Dim CadAux As String
        
        Txt33Csb = "Compensa: "
        Txt41Csb = ""
        For i = 1 To Me.lwCompenCli.ListItems.Count - 1
            If Me.lwCompenCli.ListItems(i).Checked Then
            
            
                Observaciones = Observaciones & lwCompenCli.ListItems(i).Text & lwCompenCli.ListItems(i).SubItems(1) & " " & Trim(lwCompenCli.ListItems(i).SubItems(2))
                Observaciones = Observaciones & " " & Trim(lwCompenCli.ListItems(i).SubItems(6) & " " & Trim(lwCompenCli.ListItems(i).SubItems(7)))
                If i = CONT Then Observaciones = Observaciones & "   Vto destino . Resultado final: " & Me.txtimpNoEdit(2).Text
                Observaciones = Observaciones & vbCrLf
            
            
            
                CadAux = Trim(lwCompenCli.ListItems(i).Text & lwCompenCli.ListItems(i).SubItems(1)) & " " & Trim(lwCompenCli.ListItems(i).SubItems(2))
                If Len(Txt33Csb & " " & CadAux) < 80 Then
                    Txt33Csb = Txt33Csb & " " & CadAux
                Else
                    If Len(Txt41Csb & " " & CadAux) < 60 Then
                        Txt41Csb = Txt41Csb & CadAux
                    Else
                        Txt41Csb = Txt41Csb & ".."
                        Exit For
                    End If
                End If
            End If
        Next i
        

    
    If MsgBox("Seguro que desea realizar la compensación?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    
    Me.FrameCompensaAbonosCliente.Enabled = False
    Me.Refresh
    Screen.MousePointer = vbHourglass
    
    RealizarCompensacionAbonosClientes
    Me.FrameCompensaAbonosCliente.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdVtoDestino(Index As Integer)
    
    If Index = 0 Then
        TotalRegistros = 0
        If Not Me.lwCompenCli.SelectedItem Is Nothing Then TotalRegistros = Me.lwCompenCli.SelectedItem.Index
    
    
        For i = 1 To Me.lwCompenCli.ListItems.Count
            If Me.lwCompenCli.ListItems(i).Bold Then
                Me.lwCompenCli.ListItems(i).Bold = False
                Me.lwCompenCli.ListItems(i).ForeColor = vbBlack
                For CONT = 1 To Me.lwCompenCli.ColumnHeaders.Count - 1
                    Me.lwCompenCli.ListItems(i).ListSubItems(CONT).ForeColor = vbBlack
                    Me.lwCompenCli.ListItems(i).ListSubItems(CONT).Bold = False
                Next
            End If
        Next
        Me.Refresh
        
        If TotalRegistros > 0 Then
            i = TotalRegistros
            Me.lwCompenCli.ListItems(i).Bold = True
            Me.lwCompenCli.ListItems(i).ForeColor = vbRed
            For CONT = 1 To Me.lwCompenCli.ColumnHeaders.Count - 1
                Me.lwCompenCli.ListItems(i).ListSubItems(CONT).ForeColor = vbRed
                Me.lwCompenCli.ListItems(i).ListSubItems(CONT).Bold = True
            Next
        End If
        lwCompenCli.Refresh
        
        PonerFocoLw Me.lwCompenCli

    Else
    
        frmTESCompensaAboCliImp.pCodigo = Me.lw1.SelectedItem
        frmTESCompensaAboCliImp.Show vbModal

    End If
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        If CadenaDesdeOtroForm <> "" Then
            Cuentas = Mid(CadenaDesdeOtroForm, 1, Len(CadenaDesdeOtroForm) - 1)
            txtCta(17).Text = RecuperaValorNew(CadenaDesdeOtroForm, ",", 1)
            txtCta(17).Text = Mid(txtCta(17).Text, 2, Len(txtCta(17).Text) - 2)
            
            BotonVerTodos False
            
            BotonAnyadir
            VerTodos = True
            txtCta_LostFocus (17)
            
            CadenaDesdeOtroForm = ""
        Else
            If Not Frame1.visible Then
                If CadenaDesdeOtroForm <> "" Then
                    txtCta(17).Text = CadenaDesdeOtroForm
                    txtCta_LostFocus 17
                Else
                    PonFoco txtCta(17)
                End If
                CadenaDesdeOtroForm = ""
            End If
            CargaList
             Me.cmdCancelar(1).Cancel = True
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub



    
Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim Img As Image


    Limpiar Me
    Me.Icon = frmppal.Icon
    CargaImagenesAyudas Me.Image3, 1, "Cuenta contable"
    
    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With
    
    
    'La toolbar
    With Toolbar2
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 4
        .Buttons(2).Image = 1
        .Buttons(4).Image = 7
        .Buttons(5).Image = 8
        
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    
    'Limpiamos el tag
    PrimeraVez = True
    CommitConexion  'Porque son listados. No hay nada dentro transaccion
    
    H = FrameCompensaAbonosCliente.Height + 120
    W = FrameCompensaAbonosCliente.Width
    
    FrameCompensaAbonosCliente.visible = False
    Me.Frame1.visible = True
    
    VerTodos = False
    
    Me.Width = W + 300
    Me.Height = H + 400
    
   
    
    PonerModoUsuarioGnral 0, "ariconta"
    
    Orden = True
    
End Sub


Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 1)
    DtxtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub Image3_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmCta = New frmColCtas
    RC = Index
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 3
    frmCta.Show vbModal
    Set frmCta = Nothing
    If Index = 17 Then PonerVtosCompensacionCliente
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim IT
Dim i As Integer
    For i = 1 To Me.lwCompenCli.ListItems.Count
        Set IT = lwCompenCli.ListItems(i)
        lwCompenCli.ListItems(i).Checked = (Index = 1)
        lwCompenCli_ItemCheck (IT)
        Set IT = Nothing
    Next i
End Sub

Private Sub lw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Campo2 As Integer

    Orden = Not Orden
    
    Select Case ColumnHeader
        Case "Código"
            CampoOrden = "compensa.codigo"
        Case "Fecha"
            CampoOrden = "compensa.fecha"
        Case "Cuenta"
            CampoOrden = "compensa.codmacta"
        Case "Nombre"
            CampoOrden = "compensa.nommacta"
    End Select
    CargaList


End Sub

Private Sub lw1_DblClick()
    ' ver facturas de las compensaciones
    
    Set frmMens = New frmMensajes
    
    frmMens.Opcion = 57
    frmMens.Parametros = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|"
    frmMens.Show vbModal
    
    Set frmMens = Nothing
    
End Sub

Private Sub lwCompenCli_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim C As Currency
Dim Cobro As Boolean

    Cobro = True
    C = Item.Tag
    If Trim(Item.SubItems(6)) = "" Then
        'Es un abono
        Cobro = False
        C = -C
    
    End If
    
    'Si no es checkear cambiamos los signos
    If Not Item.Checked Then C = -C
    
    i = 0
    If Not Cobro Then i = 1
    
    Me.txtimpNoEdit(i).Tag = Me.txtimpNoEdit(i).Tag + C
    txtimpNoEdit(i).Text = Format(Abs(txtimpNoEdit(i).Tag))
    txtimpNoEdit(2).Text = Format(CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag), FormatoImporte)
            
    If ComprobarCero(txtimpNoEdit(0).Text) = 0 Then txtimpNoEdit(0).Text = ""
    If ComprobarCero(txtimpNoEdit(1).Text) = 0 Then txtimpNoEdit(1).Text = ""
    If ComprobarCero(txtimpNoEdit(2).Text) = 0 Then txtimpNoEdit(2).Text = ""
    
            
End Sub

Private Sub HacerToolBar(Boton As Integer)

    Select Case Boton
        Case 1
            BotonAnyadir
        Case 2
'            BotonModificar
        Case 3
'            BotonEliminar False
        Case 5
'            BotonBuscar
        Case 6 ' ver todos
            CargaList
        Case 8
            'Imprimir factura
            
             cmdVtoDestino (1)

    End Select
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar2 Button.Index
End Sub

Private Sub HacerToolBar2(Boton As Integer)

    Select Case Boton
        Case 1
            cmdVtoDestino (0)
        Case 2 ' ver todos
            BotonVerTodos False
        Case 4 'cuenta anterior
            Desplazamiento 0
        Case 5 'cuenta siguiente
            Desplazamiento 1
    End Select
End Sub

Private Sub Desplazamiento(Index As Integer)
    If data1.Recordset.EOF Then Exit Sub
    
    Select Case Index
        Case 0
            data1.Recordset.MovePrevious
            If data1.Recordset.BOF Then data1.Recordset.MoveFirst
            
        Case 1
            data1.Recordset.MoveNext
            If data1.Recordset.EOF Then data1.Recordset.MoveLast
    End Select
    txtCta(17).Text = data1.Recordset.Fields(0)
    DtxtCta(17).Text = DBLet(data1.Recordset.Fields(1), "T")
    PonerModoUsuarioGnral 0, "ariconta"
    PonerVtosCompensacionCliente
End Sub


Private Sub BotonVerTodos(Limpiar As Boolean)
Dim SQL As String
    'Ver todos
    
    VerTodos = True
    
    SQL = "select distinct cobros.codmacta, cobros.nomclien from cobros  where (1=1) "
    If Cuentas <> "" Then SQL = SQL & " and cobros.codmacta in (" & Cuentas & ")"
    SQL = SQL & " and impvenci + coalesce(gastos,0) - coalesce(impcobro,0) < 0"
    If Limpiar Then SQL = SQL & " and cobros.codmacta is null"
    
    
    
    If TotalRegistrosConsulta(SQL) = 0 Then
        If Not Limpiar Then MsgBox "No hay cuentas con abonos.", vbExclamation
        
        VerTodos = False
    End If
    
    
    data1.ConnectionString = Conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    data1.RecordSource = SQL
    data1.Refresh
    
    If VerTodos Then
        txtCta(17).Text = data1.Recordset.Fields(0)
        DtxtCta(17).Text = data1.Recordset.Fields(1)
    Else
        txtCta(17).Text = ""
        DtxtCta(17).Text = ""
    End If
    PonerModoUsuarioGnral 0, "ariconta"
    PonerVtosCompensacionCliente
    
End Sub

Private Sub BotonAnyadir()

    Frame1.visible = False
    Frame1.Enabled = False

    Me.FrameCompensaAbonosCliente.visible = True
    Me.FrameCompensaAbonosCliente.Enabled = True
    
    VerTodos = False
    
    PonleFoco txtCta(17)

End Sub


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(Index)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
    txtCta(Index).Text = Trim(txtCta(Index).Text)
    
     
     
    If txtCta(Index).Text = "" Then
        DtxtCta(Index).Text = ""
       ' txtCta(6).Tag = txtCta(6).Text
        Exit Sub
    End If
    
    If Not IsNumeric(txtCta(Index).Text) Then
        MsgBox "La cuenta debe ser numérica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        DtxtCta(Index).Text = ""
        txtCta(6).Tag = txtCta(6).Text
        PonFoco txtCta(Index)
        
        If Index = 17 Then PonerVtosCompensacionCliente
        
        Exit Sub
    End If
    
    Select Case Index
    Case Else
        'DE ULTIMO NIVEL
        Cta = (txtCta(Index).Text)
        If CuentaCorrectaUltimoNivel(Cta, SQL) Then
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = SQL
            
            
        Else
            MsgBox SQL, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
            txtCta(Index).SetFocus
        End If
        If Index = 17 Then PonerVtosCompensacionCliente
        
    End Select
End Sub

Private Function ComprobarCuentas(Indice1 As Integer, Indice2 As Integer) As Boolean
Dim L1 As Integer
Dim L2 As Integer
    ComprobarCuentas = False
    If txtCta(Indice1).Text <> "" And txtCta(Indice2).Text <> "" Then
        L1 = Len(txtCta(Indice1).Text)
        L2 = Len(txtCta(Indice2).Text)
        If L1 > L2 Then
            L2 = L1
        Else
            L1 = L2
        End If
        If Val(Mid(txtCta(Indice1).Text & "000000000", 1, L1)) > Val(Mid(txtCta(Indice2).Text & "0000000000", 1, L1)) Then
            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
            Exit Function
        End If
    End If
    ComprobarCuentas = True
End Function


Private Sub SubSetFocus(Obje As Object)
    On Error Resume Next
    Obje.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


'Si tiene valor el campo fecha, entonces lo ponemos con el BD
Private Function CampoABD(ByRef T As TextBox, Tipo As String, CampoEnLaBD, Mayor_o_Igual As Boolean) As String

    CampoABD = ""
    If T.Text <> "" Then
        If Mayor_o_Igual Then
            CampoABD = " >= "
        Else
            CampoABD = " <= "
        End If
        Select Case Tipo
        Case "F"
            CampoABD = CampoEnLaBD & CampoABD & "'" & Format(T.Text, FormatoFecha) & "'"
        Case "T"
            CampoABD = CampoEnLaBD & CampoABD & "'" & T.Text & "'"
        Case "N"
            CampoABD = CampoEnLaBD & CampoABD & T.Text
        End Select
    End If
End Function



Private Function CampoBD_A_SQL(ByRef C As ADODB.Field, Tipo As String, Nulo As Boolean) As String

    If IsNull(C) Then
        If Nulo Then
            CampoBD_A_SQL = "NULL"
        Else
            If Tipo = "T" Then
                CampoBD_A_SQL = "''"
            Else
                CampoBD_A_SQL = "0"
            End If
        End If

    Else
    
        Select Case Tipo
        Case "F"
            CampoBD_A_SQL = "'" & Format(C.Value, FormatoFecha) & "'"
        Case "T"
            CampoBD_A_SQL = "'" & DevNombreSQL(C.Value) & "'"
        Case "N"
            CampoBD_A_SQL = TransformaComasPuntos(CStr(C.Value))
        End Select
    End If
End Function

Private Sub PonerFrameProgressVisible(Optional TEXTO As String)
        If TEXTO = "" Then TEXTO = "Generando datos"
        Me.Refresh
End Sub


'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'
'       Compensaciones Cliente. Abonos vs Cobros
'
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

Private Sub PonerVtosCompensacionCliente()
Dim IT


    lwCompenCli.ListItems.Clear
    Me.txtimpNoEdit(0).Tag = 0
    Me.txtimpNoEdit(1).Tag = 0
    Me.txtimpNoEdit(0).Text = ""
    Me.txtimpNoEdit(1).Text = ""
    If Me.txtCta(17).Text = "" Then Exit Sub
    Set Me.lwCompenCli.SmallIcons = frmppal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    Cad = "Select cobros.*,nomforpa from cobros,formapago where cobros.codforpa=formapago.codforpa "
    Cad = Cad & " AND codmacta = '" & Me.txtCta(17).Text & "'"
    Cad = Cad & " AND (transfer =0 or transfer is null) and codrem is null"
    Cad = Cad & " and recedocu=0 and situacion = 0" ' pendientes de cobro
    Cad = Cad & " ORDER BY fecvenci"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCompenCli.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!NumFactu, "0000000")
        IT.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!numorden
        IT.SubItems(4) = miRsAux!FecVenci
        IT.SubItems(5) = miRsAux!nomforpa
    
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!ImpVenci
        
        'Si ya he cobrado algo
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        
        If Importe > 0 Then
            IT.SubItems(6) = Format(Importe, FormatoImporte)
            IT.SubItems(7) = " "
        Else
            IT.SubItems(6) = " "
            IT.SubItems(7) = Format(-Importe, FormatoImporte)
        End If
        IT.Tag = Abs(Importe)  'siempre valor absoluto
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



Private Sub RealizarCompensacionAbonosClientes()
Dim Borras As Boolean
    
    If BloqueoManual(True, "COMPEABONO", "1") Then

        Cad = DevuelveDesdeBD("max(codigo)", "compensa", "1", "1")
        If Cad = "" Then Cad = "0"
        CONT = Val(Cad) + 1 'ID de la operacion
        
         Observaciones = "Compensacion " & CONT & " fecha: " & Now & vbCrLf & Observaciones
        
        
        
        Cad = "INSERT INTO compensa(codigo,fecha,login,PC,codmacta,nommacta) VALUES (" & CONT
        Cad = Cad & ",now(),'" & DevNombreSQL(vUsu.Login) & "','" & DevNombreSQL(vUsu.PC)
        Cad = Cad & "','" & txtCta(17).Text & "','" & DevNombreSQL(DtxtCta(17).Text) & "')"
        
        Set miRsAux = New ADODB.Recordset
        Borras = True
        If Ejecuta(Cad) Then
            
            Borras = Not RealizarProcesoCompensacionAbonos
        
        End If


        Set miRsAux = Nothing
        If Borras Then
            Conn.Execute "DELETE FROM compensa WHERE codigo = " & CONT
            Conn.Execute "DELETE FROM compensa_facturas WHERE codigo = " & CONT
            
        End If

        'Desbloquamos proceso
        BloqueoManual False, "COMPEABONO", ""
        DevfrmCCtas = ""
        
        PonerVtosCompensacionCliente   'Volvemos a cargar los vencimientos
        
        'El nombre del report
        CadenaDesdeOtroForm = Me.Tag
        Me.Tag = ""
        If Not Borras Then
            Screen.MousePointer = vbDefault
            frmTESCompensaAboCliImp.pCodigo = CONT
            frmTESCompensaAboCliImp.Show vbModal
        End If
        
        Set miRsAux = Nothing
    Else
        MsgBox "Proceso bloqueado", vbExclamation
    End If

End Sub




Private Function RealizarProcesoCompensacionAbonos() As Boolean
Dim Destino As Byte
Dim J As Integer

    'NO USAR CONT

    RealizarProcesoCompensacionAbonos = False


    'Vamos a seleccionar los vtos
    '(numserie,codfaccl,fecfaccl,numorden)
    'EN SQL
    SQLVtosSeleccionadosCompensacion NumRegElim, False    'todos  -> Numregelim tendr el destino
    
    'Metemos los campos en el la tabla de lineas
    ' Esto guarda el valor en CAD
    FijaCadenaSQLCobrosCompen
    
    
    'Texto compensacion
    DevfrmCCtas = ""
    
    RC = "Select " & Cad & ", gastos, impvenci, impcobro, fecvenci FROM cobros where (numserie,numfactu,fecfactu,numorden) IN (" & SQL & ")"
    miRsAux.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Error. EOF vencimientos devueltos ", vbExclamation
        Exit Function
    End If
    
    
    i = 0
    
    While Not miRsAux.EOF
        i = i + 1
        BACKUP_Tabla miRsAux, RC
        'Quito los parentesis
        RC = Mid(RC, 1, Len(RC) - 1)
        RC = Mid(RC, 2)
        
        Destino = 0
        If miRsAux!NUmSerie = Me.lwCompenCli.ListItems(NumRegElim).Text Then
            If miRsAux!NumFactu = Val(Me.lwCompenCli.ListItems(NumRegElim).SubItems(1)) Then
                If Format(miRsAux!FecFactu, "dd/mm/yyyy") = Me.lwCompenCli.ListItems(NumRegElim).SubItems(2) Then
                    If miRsAux!numorden = Val(Me.lwCompenCli.ListItems(NumRegElim).SubItems(3)) Then Destino = 1
                End If
            End If
        End If
        
        RC = "INSERT INTO compensa_facturas (codigo,linea,destino," & Cad & ",impvenci,gastos,impcobro,fecvenci) VALUES (" & CONT & "," & i & "," & Destino & "," & DBSet(miRsAux!NUmSerie, "T")
        RC = RC & "," & DBSet(miRsAux!NumFactu, "N") & "," & DBSet(miRsAux!FecFactu, "F") & "," & DBSet(miRsAux!numorden, "N") & "," & DBSet(miRsAux!ImpVenci, "N")
        RC = RC & "," & DBSet(miRsAux!Gastos, "N") & "," & DBSet(miRsAux!impcobro, "N") & "," & DBSet(miRsAux!FecVenci, "F") & ")"
        Conn.Execute RC
        
        'Para las observaciones de despues
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!ImpVenci
        'Si ya he cobrado algo
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        
        If Destino = 0 Then 'El destino
            DevfrmCCtas = DevfrmCCtas & miRsAux!NUmSerie & Format(miRsAux!NumFactu, "0000000") & " " & Format(miRsAux!FecFactu, "dd/mm/yyyy")
            DevfrmCCtas = DevfrmCCtas & " Vto:" & Format(miRsAux!FecVenci, "dd/mm/yy") & " " & Importe
            DevfrmCCtas = DevfrmCCtas & "|"
        Else
            'El DESTINO siempre ira en la primera observacion del texto
            RC = "Importe anterior vto: " & Importe
            DevfrmCCtas = RC & "|" & DevfrmCCtas
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Acutalizaremos el VTO destino
    
    Conn.BeginTrans
        'Insertaremos registros en cobros_realizados BORRAREMOS LOS VENCIMIENTOS QUE NO SEAN DESTINO a no ser que el importe restante sea 0
        Destino = 1
        If txtimpNoEdit(0).Text = txtimpNoEdit(1).Text Then Destino = 0
        
        SQLVtosSeleccionadosCompensacion 0, Destino = 1  'sin o con el destino
        
        'Para saber si ha ido bien
        Destino = 0    '0 mal,1 bien
        If InsertarCobrosRealizados(SQL) Then
            If txtimpNoEdit(0).Text = txtimpNoEdit(1).Text Then
                Destino = 1
            Else
                'Updatearemos los campos csb del vto restante. A partir del segundo
                'La variable CadenaDesdeOtroForm  tiene los que vamos a actualizar
                
                Cad = ""
                J = 0
                SQL = ""
                
                
                Importe = CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag)  'txtimpNoEdit(1).Tag es negativo
                
                RC = "gastos=null, impcobro=null,fecultco=null,impvenci=" & TransformaComasPuntos(CStr(Importe))
                RC = RC & ",text33csb=" & DBSet(Txt33Csb, "T")
                RC = RC & ",text41csb=" & DBSet(Txt41Csb, "T")
                
                SQL = RC & SQL
                SQL = "UPDATE cobros SET " & SQL
                
                SQL = SQL & ", observa =concat(coalesce(observa,''),' '," & DBSet(Observaciones, "T") & ")"
                
                
                'WHERE
                RC = ""
                For J = 1 To Me.lwCompenCli.ListItems.Count
                    If Me.lwCompenCli.ListItems(J).Bold Then
                        'Este es el destino
                        RC = "NUmSerie = '" & Me.lwCompenCli.ListItems(J).Text
                        RC = RC & "' AND numfactu = " & Val(Me.lwCompenCli.ListItems(J).SubItems(1))
                        RC = RC & " AND fecfactu = '" & Format(Me.lwCompenCli.ListItems(J).SubItems(2), FormatoFecha)
                        RC = RC & "' AND numorden = " & Val(Me.lwCompenCli.ListItems(J).SubItems(3))
                        Exit For
                    End If
                Next
                If RC <> "" Then
                    Cad = SQL & " WHERE " & RC
                    If Ejecuta(Cad) Then Destino = 1
                Else
                    MsgBox "No encontrado destino", vbExclamation
                    
                End If
            End If
        End If
        If Destino = 1 Then
            Conn.CommitTrans
            RealizarProcesoCompensacionAbonos = True
        Else
            Conn.RollbackTrans
        End If
        
End Function

Private Function InsertarCobrosRealizados(facturas As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Cade As String
Dim NumLin As Long

    On Error GoTo eInsertarCobrosRealizados

    InsertarCobrosRealizados = True

    SQL = "select * from cobros where (numserie, numfactu, fecfactu, numorden) in (" & facturas & ")"
    
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        ' actualizamos la cabecera del cobro pq ya no lo eliminamos
        SQL = "update cobros set situacion = 1, impcobro = impvenci + coalesce(gastos,0),fecultco = " & DBSet(Now, "F")
        Cade = DBLet(Rs!observa, "T")
        If Cade <> "" Then Cade = Cade & vbCrLf
        Cade = Cade & Observaciones
        SQL = SQL & " , observa =" & DBSet(Cade, "T")
        SQL = SQL & " where numserie = " & DBSet(Rs!NUmSerie, "T")
        SQL = SQL & " and numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F") & " and numorden = " & DBSet(Rs!numorden, "N")
        
        
        
        
        
        Conn.Execute SQL
        
        Rs.MoveNext
    Wend
     
    Set Rs = Nothing
    Exit Function
    
eInsertarCobrosRealizados:
    InsertarCobrosRealizados = False
End Function




Private Sub SQLVtosSeleccionadosCompensacion(ByRef RegistroDestino As Long, SinDestino As Boolean)
Dim Insertar As Boolean
    SQL = ""
    For i = 1 To Me.lwCompenCli.ListItems.Count
        If Me.lwCompenCli.ListItems(i).Checked Then
        
            Insertar = True
            If Me.lwCompenCli.ListItems(i).Bold Then
                RegistroDestino = i
                If SinDestino Then Insertar = False
            End If
            If Insertar Then
                SQL = SQL & ", ('" & lwCompenCli.ListItems(i).Text & "'," & lwCompenCli.ListItems(i).SubItems(1)
                SQL = SQL & ",'" & Format(lwCompenCli.ListItems(i).SubItems(2), FormatoFecha) & "'," & lwCompenCli.ListItems(i).SubItems(3) & ")"
            End If
            
        End If
    Next
    SQL = Mid(SQL, 2)
            
End Sub


Private Sub FijaCadenaSQLCobrosCompen()

    Cad = "numserie, numfactu, fecfactu, numorden "
    
End Sub



Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N")
        Toolbar1.Buttons(2).Enabled = False 'DBLet(RS!Modificar, "N") And (Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(3).Enabled = False 'DBLet(RS!creareliminar, "N") And (Modo = 2) And DesdeNorma43 = 0
        
        Toolbar1.Buttons(5).Enabled = False 'DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N")
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And Modo = 2
    
        Toolbar2.Buttons(1).Enabled = True 'establecer cta
        Toolbar2.Buttons(2).Enabled = True 'ver todos
        Toolbar2.Buttons(4).Enabled = VerTodos
        Toolbar2.Buttons(5).Enabled = VerTodos
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub



Private Sub CargaList()
Dim IT

    lw1.ListItems.Clear
    Set Me.lw1.SmallIcons = frmppal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    Cad = "Select codigo,fecha,codmacta,nommacta from compensa "
    
    
    If CampoOrden = "" Then CampoOrden = "compensa.codigo"
    Cad = Cad & " ORDER BY " & CampoOrden
    If Orden Then Cad = Cad & " DESC"
    
    
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = miRsAux!Codigo
        IT.SubItems(1) = Format(miRsAux!Fecha, "dd/mm/yyyy hh:mm:ss")
        IT.SubItems(2) = miRsAux!codmacta
        IT.SubItems(3) = miRsAux!Nommacta
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

    If lw1.ListItems.Count > 0 Then
        Modo = 2
    Else
        Modo = 0
    End If
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
End Sub

