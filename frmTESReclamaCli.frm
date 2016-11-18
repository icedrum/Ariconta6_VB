VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESReclamaCli 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reclamaciones"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13050
   Icon            =   "frmTESReclamaCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   13050
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6930
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameReclamacionesCliente 
      BorderStyle     =   0  'None
      Height          =   7005
      Left            =   90
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   12735
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   3495
         Left            =   240
         TabIndex        =   24
         Top             =   2880
         Width           =   12405
         Begin MSComctlLib.ListView lwReclamCli 
            Height          =   3075
            Left            =   0
            TabIndex        =   25
            Top             =   330
            Width           =   12075
            _ExtentX        =   21299
            _ExtentY        =   5424
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
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Importe"
               Object.Width           =   3590
            EndProperty
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   11790
            Picture         =   "frmTESReclamaCli.frx":000C
            ToolTipText     =   "Puntear al Debe"
            Top             =   30
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   11430
            Picture         =   "frmTESReclamaCli.frx":0156
            ToolTipText     =   "Quitar al Debe"
            Top             =   30
            Width           =   240
         End
      End
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   300
         TabIndex        =   22
         Top             =   6390
         Width           =   1755
         Begin VB.Label lblIndicador 
            Alignment       =   2  'Center
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   210
            Width           =   1200
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FEF7E4&
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   0
         Tag             =   "Codigo|N|S|||reclama|codigo||S|"
         Top             =   480
         Width           =   1245
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frmTESReclamaCli.frx":02A0
         Left            =   10470
         List            =   "frmTESReclamaCli.frx":02A2
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Tipo Carta|N|N|0|2|reclama|carta|||"
         Top             =   540
         Width           =   1830
      End
      Begin VB.TextBox Text1 
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
         Height          =   360
         Index           =   2
         Left            =   270
         TabIndex        =   2
         Tag             =   "Cuenta|T|S|||reclama|codmacta|||"
         Top             =   1230
         Width           =   1245
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
         Height          =   840
         Index           =   0
         Left            =   270
         TabIndex        =   5
         Tag             =   "Observaciones|T|S|||reclama|observaciones|||"
         Top             =   1980
         Width           =   12045
      End
      Begin VB.TextBox Text1 
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
         Height          =   360
         Index           =   1
         Left            =   4410
         TabIndex        =   1
         Tag             =   "Fecha Reclamación|F|N|||reclama|fecreclama|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   480
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   10470
         TabIndex        =   4
         Tag             =   "Importe|N|N|||reclama|importes|||"
         Top             =   1200
         Width           =   1815
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
         Index           =   3
         Left            =   1590
         TabIndex        =   9
         Tag             =   "Cuenta|T|S|||reclama|nommacta|||"
         Text            =   "Text5"
         Top             =   1230
         Width           =   6645
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
         Index           =   0
         Left            =   10140
         TabIndex        =   6
         Top             =   6540
         Width           =   975
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
         Left            =   11340
         TabIndex        =   7
         Top             =   6540
         Width           =   975
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
      Begin VB.Image Image3 
         Height          =   240
         Index           =   0
         Left            =   1860
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   510
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "Carta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9390
         TabIndex        =   20
         Top             =   570
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
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
         Left            =   270
         TabIndex        =   19
         Top             =   1650
         Width           =   1440
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
         Index           =   5
         Left            =   3120
         TabIndex        =   18
         Top             =   510
         Width           =   795
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4080
         Picture         =   "frmTESReclamaCli.frx":02A4
         Top             =   510
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
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   72
         Left            =   9390
         TabIndex        =   11
         Top             =   1215
         Width           =   765
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
         TabIndex        =   10
         Top             =   930
         Width           =   1440
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   1770
         Top             =   930
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6885
      Left            =   90
      TabIndex        =   12
      Top             =   -30
      Visible         =   0   'False
      Width           =   12735
      Begin VB.Frame FrameBotonGnral2 
         Height          =   705
         Left            =   4020
         TabIndex        =   26
         Top             =   180
         Width           =   1095
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   330
            Left            =   210
            TabIndex        =   27
            Top             =   240
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
                  Object.ToolTipText     =   "Efectuar reclamacion "
               EndProperty
            EndProperty
         End
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
         Index           =   1
         Left            =   11550
         TabIndex        =   15
         Top             =   6300
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   240
         TabIndex        =   13
         Top             =   180
         Width           =   3585
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   180
            TabIndex        =   16
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
         TabIndex        =   14
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cuenta"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   5468
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Envio"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Importe"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Observac"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Codigo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Carta"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   12060
         TabIndex        =   17
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
Attribute VB_Name = "frmTESReclamaCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SaltoLinea = """ + chr(13) + """

Private Const IdPrograma = 608


    
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1

Dim Sql As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim cad As String
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
Dim Indice As Integer
Dim Codigo As Long

Private Sub PonFoco(ByRef T1 As TextBox)
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
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
        Frame1.Visible = True
        Frame1.Enabled = True
        
        FrameReclamacionesCliente.Visible = False
        FrameReclamacionesCliente.Enabled = False
        
        CargaList
        Codigo = ComprobarCero(Text1(5))
    Else
        Unload Me
    End If
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
    Select Case Index
        Case 0
            Select Case Modo
                Case 3  ' insertar
                    If DatosOK Then
                        Insertar
                        cmdCancelar_Click (0)
                    End If
                Case 4  ' modificar
                    If DatosOK Then
                        ModificaDesdeFormulario Me
                        cmdCancelar_Click (0)
                    End If
            End Select
    End Select
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean

    DatosOK = False

    'comprobamos datos OK de la tabla scafac
    B = CompForm2(Me, 2, "FrameReclamacionesCliente")
    Text1(5).BackColor = vbLightBlue '&HFEF7E4
    If Not B Then Exit Function
    
    DatosOK = B

End Function

Private Sub Insertar()
Dim NumF As Long
Dim B As Boolean

    On Error GoTo eInsertar
    
    Conn.BeginTrans
    
    NumF = SugerirCodigoSiguienteStr("reclama", "codigo")
    Text1(5).Text = NumF
    Codigo = Text1(5)
    B = InsertarDesdeForm(Me)
    If B Then InsertarLineas
    
eInsertar:
    If Err.Number = 0 And B Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
    End If
End Sub

Private Function InsertarLineas() As Boolean
Dim Rs As ADODB.Recordset
Dim CadValues As String
Dim CadInsert As String

    On Error GoTo eInsertarLineas

    InsertarLineas = False

    CadInsert = "insert into reclama_facturas (codigo,numlinea,numserie,numfactu,fecfactu,numorden,impvenci) values "

    CadValues = ""
    For i = 1 To lwReclamCli.ListItems.Count
        If lwReclamCli.SelectedItem.Checked Then
            CadValues = CadValues & "(" & DBSet(Text1(5).Text, "N") & "," & DBSet(i, "N") & "," & DBSet(lwReclamCli.ListItems(i).Text, "T") & ","
            CadValues = CadValues & DBSet(lwReclamCli.ListItems(i).SubItems(1), "N") & "," & DBSet(lwReclamCli.ListItems(i).SubItems(2), "F") & ","
            CadValues = CadValues & DBSet(lwReclamCli.ListItems(i).SubItems(3), "N") & "," & DBSet(lwReclamCli.ListItems(i).SubItems(6), "N") & "),"
        End If
    Next i
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute CadInsert & CadValues
    End If
    
    InsertarLineas = True
    Exit Function
    
eInsertarLineas:
    MuestraError Err.Number, "Insertar Lineas", Err.Description
End Function

Private Sub cmdVtoDestino(Index As Integer)
    
    If Index = 0 Then
        TotalRegistros = 0
        If Not Me.lwReclamCli.SelectedItem Is Nothing Then TotalRegistros = Me.lwReclamCli.SelectedItem.Index
    
    
        For i = 1 To Me.lwReclamCli.ListItems.Count
            If Me.lwReclamCli.ListItems(i).Bold Then
                Me.lwReclamCli.ListItems(i).Bold = False
                Me.lwReclamCli.ListItems(i).ForeColor = vbBlack
                For CONT = 1 To Me.lwReclamCli.ColumnHeaders.Count - 1
                    Me.lwReclamCli.ListItems(i).ListSubItems(CONT).ForeColor = vbBlack
                    Me.lwReclamCli.ListItems(i).ListSubItems(CONT).Bold = False
                Next
            End If
        Next
        Me.Refresh
        
        If TotalRegistros > 0 Then
            i = TotalRegistros
            Me.lwReclamCli.ListItems(i).Bold = True
            Me.lwReclamCli.ListItems(i).ForeColor = vbRed
            For CONT = 1 To Me.lwReclamCli.ColumnHeaders.Count - 1
                Me.lwReclamCli.ListItems(i).ListSubItems(CONT).ForeColor = vbRed
                Me.lwReclamCli.ListItems(i).ListSubItems(CONT).Bold = True
            Next
        End If
        lwReclamCli.Refresh
        
        PonerFocoLw Me.lwReclamCli

    Else
    
    End If
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Not Frame1.Visible Then
            If CadenaDesdeOtroForm <> "" Then
                Text1(2).Text = CadenaDesdeOtroForm
                Text1_LostFocus 2
            Else
                PonFoco Text1(2)
            End If
            CadenaDesdeOtroForm = ""
        End If
        CargaList
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
    
    ' Botonera Principal 2
    With Me.Toolbar2
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 47
    End With
    
    
    
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    
    'Limpiamos el tag
    PrimeraVez = True
    CommitConexion  'Porque son listados. No hay nada dentro transaccion
    
        
    H = FrameReclamacionesCliente.Height + 120
    W = FrameReclamacionesCliente.Width
    
    FrameReclamacionesCliente.Visible = False
    Me.Frame1.Visible = True
    
    VerTodos = False
    
    Me.Width = W + 300
    Me.Height = H + 400
    
    Me.cmdCancelar(0).Cancel = True
    
    PonerModoUsuarioGnral 0, "ariconta"
    
    Orden = True
    
    CargaCombo

End Sub


Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    Text1(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub Image3_Click(Index As Integer)

    Select Case Index
        Case 1 ' cuenta contable
            Screen.MousePointer = vbHourglass
            
            Set frmCta = New frmColCtas
            RC = Index + 1
            frmCta.DatosADevolverBusqueda = "0|1"
            frmCta.ConfigurarBalances = 3
            frmCta.Show vbModal
            Set frmCta = Nothing
            If Index = 2 Then PonerVtosReclamacionCliente False
    
        Case 0 ' observaciones
            Screen.MousePointer = vbDefault
            
            Indice = 0
            
            Set frmZ = New frmZoom
            frmZ.pValor = Text1(Indice).Text
            frmZ.pModo = Modo
            frmZ.Caption = "Observaciones Reclamaciones Cliente"
            frmZ.Show vbModal
            Set frmZ = Nothing
            
    End Select
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim IT
Dim i As Integer
    For i = 1 To Me.lwReclamCli.ListItems.Count
        Set IT = lwReclamCli.ListItems(i)
        lwReclamCli.ListItems(i).Checked = (Index = 1)
        lwReclamCli_ItemCheck (IT)
        Set IT = Nothing
    Next i
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)
    'FECHA FACTURA
    Indice = 1
    
    Set frmF = New frmCal
    frmF.Fecha = Now
    If Text1(1).Text <> "" Then frmF.Fecha = CDate(Text1(1).Text)
    frmF.Show vbModal
    Set frmF = Nothing
    PonFoco Text1(1)

End Sub

Private Sub lw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Campo2 As Integer

    Orden = Not Orden
    
    Select Case ColumnHeader
        Case "Código"
            CampoOrden = "reclama.codigo"
        Case "Fecha"
            CampoOrden = "reclama.fecreclama"
        Case "Cuenta"
            CampoOrden = "reclama.codmacta"
        Case "Nombre"
            CampoOrden = "reclama.nommacta"
        Case "Carta"
            CampoOrden = "reclama.carta"
    End Select
    CargaList


End Sub

Private Sub lw1_DblClick()
    'detalle de facturas
    Set frmMens = New frmMensajes
    
    frmMens.Opcion = 50
    frmMens.Parametros = lw1.SelectedItem.SubItems(6) & "|" & lw1.SelectedItem.SubItems(2) & "|" & lw1.SelectedItem.Text & "|"
    frmMens.Show vbModal
    
    Set frmMens = Nothing

End Sub

Private Sub lwReclamCli_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim C As Currency
Dim Cobro As Boolean

    Cobro = True
    C = Item.Tag
    
    Importe = 0
    For i = 1 To lwReclamCli.ListItems.Count
        If lwReclamCli.ListItems(i).Checked Then Importe = Importe + lwReclamCli.ListItems(i).SubItems(6)
    Next i
    Text1(4).Text = Format(Importe, "###,###,##0.00")
    
    If ComprobarCero(Text1(4).Text) = 0 Then Text1(4).Text = ""
            
End Sub

Private Sub HacerToolBar(Boton As Integer)

    Select Case Boton
        Case 1
            BotonAnyadir
        Case 2
            BotonModificar
        Case 3
            BotonEliminar
        Case 5
'            BotonBuscar
        Case 6 ' ver todos
            CargaList
        Case 8
            'Imprimir factura
            frmTESReclamaCliList.Show vbModal

    End Select
End Sub

Private Sub BotonEliminar()
Dim Sql As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If Me.lw1.SelectedItem = "" Then Exit Sub
        
    '*************** canviar els noms i el DELETE **********************************
    Sql = "¿Seguro que desea eliminar la Reclamación?"
    Sql = Sql & vbCrLf & "Código: " & lw1.SelectedItem.SubItems(6)
    Sql = Sql & vbCrLf & " de fecha: " & lw1.SelectedItem.Text
    Sql = Sql & vbCrLf & " de " & lw1.SelectedItem.SubItems(1) & "-" & lw1.SelectedItem.SubItems(2)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = lw1.SelectedItem.SubItems(6)
        
        Sql = "Delete from reclama_facturas where codigo=" & lw1.SelectedItem.SubItems(6)
        Conn.Execute Sql
        
        Sql = "Delete from reclama where codigo=" & lw1.SelectedItem.SubItems(6)
        Conn.Execute Sql
        
        
        lw1.ListItems.Remove (lw1.SelectedItem.Index)
        If lw1.ListItems.Count > 0 Then
            lw1.SetFocus
        End If
        
'        CargaList
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub

Private Sub Desplazamiento(Index As Integer)
    If Data1.Recordset.EOF Then Exit Sub
    
    Select Case Index
        Case 0
            Data1.Recordset.MovePrevious
            If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
            
        Case 1
            Data1.Recordset.MoveNext
            If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
    End Select
    Text1(2).Text = Data1.Recordset.Fields(0)
    Text1(3).Text = Data1.Recordset.Fields(1)
    PonerModoUsuarioGnral 0, "ariconta"
    PonerVtosReclamacionCliente False
End Sub





Private Sub BotonAnyadir()

    Frame1.Visible = False
    Frame1.Enabled = False

    Me.FrameReclamacionesCliente.Visible = True
    Me.FrameReclamacionesCliente.Enabled = True
    
    VerTodos = False
    
    LimpiarCampos
    
    Combo1.ListIndex = 0
    
    Modo = 3
    PonerModo Modo
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    PonleFoco Text1(1)

End Sub

Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    Limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    Combo1.ListIndex = -1
    
    Me.lwReclamCli.ListItems.Clear
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub BotonModificar()

    If lw1.SelectedItem = "" Then Exit Sub

    Frame1.Visible = False
    Frame1.Enabled = False

    Me.FrameReclamacionesCliente.Visible = True
    Me.FrameReclamacionesCliente.Enabled = True
    
    VerTodos = False
    
    Modo = 4
    PonerModo Modo
    
    Text1(5).Text = lw1.SelectedItem.SubItems(6)
    Text1(1).Text = lw1.SelectedItem.Text
    Text1(2).Text = lw1.SelectedItem.SubItems(1)
    Text1(3).Text = lw1.SelectedItem.SubItems(2)
    PosicionarCombo Combo1, lw1.SelectedItem.SubItems(7)
    Text1(4).Text = lw1.SelectedItem.SubItems(4)
    Text1(0).Text = lw1.SelectedItem.SubItems(5)
    
    PonerVtosReclamacionCliente True
    
    PonleFoco Text1(1)
End Sub




Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    
    PonerIndicador lblIndicador, Modo
    
    ' la cuenta no se puede modificar pq cambiarian las líneas
    Text1(2).Locked = (Modo = 4)
    Text1(3).Locked = (Modo = 4)
    Image3(1).Visible = (Modo = 3)
    Image3(1).Enabled = (Modo = 3)
    Me.Frame4.Enabled = (Modo = 3)
    
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar2 Button.Index
End Sub

Private Sub HacerToolBar2(Boton As Integer)
    Select Case Boton
        Case 1
            frmTESReclamaCliEfe.Show vbModal
            CargaList
            
    End Select
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    PonFoco Text1(Index)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
    Text1(Index).Text = Trim(Text1(Index).Text)
    
     
    If Text1(Index).Text = "" Then
        Exit Sub
    End If
    
    Select Case Index
        Case 1 ' fecha
            PonerFormatoFecha Text1(Index)
        
        Case 2 ' cuenta
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "La cuenta debe ser numérica: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
                Text1(3).Text = ""
                Text1(6).Tag = Text1(6).Text
                PonFoco Text1(Index)
                
                If Modo = 3 Then PonerVtosReclamacionCliente False
                
                Exit Sub
            End If
            
            Select Case Index
            Case Else
                'DE ULTIMO NIVEL
                Cta = (Text1(Index).Text)
                If CuentaCorrectaUltimoNivel(Cta, Sql) Then
                    Text1(Index).Text = Cta
                    Text1(3).Text = Sql
                    
                    
                Else
                    MsgBox Sql, vbExclamation
                    Text1(Index).Text = ""
                    Text1(3).Text = ""
                    Text1(Index).SetFocus
                End If
                If Modo = 3 Then PonerVtosReclamacionCliente False
                
            End Select
        Case 4
            PonerFormatoDecimal Text1(Index), 1
    End Select
End Sub

Private Function ComprobarCuentas(Indice1 As Integer, Indice2 As Integer) As Boolean
Dim L1 As Integer
Dim L2 As Integer
    ComprobarCuentas = False
    If Text1(Indice1).Text <> "" And Text1(Indice2).Text <> "" Then
        L1 = Len(Text1(Indice1).Text)
        L2 = Len(Text1(Indice2).Text)
        If L1 > L2 Then
            L2 = L1
        Else
            L1 = L2
        End If
        If Val(Mid(Text1(Indice1).Text & "000000000", 1, L1)) > Val(Mid(Text1(Indice2).Text & "0000000000", 1, L1)) Then
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


Private Sub PonerVtosReclamacionCliente(Modificar As Boolean)
Dim IT
Dim ImporteTot As Currency

    lwReclamCli.ListItems.Clear
    If Not Modificar Then Text1(4).Text = ""
    
    ImporteTot = 0
    
    If Me.Text1(2).Text = "" Then Exit Sub
    Set Me.lwReclamCli.SmallIcons = frmppal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    If Modificar Then
        cad = "Select reclama_facturas.numlinea,reclama_facturas.numserie,reclama_facturas.numfactu,reclama_facturas.fecfactu,reclama_facturas.numorden,reclama_facturas.impvenci importe,"
        cad = cad & " cobros.codforpa,cobros.fecvenci, cobros.gastos, cobros.impvenci, cobros.impcobro,nomforpa from reclama_facturas,cobros,formapago where cobros.codforpa=formapago.codforpa "
        cad = cad & " and reclama_facturas.numserie = cobros.numserie "
        cad = cad & " and reclama_facturas.numfactu = cobros.numfactu "
        cad = cad & " and reclama_facturas.fecfactu = cobros.fecfactu "
        cad = cad & " and reclama_facturas.numorden = cobros.numorden "
        cad = cad & " AND reclama_facturas.codigo = " & Me.Text1(5).Text
        cad = cad & " ORDER BY 1"
    Else
        cad = "Select cobros.*,nomforpa from cobros,formapago where cobros.codforpa=formapago.codforpa "
        cad = cad & " AND codmacta = '" & Me.Text1(2).Text & "'"
        cad = cad & " AND (transfer =0 or transfer is null) and codrem is null"
        cad = cad & " and recedocu=0 and situacion = 0" ' pendientes de cobro
        cad = cad & " ORDER BY fecvenci"
    End If
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwReclamCli.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!NumFactu, "0000000")
        IT.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!numorden
        IT.SubItems(4) = miRsAux!FecVenci
        IT.SubItems(5) = miRsAux!nomforpa
    
        If Modificar Then
            IT.SubItems(6) = Format(DBLet(miRsAux!Importe, "N"), FormatoImporte)
            
            IT.Checked = True
        
            IT.Tag = DBLet(miRsAux!Importe, "N")  'siempre valor absoluto
        
        Else
            Importe = DBLet(miRsAux!Gastos, "N")
            Importe = Importe + miRsAux!ImpVenci
            
            'Si ya he cobrado algo
            If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
            
            IT.SubItems(6) = Format(Importe, FormatoImporte)
            
            ImporteTot = ImporteTot + Importe

            IT.Tag = Abs(Importe)  'siempre valor absoluto
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    

End Sub



Private Sub SQLVtosSeleccionadosCompensacion(ByRef RegistroDestino As Long, SinDestino As Boolean)
Dim Insertar As Boolean
    Sql = ""
    For i = 1 To Me.lwReclamCli.ListItems.Count
        If Me.lwReclamCli.ListItems(i).Checked Then
        
            Insertar = True
            If Me.lwReclamCli.ListItems(i).Bold Then
                RegistroDestino = i
                If SinDestino Then Insertar = False
            End If
            If Insertar Then
                Sql = Sql & ", ('" & lwReclamCli.ListItems(i).Text & "'," & lwReclamCli.ListItems(i).SubItems(1)
                Sql = Sql & ",'" & Format(lwReclamCli.ListItems(i).SubItems(2), FormatoFecha) & "'," & lwReclamCli.ListItems(i).SubItems(3) & ")"
            End If
            
        End If
    Next
    Sql = Mid(Sql, 2)
            
End Sub


Private Sub FijaCadenaSQLCobrosCompen()

    cad = "numserie, numfactu, fecfactu, numorden "
    
End Sub



Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N")
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = False 'DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N")
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And Modo = 2
    
        Toolbar2.Buttons(1).Enabled = DBLet(Rs!especial, "N")
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub



Private Sub CargaList()
Dim IT

    lw1.ListItems.Clear
    Set Me.lw1.SmallIcons = frmppal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    cad = "Select codigo,fecreclama,codmacta,nommacta,carta,CASE carta WHEN 0 THEN 'Carta' WHEN 1 THEN 'EMail' WHEN 2 THEN 'Teléfono' END as TCarta,importes,observaciones from reclama "
    
    
    If CampoOrden = "" Then CampoOrden = "reclama.fecreclama"
    cad = cad & " ORDER BY " & CampoOrden
    If Orden Then cad = cad & " DESC"
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = Format(miRsAux!Fecreclama, "dd/mm/yyyy")
        IT.SubItems(1) = DBLet(miRsAux!codmacta, "T")
        IT.SubItems(2) = DBLet(miRsAux!Nommacta, "T")
        IT.SubItems(3) = miRsAux!tcarta
        IT.SubItems(4) = Format(miRsAux!Importes, "###,###,##0.00")
        IT.SubItems(5) = DBLet(miRsAux!observaciones, "T")
        IT.SubItems(6) = miRsAux!Codigo
        IT.SubItems(7) = miRsAux!carta
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

Private Sub CargaCombo()
    Combo1.Clear
    Combo1.AddItem "Carta"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Email"
    Combo1.ItemData(Combo1.NewIndex) = 1
    Combo1.AddItem "Teléfono"
    Combo1.ItemData(Combo1.NewIndex) = 2
    
End Sub

