VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTESRemesas2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remesas"
   ClientHeight    =   7620
   ClientLeft      =   450
   ClientTop       =   525
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmdCta1 
         Caption         =   "+"
         Height          =   290
         Left            =   1140
         TabIndex        =   15
         Top             =   3480
         Width           =   195
      End
      Begin VB.CommandButton cmdDivideImporte 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   3
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdDivideImporte 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   2
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdBanco 
         Height          =   375
         Index           =   2
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Eliminar"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdBanco 
         Height          =   375
         Index           =   1
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Modificar"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdBanco 
         Height          =   375
         Index           =   0
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Insertar"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "Dat"
         Top             =   3480
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "Dat"
         Top             =   3480
         Width           =   800
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Index           =   1
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   16
         Text            =   "Dato2"
         Top             =   3480
         Width           =   1395
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2535
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   120
         Top             =   4680
         Width           =   1200
         _ExtentX        =   2117
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
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   3480
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha remesa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   9
         Left            =   3480
         TabIndex        =   19
         Top             =   1920
         Width           =   660
      End
      Begin VB.Label Label5 
         Caption         =   "Dividir importe bancos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame frame3 
      Height          =   3735
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6495
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         TabIndex        =   29
         Top             =   3225
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   4080
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   240
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   1935
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin VB.Label Label4 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   25
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Nº"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   24
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "RESUMEN BANCO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   2100
      End
      Begin VB.Label Label2 
         Caption         =   "REMESA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7095
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8775
      Begin MSComctlLib.ListView ListView1 
         Height          =   6255
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   11033
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha Vto."
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nº"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Banco"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "FechaOrden"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ImporteOrden"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Fecfaccl"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblImporteRemesa 
         Alignment       =   1  'Right Justify
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   28
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "DIVIDIR REMESAS ENTRE BANCOS"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame FrameDevolucionRemesa 
      Height          =   7575
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton cmdDevRem 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   7440
         TabIndex        =   35
         Top             =   7080
         Width           =   1095
      End
      Begin VB.CommandButton cmdDevRem 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   6120
         TabIndex        =   34
         Top             =   7080
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView22 
         Height          =   6255
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   11033
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Serie"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Factura"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Vto"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cta"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "fechaorden"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Importeorden"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   36
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Devolución de remesas: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   4
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   5730
      End
   End
   Begin VB.Menu mnPopUp 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnQuitarRecibo 
         Caption         =   "Quitar"
      End
      Begin VB.Menu mnBarra 
         Caption         =   "-"
      End
      Begin VB.Menu mnBanco 
         Caption         =   "b1"
         Index           =   1
      End
      Begin VB.Menu mnBanco 
         Caption         =   "b2"
         Index           =   2
      End
      Begin VB.Menu mnBanco 
         Caption         =   "b3"
         Index           =   3
      End
      Begin VB.Menu mnBanco 
         Caption         =   "b4"
         Index           =   4
      End
      Begin VB.Menu mnBanco 
         Caption         =   "b5"
         Index           =   5
      End
      Begin VB.Menu mnBanco 
         Caption         =   "b6"
         Index           =   6
      End
      Begin VB.Menu mnBanco 
         Caption         =   "b7"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmTESRemesas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vSQL As String
Public Opcion As Byte      ' 0.- Nueva remesa    1.- Modifcar remesa
                           ' 2.- Devolucion remesa
Public vRemesa As String   ' nºremesa|fecha remesa
Public ImporteRemesa As Currency

Public ValoresDevolucionRemesa As String
        'NOV 2009
        'antes: 4 campos     AHORA 5 campos
        'Concepto|ampliacion|
        'Concepto banco|ampliacion banco|
        'ahora+ Agrupa vtos
        
Private WithEvents frmBan As frmBanco
Attribute frmBan.VB_VarHelpID = -1

Private Estado As Byte
    '0.- Datos basicos
    '1.- Separar por bancos
    '2.- Asignar a bancos
    
    
Dim Cad As String
Dim jj As Integer
Dim PrimeraVez As Boolean
Dim ImporteQueda As Currency
Dim ImporteModificando As Currency



Private Sub cmdBanco_Click(Index As Integer)
    Select Case Index
    Case 0
        If ImporteQueda <> 0 Then
            cmdBanco(0).Tag = 0
            PonerBotoncitos False
            AnyadirModificarBanco True
        End If
    Case 1
        If Not Adodc1.Recordset.EOF Then
            cmdBanco(0).Tag = 1
            PonerBotoncitos False
            'Modificar
            AnyadirModificarBanco False
        End If
    Case 2
        If Not Adodc1.Recordset.EOF Then
        
            
            Cad = RecuperaValor(vRemesa, 4)
            If Cad <> "" Then
                If Adodc1.Recordset!Cta = Cad Then
                    MsgBox "No puede eliminar el banco por defecto", vbExclamation
                    Exit Sub
                End If
            End If
        
        
            'Eliminamos
            ImporteQueda = ImporteQueda + Adodc1.Recordset!acumperd
            Adodc1.Recordset.Delete
            CargaGrid1 False
        End If
        
    End Select
    
End Sub

Private Sub cmdCta1_Click()
        Set frmBan = New frmBanco
        frmBan.DatosADevolverBusqueda = "0|1|"
        frmBan.Show vbModal
        Set frmBan = Nothing
End Sub

Private Sub cmdDevRem_Click(Index As Integer)
Dim AUX As String
        
        
        If Index = 0 Then
            'Vemos si ha selecionado alguno
            Cad = ""
            For jj = 1 To ListView22.ListItems.Count
                If ListView22.ListItems(jj).Checked Then
                    Cad = Cad & "1"
                End If
            Next jj
            If Cad = "" Then
                MsgBox "Seleccione los efectos devueltos", vbExclamation
                Exit Sub
            End If
            Cad = Len(Cad) & " efecto(s)."
            
            'Llegado aqui hago la pregunta
            Cad = "Va a realizar la devolución de " & Cad & vbCrLf
            
            If InStr(1, Label6.Caption, ":") > 0 Then
            
                Cad = Cad & vbCrLf & "Importe total de la devolución: "
                Cad = Cad & Mid(Label6.Caption, InStr(1, Label6.Caption, ":")) & "€" & vbCrLf & vbCrLf
            End If
            
            AUX = RecuperaValor(vRemesa, 5)
            If AUX = "%" Then
                AUX = "Porcentaje por recibo: " & ImporteRemesa & "%" & vbCrLf
                If RecuperaValor(vRemesa, 6) <> "" Then
                    AUX = AUX & "Gasto mínimo: " & RecuperaValor(vRemesa, 6) & " €" & vbCrLf
                End If
            Else
                AUX = "Gasto por recibo: " & ImporteRemesa & " €" & vbCrLf
            End If
            
            Cad = Cad & AUX & vbCrLf
            
            'Gasto tramitacion devolucion
            AUX = RecuperaValor(vRemesa, 7)
            If AUX <> "" Then
                AUX = "Gasto bancario : " & AUX & "€" & vbCrLf
                Cad = Cad & vbCrLf & AUX
            End If
            
            Cad = Cad & vbCrLf & "¿Desea continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            
            If Not RealizarDevolucion Then Exit Sub

            Unload Me
        Else
            Unload Me
        End If
     
End Sub

Private Sub cmdDivideImporte_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    DivideImporte Index
    Screen.MousePointer = vbDefault
End Sub

Private Sub DivideImporte(Index As Integer)
Dim Importe As Currency
    
    If Index = 1 Then
        LLamaLineas 0, 0
        DataGrid1.AllowAddNew = False
        CargaGrid1 False
        PonerBotoncitos True
        If ImporteQueda = 0 Then
            Command1(0).Enabled = True
        Else
            If Adodc1.Recordset.EOF Then
                Command1(0).Enabled = False

            Else
               HabilitaSeguir
            End If
        End If
    Else

        'Los campos no pueden ser ""
        If txtAux1(0).Text = "" Or txtAux1(1).Text = "" Or txtAux1(2).Text = "" Then
            MsgBox "Los campos no puden ser nulos", vbExclamation
            Exit Sub
        End If
        
        
        Importe = CCur(txtAux1(2).Text)
        If Importe > ImporteQueda Then
            MsgBox "Importe mayor del perimitido", vbExclamation
            Exit Sub
        End If
        If CBool(txtAux1(0).Tag) Then
            'INSERTAR
            Cad = "INSERT INTO tmpcierre1 (codusu, cta, nomcta, acumPerD) VALUES (" & vUsu.Codigo
            Cad = Cad & ",'" & txtAux1(0).Text & "','" & txtAux1(1).Text & "'," & TransformaComasPuntos(txtAux1(2).Text) & ")"
            If Not Ejecuta(Cad) Then Exit Sub
            
        Else
            'Modificar
            Cad = "UPDATE tmpcierre1 SET acumperd=" & TransformaComasPuntos(txtAux1(2).Text)
            Cad = Cad & " where codusu = " & vUsu.Codigo
            Cad = Cad & " AND cta='" & txtAux1(0).Text & "'"
            If Not Ejecuta(Cad) Then Exit Sub
            
        End If
        espera 0.5
        DataGrid1.AllowAddNew = False
        CargaGrid1 False
        If ImporteQueda > 0 Then
            'Añadimos otro
            espera 0.5
            AnyadirModificarBanco True
        Else
            
            PonerBotoncitos True
            LLamaLineas 0, 0
            HabilitaSeguir
        End If

    End If
End Sub


Private Sub HabilitaSeguir()
    Me.Command1(0).Enabled = True
    Me.Command1(0).SetFocus
End Sub


Private Sub Command1_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    'Haremos comprobaciones
    If Estado < 2 Then
        If Estado = 0 Then
        
            'Primera vez que comprobamos el riesgo con el banco
            If ComprobarRiesgoBanco Then
                
            
                'ya tenemos cuanto va a cada banco
                If RecuperaValor(vRemesa, 4) <> "" Then
                    MostrarVencimientosRemesasPorEntidad
                Else
                    'NORMAL
                    MostrarVencimientosRemesasNormal
                End If
            Else
                Exit Sub
            End If
                
        Else
            Cad = ""
            For jj = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(jj).SubItems(6) <> "" Then
                    Cad = "OK"
                    Exit For
                End If
            Next jj
            
            If Cad = "" Then
                MsgBox "Seleccione algun vencimiento", vbExclamation
                Exit Sub
            End If
            
            If Not ComprobarRiesgoAsignandoVto Then Exit Sub
            
            
            'Ponemos datos remesa
            Text1(0).Text = RecuperaValor(vRemesa, 1)
            Text1(1).Text = RecuperaValor(vRemesa, 2)
            Text1(2).Text = RecuperaValor(vRemesa, 3)
            
            
            Cad = "Select descripcion from remesas where codigo=" & Text1(0).Text & " AND anyo =" & Text1(1).Text
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = ""
            If Not miRsAux.EOF Then Cad = DBLet(miRsAux.Fields(0), "T")
            miRsAux.Close
            Set miRsAux = Nothing
            Text2.Text = Cad
            
            'Reajustamos los importes k iran a cada banco
            RejusteImportesBanco
            
        End If
    
        'Siguiente
        Estado = Estado + 1
        PonerEstado
    Else
        'Remesas
        If Opcion = 0 Then
            Cad = "Seguro que desea generar la remesa?"
        Else
            Cad = "Seguro que desea guardar los cambios de la remesa?"
        End If
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        GenerarRemesa
        Unload Me
    End If
    
End Sub




Private Sub Tamañosydemas(ByRef F As frame)
Dim h As Integer
Dim W As Integer

    h = F.Height + 120
    W = F.Width + 120
    
    If F.Name = "FrameDevolucionRemesa" Then
        Me.Width = W
        Me.Height = h + 240
        Exit Sub
    End If
    
    Me.Command1(0).Top = h
    Me.Command1(1).Top = Me.Command1(0).Top
    
    Me.Command1(1).Left = W - Me.Command1(1).Width - 120
    Me.Command1(0).Left = Me.Command1(1).Left - Me.Command1(0).Width - 120
    
    
    Me.Width = W
    
    
    If W > 8000 Then
        Me.Top = 2000
        Me.Left = 3500
    Else
        Me.Top = 3195
        Me.Left = 4185
    End If
    'mas los botones
    h = h + Me.Command1(0).Height + 520
    Me.Height = h
    
    Me.Refresh
End Sub


Private Sub PonerEstado()

    
    Frame1.Visible = Estado = 0
    Frame2.Visible = Estado = 1
    frame3.Visible = Estado = 2
        
        
    Label4.Visible = Estado <> 0
    Text2.Visible = Estado <> 0
    
    If Estado < 2 Then
        If Estado = 0 Then
            Tamañosydemas Frame1
            Me.Command1(0).Enabled = False
        Else
            Tamañosydemas Frame2
        End If
    
        Me.Command1(0).Caption = "Siguiente"
    Else
        Tamañosydemas frame3
        Me.Command1(0).Caption = "Remesar"


    End If
    
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        lblImporteRemesa.Visible = False
        
        If Opcion = 2 Then
            ImporteQueda = 0
            CargaDevolucion
            
            Label6.Caption = "Devuelto: " & Format(ImporteQueda, FormatoImporte)
            cmdDevRem(1).SetFocus
            
        Else
            If Estado = 0 Then
                PonerBotoncitos False
                
                'El 4 es la opcion de remesar por entidad.
                ' = ""   LO NORMAL
                ' <>
                Cad = Trim(RecuperaValor(vRemesa, 4))
                
                If Cad = "" Then
                    AnyadirModificarBanco True  'esto seria como estaba
                Else
                    
                    
                    
                    cmdBanco(2).Enabled = True
                    cmdDivideImporte(0).Enabled = False
                    cmdDivideImporte(1).Enabled = False

                    HabilitaSeguir
                    
                End If
                    
            Else
                lblImporteRemesa.Caption = "importe: " & Format(ImporteModificando, FormatoImporte)
                lblImporteRemesa.Visible = True
                'Añadir / eliminar vencimientos a la remesa
                Me.Refresh
                espera 0.1
                CargaGrid1 True
                MostrarVencimientosRemesasNormal
            End If
            Label4.Visible = Estado <> 0
            Text2.Visible = Estado <> 0
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    Me.lblImporteRemesa.Caption = ""
    
   
    Caption = "REMESAS      "
    
    If Opcion = 2 Then
        Limpiar Me
       
        'DEVOLUCION REMESA
        FrameDevolucionRemesa.Visible = True
        Tamañosydemas FrameDevolucionRemesa
        
        Frame1.Visible = False
        Frame2.Visible = False
        frame3.Visible = False
        
        Command1(0).Visible = False
        Command1(1).Visible = False
        CargaIconoListview ListView22
        
    Else
        CargaIconoListview ListView1   'EL FONDO
        FrameDevolucionRemesa.Visible = False



        If Opcion = 0 Then
            Estado = 0
            Me.txtAux1(0).Visible = False
            Me.txtAux1(1).Visible = False
            Me.cmdCta1.Visible = False
            
            'Cargamos GRID1
            Set DataGrid1.DataSource = Me.Adodc1
            CargaGrid1 False
                
            'Lable 3
            Label3.Caption = "Importe total a remesar:   " & Format(ImporteRemesa, FormatoImporte)
        
        Else
            
            Estado = 1
            
        End If
        PonerEstado
    End If
    
    
    
End Sub

Private Sub PonerBotoncitos(Ena As Boolean)
    Me.cmdBanco(0).Enabled = Ena
    Me.cmdBanco(1).Enabled = Ena
    Me.cmdBanco(2).Enabled = Ena
    cmdDivideImporte(0).Enabled = Not Ena
    cmdDivideImporte(1).Enabled = Not Ena
End Sub

Private Sub CargaGrid1(SoloDatosBanco As Boolean)
Dim SQL As String
Dim I As Integer
    
    espera 0.5
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = "Select cta,nomcta,acumperd from tmpcierre1 where codusu =" & vUsu.Codigo
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockOptimistic
    Adodc1.Refresh
    If SoloDatosBanco Then Exit Sub
    ImporteQueda = 0
    If Not Adodc1.Recordset.EOF Then
        While Not Adodc1.Recordset.EOF
            ImporteQueda = ImporteQueda + Adodc1.Recordset.Fields(2)
            Adodc1.Recordset.MoveNext
        Wend
        Adodc1.Recordset.MoveFirst
    End If
    ImporteQueda = ImporteRemesa - ImporteQueda
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    
    'Cuenta contable
    I = 0
        DataGrid1.Columns(I).Caption = "Cuenta"
        DataGrid1.Columns(I).Width = 1100
    
    'Descripcion NOMMACTA
    I = 1
        DataGrid1.Columns(I).Caption = "Descripción"
        DataGrid1.Columns(I).Width = 3400
 
    
    'Entidad
    I = 2
        DataGrid1.Columns(I).Caption = "Importe"
        DataGrid1.Columns(I).Width = 1200
        
        
    
        
    For I = 0 To 2
        DataGrid1.Columns(I).AllowSizing = False
    Next I
        
           
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux1(0).Left = DataGrid1.Left + 340
        txtAux1(0).Width = DataGrid1.Columns(0).Width - 120
        
        
        txtAux1(1).Left = txtAux1(0).Left + txtAux1(0).Width + cmdCta1.Width - 60
        txtAux1(1).Width = DataGrid1.Columns(1).Width - 60
        
        
        txtAux1(2).Width = DataGrid1.Columns(2).Width - 60
        txtAux1(2).Left = txtAux1(1).Left + txtAux1(1).Width + 30
        
        
        
        
        'El botoncito para la cuenta
        cmdCta1.Left = txtAux1(1).Left - 240
    
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
    
    End If
        
End Sub



Private Sub AnyadirModificarBanco(Anadir As Boolean)
Dim J As Integer
Dim anc As Single

    txtAux1(0).Tag = Anadir
    If Anadir Then
        DataGrid1.AllowAddNew = True
        If Not Adodc1.Recordset.EOF Then
            DataGrid1.HoldFields
            Adodc1.Recordset.MoveLast
            'DataGrid1.Row = DataGrid1.Row + 1
            J = DataGrid1.Row + 1
        Else
            J = -1
        End If
        
    Else
        DataGrid1.AllowAddNew = False
        J = DataGrid1.Row
    End If
   
    If J < 0 Then
        anc = DataGrid1.RowTop(0)
        Else
        anc = DataGrid1.RowTop(J)
    End If
    anc = DataGrid1.Top + anc + 9
    For jj = 0 To 2
        txtAux1(jj).Text = ""
    Next jj
    
    LLamaLineas anc, 1
    txtAux1(0).Enabled = Anadir
    cmdCta1.Visible = Anadir
    If Anadir Then
        txtAux1(0).Text = "": txtAux1(1).Text = ""
        txtAux1(2).Text = ImporteQueda
    Else
        txtAux1(0).Text = Adodc1.Recordset!Cta
        txtAux1(1).Text = Adodc1.Recordset!nomcta
        txtAux1(2).Text = Adodc1.Recordset!acumperd
        ImporteQueda = ImporteQueda + Adodc1.Recordset!acumperd
    End If
    'Ponemos el foco
    If Anadir Then
        txtAux1(0).SetFocus
    Else
        txtAux1(2).SetFocus
    End If
    
    
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
DeseleccionaGrid

cmdCta1.Top = alto
cmdCta1.Visible = xModo > 0
'Fijamos el ancho
For jj = 0 To 2
    txtAux1(jj).Visible = xModo > 0
    txtAux1(jj).Top = alto
Next jj
End Sub

Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub





Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    txtAux1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux1(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Orden As Integer
    
    If ColumnHeader.Index = 3 Then
        'Ordena por fecha. Para eso hemos metido una columna mas (0 de ancho)
        'con la fecha en formato yyyymmdd
        Orden = 9
    Else
        If ColumnHeader.Index = 6 Then
            'Igual que lo anterior. Solo que el importe va en formato 00000000
            Orden = 10
        Else
            Orden = ColumnHeader.Index
        End If
    End If
    Orden = Orden - 1
    Cad = ColumnHeader.Text
    Cad = "Desea ordenar por el campo """ & Cad & """ "
    If ListView1.SortKey = Orden Then
        If ListView1.SortOrder = lvwAscending Then
            Cad = Cad & "(descendente)?"
        Else
            Cad = Cad & "(ascendente)?"
        End If
    End If
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    If ListView1.SortKey <> Orden Then
        ListView1.SortOrder = lvwAscending
        ListView1.SortKey = Orden
     Else
        If ListView1.SortOrder = lvwAscending Then
            Orden = 1
        Else
            Orden = 0
        End If
        ListView1.SortOrder = Orden
    End If
    ListView1.Sorted = True
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu Me.mnPopUp
    End If
End Sub

Private Sub ListView22_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Orden As Integer
  
    If ColumnHeader.Index = 3 Then
        'Ordena por fecha. Para eso hemos metido una columna mas (0 de ancho)
        'con la fecha en formato yyyymmdd
        Orden = ColumnHeader.Index
    Else
        If ColumnHeader.Index = 6 Then
            'Igual que lo anterior. Solo que el importe va en formato 00000000
            Orden = 8
        Else
            Orden = ColumnHeader.Index
        End If
    End If
   
    Orden = Orden - 1
    Cad = ColumnHeader.Text
    Cad = "Desea ordenar por el campo """ & Cad & """ "
    If ListView22.SortKey = Orden Then
        If ListView22.SortOrder = lvwAscending Then
            Cad = Cad & "(descendente)?"
        Else
            Cad = Cad & "(ascendente)?"
        End If
    End If
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    If ListView22.SortKey <> Orden Then
        ListView22.SortOrder = lvwAscending
        ListView22.SortKey = Orden
     Else
        If ListView22.SortOrder = lvwAscending Then
            Orden = 1
        Else
            Orden = 0
        End If
        ListView22.SortOrder = Orden
    End If
    ListView22.Sorted = True
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub ListView22_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim I As Currency
    Set ListView22.SelectedItem = Item
    I = ImporteFormateado(Item.SubItems(5))
    If Not Item.Checked Then I = -1 * I
    ImporteQueda = ImporteQueda + I
    Label6.Caption = "Devuelto: " & Format(ImporteQueda, FormatoImporte)
End Sub

'Private Sub ListView2_DblClick()
'Dim i As Integer
'Dim cad As String
'
'    For i = 1 To Me.ListView2.ColumnHeaders.Count
'        cad = cad & Me.ListView2.ColumnHeaders(i).Text & " : " & Me.ListView2.ColumnHeaders(i).Width & vbCrLf
'    Next i
'    MsgBox cad
'End Sub

Private Sub mnBanco_Click(Index As Integer)

    If ListView1.SelectedItem Is Nothing Then Exit Sub
    For jj = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(jj).Selected Then
            If ListView1.ListItems(jj).SubItems(7) = "" Then
                ImporteModificando = ImporteModificando + TextoAimporte(ListView1.ListItems(jj).SubItems(5))
                lblImporteRemesa.Caption = "importe: " & Format(ImporteModificando, FormatoImporte)
            End If
            'ListView1.SelectedItem.SubItems(5) = mnBanco(Index).Caption
            'ListView1.SelectedItem.SubItems(6) = mnBanco(Index).Tag
            ListView1.ListItems(jj).SubItems(6) = mnBanco(Index).Caption
            ListView1.ListItems(jj).SubItems(7) = mnBanco(Index).Tag
            ListView1.ListItems(jj).Selected = False
            
        End If
    Next jj
    ListView1.SelectedItem = Nothing
End Sub


Private Sub CambiarSeleccionado(Indice As Integer)
Dim Imp1 As Currency

    Imp1 = ImporteFormateado(ListView1.ListItems(Indice).SubItems(1))
    Imp1 = CCur(Label4.Tag) + Imp1
    
End Sub

Private Sub mnQuitarRecibo_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub


    For jj = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(jj).Selected Then
            ' ListView1.ListItems(jj).SubItems(5) = 0
            If ListView1.ListItems(jj).SubItems(7) <> "" Then
                ImporteModificando = ImporteModificando - TextoAimporte(ListView1.ListItems(jj).SubItems(5))
                lblImporteRemesa.Caption = "importe: " & Format(ImporteModificando, FormatoImporte)
            End If
            ListView1.ListItems(jj).SubItems(7) = ""
            ListView1.ListItems(jj).SubItems(6) = ""
            ListView1.ListItems(jj).Selected = False
           
        End If
    Next jj
    ListView1.SelectedItem = Nothing
End Sub



Private Sub txtAux1_GotFocus(Index As Integer)
    With txtAux1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtAux1_LostFocus(Index As Integer)
Dim D As String
    If Index = 0 Then
        'Cta
        Cad = Trim(txtAux1(0).Text)
        If Cad <> "" Then
            If CuentaCorrectaUltimoNivel(Cad, D) Then
                txtAux1(0).Text = Cad
                Cad = DevuelveDesdeBD("entidad", "ctabancaria", "codmacta", Cad, "T")
                If Cad = "" Then
                    MsgBox "Cuenta no relacionada con ningún banco", vbExclamation
                    txtAux1(1).Text = ""
                    txtAux1(0).SetFocus
                Else
                    
                    txtAux1(1).Text = D
                End If
            Else
                MsgBox D, vbExclamation
                txtAux1(0).Text = ""
                txtAux1(1).Text = ""
                PonFoco txtAux1(0)
            End If
        Else
            txtAux1(1).Text = ""
        End If
    Else
        'INDEX=2  ... Importe
        txtAux1(2).Text = Trim(txtAux1(2).Text)
        If txtAux1(2).Text <> "" Then
            If Not IsNumeric(txtAux1(2).Text) Then
                MsgBox "Campo debe ser numerico.", vbExclamation
                txtAux1(2).Text = ""
            End If
        End If
        
    End If
    
    
    
    
End Sub



Private Sub MostrarVencimientosRemesasNormal()
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim impo As Currency
Dim PonerBanco As Boolean
Dim Imp1 As Currency
Dim AnoRem As Integer
Dim vRem As Integer


    'Primero ponemos los menus del boton derecho
    impo = 0
    jj = 0
    Adodc1.Recordset.MoveFirst
    While Not Adodc1.Recordset.EOF
        jj = jj + 1
        mnBanco(jj).Caption = Adodc1.Recordset!nomcta
        mnBanco(jj).Tag = Adodc1.Recordset!Cta
        Adodc1.Recordset.MoveNext
    Wend
    
    PonerBanco = True
    If jj < 7 Then
        Cad = jj + 1
        For jj = Val(Cad) To 7
            mnBanco(jj).Visible = False
        Next jj
    End If
    
    If Opcion = 1 Then
        vRem = Val(RecuperaValor(vRemesa, 1))
        AnoRem = Val(RecuperaValor(vRemesa, 2))
    End If
    'Cargamos los datos
    ListView1.ListItems.Clear
    Adodc1.Recordset.MoveFirst
    Cad = "Select scobro.*,nommacta  " & vSQL
    'IMPORTES >0
    Cad = Cad & " AND impvenci >0  "
    Cad = Cad & " ORDER BY fecvenci,codmacta"
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ImporteQueda = 0
    impo = 0
    
    
    
    
    While Not RS.EOF
    
        Imp1 = RS!ImpVenci - DBLet(RS!impcobro, "N") + DBLet(RS!Gastos, "N")
        If Not IsNull(RS!CodRem) Then
            If RS!CodRem = vRem And RS!AnyoRem = AnoRem Then
                'Si es esta remesa, y la situacion
                If RS!siturem = "B" Then Imp1 = RS!ImpVenci + DBLet(RS!Gastos, "N")
            End If
        End If
        'Los que sean CERO no los pinto
        If Imp1 > 0 Then
                Set ItmX = ListView1.ListItems.Add
                ItmX.Text = RS!NUmSerie
                ItmX.SubItems(1) = Format(RS!codfaccl, "00000000")
                'Enero 2011
                ItmX.SubItems(2) = Format(RS!FecVenci, "dd/mm/yyyy")
                ItmX.SubItems(3) = RS!numorden
                ItmX.SubItems(4) = RS!Nommacta
                
                
                ItmX.SubItems(5) = Format(Imp1, FormatoImporte)
                
                'Por si ordena por importe
                ItmX.SubItems(9) = Format(Imp1 * 100, "0000000000")
                
                'Fecvenci
                'Por si ordena por fecha
                ItmX.SubItems(8) = Format(RS!FecVenci, "yyyymmdd")
                
                'Fecha factura
                ItmX.SubItems(10) = Format(RS!fecfaccl, "dd/mm/yyyy")
                'ANTES
                
                'ItmX.SubItems(4) = Format(Rs!impvenci - DBLet(Rs!impcobro, "N") + DBLet(Rs!Gastos, "N"), FormatoImporte)
                'ImporteQueda = ImporteQueda + (Rs!impvenci - DBLet(Rs!impcobro, "N") + DBLet(Rs!Gastos, "N"))
                
                'Si ya ha sido devuelto
                'If Rs!contdocu = 1 Then
                If RS!Devuelto = 1 Then
                
                
                    ItmX.Bold = True
                    ItmX.ForeColor = vbRed
                End If
                
                
                
                
                'Ponog el banoc
                'Si es modificar remesa , o no
                If Opcion = 1 Then
                        ItmX.SubItems(6) = ""
                        ItmX.SubItems(7) = ""
        
                        If Not IsNull(RS!CodRem) Then
                            If RS!CodRem = vRem And RS!AnyoRem = AnoRem Then
                                ItmX.SubItems(6) = Adodc1.Recordset!nomcta
                                ItmX.SubItems(7) = Adodc1.Recordset!Cta
                                impo = impo + Imp1
                            End If
          
                        End If
                        
                Else
                    If PonerBanco Then
                        ImporteQueda = ImporteQueda + Imp1
                        ItmX.SubItems(6) = Adodc1.Recordset!nomcta
                        ItmX.SubItems(7) = Adodc1.Recordset!Cta
                    End If
                End If
                
                If PonerBanco Then
                    If Opcion = 0 Then
                        If ImporteQueda > Adodc1.Recordset!acumperd Then
                            Adodc1.Recordset.MoveNext
                            If Adodc1.Recordset.EOF Then
                                PonerBanco = False
                            Else
                                ImporteQueda = 0
                                
                            End If
                        End If
                    End If
                End If
        
        End If   'DE imp >0  Si es cero no lo pongo
        RS.MoveNext
    Wend
    RS.Close
    
    If Opcion = 1 Then
        'MODIFICAR REIBOS REMESA
        ImporteModificando = impo
    Else
        'NUEVA REMESA
        ImporteModificando = ImporteQueda
    End If
    lblImporteRemesa.Caption = "importe: " & Format(ImporteModificando, FormatoImporte)
    lblImporteRemesa.Visible = True
    If mnBanco(2).Visible Then lblImporteRemesa.Visible = False
    
    Set RS = Nothing
End Sub



Private Sub MostrarVencimientosRemesasPorEntidad()
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
'Dim AnoRem As Integer
'Dim vRem As Integer
Dim Impor As Currency
Dim Eliminados As String

    'Primero ponemos los menus del boton derecho
    
    jj = 0
    Set RS = New ADODB.Recordset
    ValoresDevolucionRemesa = RecuperaValor(vRemesa, 4)
    
    'para cargar cada entidad
        
    For jj = 1 To 7
        mnBanco(jj).Caption = ""
        mnBanco(jj).Tag = ""
        mnBanco(jj).Visible = False
    Next jj
    
    
    Adodc1.Recordset.MoveFirst
    Cad = ""
    jj = 0
    While Not Adodc1.Recordset.EOF
        jj = jj + 1
        mnBanco(jj).Caption = Adodc1.Recordset!nomcta
        mnBanco(jj).Tag = Adodc1.Recordset!Cta
        mnBanco(jj).Visible = True
        If Adodc1.Recordset!Cta = ValoresDevolucionRemesa Then NumRegElim = jj
        

        Adodc1.Recordset.MoveNext
    Wend
    
    Cad = "Select cta from tmp347 where codusu =" & vUsu.Codigo
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not RS.EOF
        Cad = Cad & ", '" & RS!Cta & "'"
        RS.MoveNext
    Wend
    RS.Close
    
    'NO PUEDE SER ""
    Cad = Mid(Cad, 2)
    Cad = "(" & Cad & ")"
    'ValoresDevolucionRemesa:  llevara entidad(0000)oficina(0000)·indice| indice=jj del menu
    Cad = "Select * from ctabancaria where codmacta in " & Cad
    Cad = Cad & " AND codmacta<>'" & ValoresDevolucionRemesa & "'"
    
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ValoresDevolucionRemesa = ""
    Eliminados = ""   'bancos eliminados
    While Not RS.EOF
        'Para cada banco entidad
        
        For jj = 1 To 7
           If mnBanco(jj).Visible Then
                If mnBanco(jj).Tag = RS!codmacta Then
                    'Cad = "·" & Format(RS!Entidad, "0000") & Format(RS!oficina, "0000") & "#" & jj
                    Cad = "·" & Format(RS!Entidad, "0000") & "#" & jj
                    ValoresDevolucionRemesa = ValoresDevolucionRemesa & Cad
                    Exit For
                End If
            End If
        Next
        If jj > 7 Then
            'Cad = "·" & Format(RS!Entidad, "0000") & Format(RS!oficina, "0000") & "#" & jj
            Cad = "·" & Format(RS!Entidad, "0000") & "#" & jj
            Eliminados = Eliminados & Cad
        End If
        RS.MoveNext
    Wend
    RS.Close
    'El banco por defecto lo tendremos en NUMREGELIM
    'Entonces buscaremos en ValoresDevolucionRemesa donde estara para cada banco
    ' que no sea el de pordefecto entidad oficina indice (en el mnBanco menu
    ' sera ·eeeeoooo#i    e:ntidad    o:ficina  i:ndice
        
    
    
   
    'Cargamos los datos
    ListView1.ListItems.Clear
    Adodc1.Recordset.MoveFirst
    Cad = "Select scobro.*,nommacta  " & vSQL
    'IMPORTES >0
    Cad = Cad & " AND impvenci >0  "
    Cad = Cad & " ORDER BY fecvenci,codmacta"
    
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

   
    ImporteQueda = 0
    
    
    While Not RS.EOF
    
        Impor = RS!ImpVenci - DBLet(RS!impcobro, "N") + DBLet(RS!Gastos, "N")
        'Cad = "·" & Format(RS!codbanco, "0000") & Format(RS!codsucur, "0000") & "#"
        Cad = "·" & Format(RS!codbanco, "0000") & "#"
        
        
        'Veremos que banco le toca
        jj = InStr(1, Eliminados, Cad)
        If jj > 0 Then
            'De los bancoas que no vamops a girar ahora
            Impor = 0
        Else
            jj = InStr(1, ValoresDevolucionRemesa, Cad)
            If jj > 0 Then
                Cad = Mid(ValoresDevolucionRemesa, jj + 6, 1)
                jj = CInt(Cad)
                If jj = 0 Then jj = NumRegElim
            Else
                jj = NumRegElim
            End If
        End If
        
        
        'Los que sean CERO no los pinto
        If Impor > 0 Then
                ImporteQueda = ImporteQueda + Impor
        
                Set ItmX = ListView1.ListItems.Add
                ItmX.Text = RS!NUmSerie
                ItmX.SubItems(1) = Format(RS!codfaccl, "00000000")
                'Enero 2011
                ItmX.SubItems(2) = Format(RS!FecVenci, "dd/mm/yyyy")
                ItmX.SubItems(3) = RS!numorden
                ItmX.SubItems(4) = RS!Nommacta
                
                
                ItmX.SubItems(5) = Format(Impor, FormatoImporte)
                
                'Por si ordena por importe
                ItmX.SubItems(9) = Format(Impor * 100, "0000000000")
                
                'Fecvenci
                'Por si ordena por fecha
                ItmX.SubItems(8) = Format(RS!FecVenci, "yyyymmdd")
                
                'Fecha factura
                ItmX.SubItems(10) = Format(RS!fecfaccl, "dd/mm/yyyy")
                'ANTES
                
                'ItmX.SubItems(4) = Format(Rs!impvenci - DBLet(Rs!impcobro, "N") + DBLet(Rs!Gastos, "N"), FormatoImporte)
                'ImporteQueda = ImporteQueda + (Rs!impvenci - DBLet(Rs!impcobro, "N") + DBLet(Rs!Gastos, "N"))
                
                'Si ya ha sido devuelto
                'If Rs!contdocu = 1 Then
                If RS!Devuelto = 1 Then
                
                
                    ItmX.Bold = True
                    ItmX.ForeColor = vbRed
                End If
               
                
                 ItmX.SubItems(6) = mnBanco(jj).Caption
                 ItmX.SubItems(7) = mnBanco(jj).Tag
              
                 
            
        
        End If   'DE imp >0  Si es cero no lo pongo
        RS.MoveNext
    Wend
    RS.Close
    

    ImporteModificando = ImporteQueda

    lblImporteRemesa.Caption = "importe: " & Format(ImporteModificando, FormatoImporte)
    lblImporteRemesa.Visible = True
    If mnBanco(2).Visible Then lblImporteRemesa.Visible = False
    
    
    
    ValoresDevolucionRemesa = ""
    Set RS = Nothing
End Sub



Private Sub RejusteImportesBanco()
Dim CadB As String
Dim I As Integer
Dim K As Integer
    'Lo k haremos sera reclaular lo k va a cada campo.. de verdad
    
    
    CadB = ""
    jj = 0
    Adodc1.Recordset.MoveFirst
    'Vemos las cuentas contables
    While Not Adodc1.Recordset.EOF
        jj = jj + 1
        CadB = CadB & Adodc1.Recordset!Cta & "|"
        Adodc1.Recordset.MoveNext
    Wend
    
    'Cerramos ya el adodc1
    Adodc1.Recordset.Close
    Set DataGrid1.DataSource = Nothing
    
    
    'UPDATEAMOS CON LOS VALORES REALES

        For I = 1 To jj
            Cad = RecuperaValor(CadB, I)
            ImporteQueda = 0
            'El banco en cad
            For K = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(K).SubItems(7) = Cad Then
                    'ES IMPORTE DEL BANCO
                    ImporteQueda = ImporteQueda + ImporteFormateado(ListView1.ListItems(K).SubItems(5))
                End If
            Next K
            'Ya tenemos el importe real para el banco
            
            Cad = "UPDATE tmpcierre1 SET acumperd = " & TransformaComasPuntos(CStr(ImporteQueda)) & " WHERE cta= '" & Cad & "'"
            Conn.Execute Cad
        
        Next I
        
        

    
    'Por si acaso algun banco esta a 0
    Cad = "DELETE FROM tmpcierre1 where codusu =" & vUsu.Codigo & " AND acumperd =0"
    Conn.Execute Cad
    
    
    Set Me.DataGrid2.DataSource = Adodc1
    
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = "Select cta,nomcta,acumperd from tmpcierre1 where codusu =" & vUsu.Codigo
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockOptimistic
    Adodc1.Refresh

    
    
    'Cuenta contable
    I = 0
        DataGrid2.Columns(I).Caption = "Cuenta"
        DataGrid2.Columns(I).Width = 1100
    
    'Descripcion NOMMACTA
    I = 1
        DataGrid2.Columns(I).Caption = "Descripción"
        DataGrid2.Columns(I).Width = 3200
 
    
    'Entidad
    I = 2
        DataGrid2.Columns(I).Caption = "Importe"
        DataGrid2.Columns(I).Width = 1200
        DataGrid2.Columns(I).Alignment = dbgRight
        DataGrid2.Columns(I).NumberFormat = FormatoImporte
    
End Sub


Private Sub GenerarRemesa()
Dim C As String
Dim NumeroRemesa As Long
Dim RS As ADODB.Recordset

    CadenaDesdeOtroForm = "MAL"
    NumeroRemesa = Val(Text1(0).Text)
    Set RS = New ADODB.Recordset
    Cad = "Select * from tmpcierre1 where codusu =" & vUsu.Codigo
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        MsgBox "Error grave. Datos temporales vacios", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Sub
    End If
    
    
    Set miRsAux = New ADODB.Recordset
    
    'Para ver si existe la remesa... pero esto no tendria k pasar
    '------------------------------------------------------------
    While Not RS.EOF
    
        
        
        Cad = "Select * from remesas where codigo=" & NumeroRemesa
        Cad = Cad & " AND anyo =" & Text1(1).Text
        Cad = Cad & " AND tiporem = 1"
    
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then Cad = ""
        miRsAux.Close
        
        
        If Opcion = 0 Then
            If Cad <> "" Then
                MsgBox "Ya existe el codigo de remesa.", vbExclamation
                Exit Sub
            End If
        Else
            If Cad = "" Then
                MsgBox "Error leyendo la remesa", vbExclamation
                Exit Sub
            End If
        End If
        
        If Opcion = 0 Then
            'Ahora insertamos la remesa
            'cad = "INSERT INTO remesas (codigo, anyo, fecremesa, fecini, fecfin, situacion) VALUES ( "
            Cad = "INSERT INTO remesas (codigo, anyo, fecremesa,situacion,codmacta,descripcion,tiporem   ) "
            Cad = Cad & " VALUES ( "
            Cad = Cad & NumeroRemesa & "," & Text1(1).Text & ",'" & Format(Text1(2).Text, FormatoFecha) & "','A','"
            Cad = Cad & RS!Cta & "',' " & DevNombreSQL(Text2.Text) & "',1)"
            Conn.Execute Cad
            
            
        Else
            'Paso la remesa a estado: A
            'Vuelvo a poner los vecnimientos a NULL para poder
            'meterlos luego
            
            '---remesa estado A
            Cad = "UPDATE Remesas SET Situacion = 'A'"
            Cad = Cad & ", descripcion ='" & DevNombreSQL(Text2.Text) & "'"
            Cad = Cad & " WHERE codigo=" & Text1(0).Text
            Cad = Cad & " AND anyo =" & Text1(1).Text
            If Not Ejecuta(Cad) Then Exit Sub
            
            Cad = "UPDATE scobro SET siturem=NULL, codrem=NULL, anyorem=NULL ,tiporem =NULL "
            Cad = Cad & " ,fecultco=NULL,ctabanc2=null, impcobro = NULL "
            Cad = Cad & " WHERE codrem = " & Text1(0).Text
            Cad = Cad & " AND anyorem=" & Text1(1).Text & " AND tiporem = 1"
            If Not Ejecuta(Cad) Then Exit Sub
        End If
        
        
        
        'Ahora cambiamos los cobros y les ponemos la remesa
        Cad = "UPDATE  scobro SET siturem= 'A',codrem= " & NumeroRemesa & ", anyorem =" & Text1(1).Text & ","
        Cad = Cad & " tiporem = 1"
        'Lo cobrado tb a NULL
        Cad = Cad & ", impcobro = NULL"
        Cad = Cad & ", fecultco = NULL"
        'ponemos la cuenta de banco donde va remesado
        Cad = Cad & ", ctabanc2 ='"
        
        
        
        
        'Para cada cobro UPDATE
        For jj = 1 To ListView1.ListItems.Count
           ' WHERE numserie='A' AND codfaccl=99588 AND fecfaccl='2003-12-01' AND numorden=4"
           With ListView1.ListItems(jj)
                'Si el subitem es del banco
                If .SubItems(7) = RS!Cta Then
           
           
                    'Cuenta de banco
                    C = .SubItems(7) & "' "
                    C = C & "WHERE numserie = '" & .Text & "' and codfaccl = "
                    C = C & Val(.SubItems(1)) & " and fecfaccl ='" & Format(.SubItems(10), FormatoFecha)
                    C = C & "' AND numorden =" & .SubItems(3)
                    
                    C = Cad & C
                    Conn.Execute C
                    
                End If
           End With
           
        Next jj
        espera 0.5
        
        
        'Hacemos un select sum para el importe
        Cad = "Select sum(impvenci),sum(impcobro),sum(gastos) from scobro "
        Cad = Cad & " WHERE codrem=" & NumeroRemesa
        Cad = Cad & " AND anyorem =" & Text1(1).Text
        Cad = Cad & " AND tiporem = 1"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        C = "0"
        If Not miRsAux.EOF Then
            If Not IsNull(miRsAux.Fields(0)) Then
                               'Impvenci                               impcobro                      gastos
                ImporteQueda = DBLet(miRsAux.Fields(0), "N") - DBLet(miRsAux.Fields(1), "N") + DBLet(miRsAux.Fields(2), "N")
                C = TransformaComasPuntos(CStr(ImporteQueda))
            End If
        End If
        miRsAux.Close
        
        Cad = "UPDATE Remesas SET importe=" & C
        Cad = Cad & " WHERE codigo=" & NumeroRemesa
        Cad = Cad & " AND anyo =" & Text1(1).Text
        Cad = Cad & " AND tiporem = 1"
        Conn.Execute Cad
        
        
        'SIGUIENTE REMESA
        RS.MoveNext
        NumeroRemesa = NumeroRemesa + 1
    Wend
    RS.Close
    Set miRsAux = Nothing
    Set RS = Nothing
    CadenaDesdeOtroForm = ""
End Sub

'Carga los efectos de la remesa indicada
' y marcara los k hayan sido devueltos
Private Sub CargaDevolucion()
Dim Itm As ListItem
Dim Col As Collection
Dim EfectoSerie As String
Dim EfectoFra As String
Dim EfectoVto As String
Dim EltoItm  As ListItem
Dim EsSepa As Boolean


    'Si viene de la opcion de devolucion por efecto, este campo tiene valor
    EfectoSerie = ""
    Set EltoItm = Nothing
    If vSQL <> "" Then
        EfectoSerie = RecuperaValor(vSQL, 1)
        EfectoFra = Format(Val(RecuperaValor(vSQL, 2)), "00000000")
        EfectoVto = RecuperaValor(vSQL, 3)
    End If


    'Veremos si viene de un fichero de devolicion, y si trae mas de una remesa
    vSQL = ""
    Cad = RecuperaValor(vRemesa, 8)
    If Cad <> "" Then
        'Fichero de dovocucion
        Cad = RecuperaValor(vRemesa, 9)
        'Vuelvo a susitiuri los # por |
        Cad = Replace(Cad, "#", "|")
        vSQL = ""
        For jj = 1 To Len(Cad)
            If Mid(Cad, jj, 1) = "·" Then vSQL = vSQL & "X"
        Next
        
        If Len(vSQL) > 1 Then
            'Tienen mas de una remesa
            vSQL = ""
            While Cad <> ""
                jj = InStr(1, Cad, "·")
                If jj = 0 Then
                    Cad = ""
                Else
                    vSQL = vSQL & ", (" & RecuperaValor(Mid(Cad, 1, jj - 1), 1) & " , " & RecuperaValor(Mid(Cad, 1, jj - 1), 2) & ")"
                    Cad = Mid(Cad, jj + 1)
                End If
            
            Wend
            vSQL = Mid(vSQL, 2) 'quitammos la preimar coma
        Else
            vSQL = ""
        End If
        
    End If


    If vSQL = "" Then
        'Normal
        vSQL = " AND codrem =" & RecuperaValor(vRemesa, 1)
        vSQL = vSQL & " AND anyorem =" & RecuperaValor(vRemesa, 2)
    
    Else
        'Multi remesa
        vSQL = " AND (codrem,anyorem) IN ( " & vSQL & ")"
        
    End If
    vSQL = "Select scobro.*,nommacta from scobro,cuentas where scobro.codmacta = cuentas.codmacta" & vSQL
    
    vSQL = vSQL & " ORDER BY numserie,codfaccl"
    Set miRsAux = New ADODB.Recordset
    ListView22.ListItems.Clear
    miRsAux.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    jj = 1
    While Not miRsAux.EOF
        Set Itm = ListView22.ListItems.Add(, "C" & jj)
        Itm.Text = miRsAux!NUmSerie
        
        Itm.SubItems(1) = Format(miRsAux!codfaccl, "00000000")
        Itm.SubItems(2) = miRsAux!numorden
        Itm.SubItems(3) = miRsAux!codmacta
        Itm.SubItems(4) = miRsAux!Nommacta
        ImporteQueda = DBLet(miRsAux!Gastos, "N")
        'No lo pongo con el importe de gastos pq pudiera ser k habiendo sido devuelto, no quiera
        ' cobrarle gastos
        If miRsAux!Devuelto = 1 Then
            Itm.Bold = True
            Itm.ForeColor = vbRed
        End If
        ImporteQueda = ImporteQueda + miRsAux!ImpVenci
        Itm.SubItems(5) = Format(ImporteQueda, FormatoImporte)
        
        'Para la ordenacion
        'Por si ordena por fecha
        'ItmX.SubItems(6) = Format(RS!fecfaccl, "yyyymmdd")
        'Por si ordena por importe
        Itm.SubItems(7) = Format(miRsAux!ImpVenci * 100, "0000000000")
        
        
        
        'En el tag meto la fecha factura
        Itm.Tag = Format(miRsAux!fecfaccl, "dd/mm/yyyy")
        
        

        If EfectoSerie <> "" Then
            If EfectoSerie = Itm.Text Then
                If EfectoFra = Itm.SubItems(1) Then
                    If EfectoVto = Itm.SubItems(2) Then
                        Set EltoItm = Itm
                        'Este es. Para que ya no busque mas
                        EfectoSerie = ""
                    End If
                End If
            End If
        End If
        
        
        jj = jj + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    ImporteQueda = 0
    Cad = RecuperaValor(vRemesa, 8)
    If Cad <> "" Then
        'DEVOLUCION CON FICHERO
        
        Me.Tag = Label2(4).Caption
        Label2(4).Caption = "Leyendo fichero de datos......"
        Me.Refresh
        Screen.MousePointer = vbHourglass
        Set Col = New Collection
        
        Dim TipoFicheroSepa As Byte
        
        TipoFicheroSepa = EsFicheroDevolucionSEPA2(Cad)
        
        If TipoFicheroSepa = 2 Then
            'SEPA XML
            ProcesaLineasFicheroDevolucionXML Cad, Col
            EsSepa = True
        Else
            ProcesaLineasFicheroDevolucion Cad, Col, EsSepa
        End If
        
        Me.Refresh
        If Not (Col Is Nothing) Then
            'Si Col no es nothing
            If Col.Count > 0 Then
                '-Aqui iremos recorriendo el COl hasta encontrar slos recibos que
                'Son a devolver.
                RecorreBuscandoRecibo Col, False, EsSepa
                If Col.Count > 0 Then RecorreBuscandoRecibo Col, True, EsSepa
            End If
            Label2(4).Caption = Me.Tag
        End If
        Me.Tag = ""
        
        
        
        'Borraremos los que no esten en el fichero
        ImporteQueda = 0
        For jj = Me.ListView22.ListItems.Count To 1 Step -1
            If Not Me.ListView22.ListItems(jj).Checked Then
                Me.ListView22.ListItems.Remove jj
            Else
                ImporteQueda = ImporteQueda + ImporteFormateado(ListView22.ListItems(jj).SubItems(5))
            End If
        Next
    Else
        If Not (EltoItm Is Nothing) Then
            'Ha encontrado un vto
            EltoItm.Checked = True
            ListView22_ItemCheck EltoItm
            
            Set EltoItm = Nothing
        End If
    End If
    
    
    
    
    
    
    Me.Refresh
    Screen.MousePointer = vbDefault
        
    
End Sub






Private Function RealizarDevolucion() As Boolean
Dim IncPorcentaje As Boolean
Dim Gasto As Currency
Dim Minimo As Currency
    
    RealizarDevolucion = False
    'Tipo de aumento del gasto de devolucion
    Cad = RecuperaValor(Me.vRemesa, 5)
    If Cad = "%" Then
        'Porcentual
        IncPorcentaje = True
        Minimo = 0
        Cad = RecuperaValor(vRemesa, 6)
        If Cad <> "" Then Minimo = Cad
    Else
        IncPorcentaje = False
    End If
        
    
    vSQL = "DELETE FROM tmpfaclin WHERE codusu =" & vUsu.Codigo
    Conn.Execute vSQL
    '                                               numero        serie vto
    vSQL = "INSERT INTO tmpfaclin (codusu, codigo, Numfac, Fecha, IVA, NIF,  "
    vSQL = vSQL & "Imponible,  ImpIVA,total,cta,cliente) VALUES (" & vUsu.Codigo & ","
    For jj = 1 To ListView22.ListItems.Count
        If ListView22.ListItems(jj).Checked Then
                                        'cdofaccl
            Cad = jj & "," & Val(ListView22.ListItems(jj).SubItems(1)) & ",'"
                                    'fecfaccl                                                   SERIE
            Cad = Cad & Format(ListView22.ListItems(jj).Tag, FormatoFecha) & "','" & ListView22.ListItems(jj).Text
                                    'numvencimiento numorden
            Cad = Cad & "'," & Val(ListView22.ListItems(jj).SubItems(2)) & ","
            ImporteQueda = ImporteFormateado(ListView22.ListItems(jj).SubItems(5))
            Cad = Cad & TransformaComasPuntos(CStr(ImporteQueda)) & ","
            
            'Calculo el gasto
            If IncPorcentaje Then
                'Importe = importe  * (importe * % )/100
                Gasto = Round((ImporteQueda * ImporteRemesa) / 100, 2)
                
                If Minimo > 0 Then If Gasto < Minimo Then Gasto = Minimo
                
                ImporteQueda = ImporteQueda + Gasto
            Else
                'Importe =importe + incremento
                Gasto = ImporteRemesa
                ImporteQueda = ImporteQueda + ImporteRemesa
            End If
            Cad = Cad & TransformaComasPuntos(CStr(Gasto)) & ","
            Cad = Cad & TransformaComasPuntos(CStr(ImporteQueda)) & ",'"
            'Cuenta cliente, y banco
            Cad = Cad & ListView22.ListItems(jj).SubItems(3) & "','"
            Cad = Cad & RecuperaValor(vRemesa, 3) & "')"
            Cad = vSQL & Cad
            If Not Ejecuta(Cad) Then Exit Function
        End If
    Next jj
    
    
    'OK. Ya tengo grabada la temporal con los recibos que devuelvo. Ahora
    'hare:
    '       - generar un asiento con los datos k devuelvo
    '       - marcar los cobros como devueltos, añadirle el gasto y insertar en la
    '           tabla de hco de devueltos
    
    jj = Val(RecuperaValor(vRemesa, 7))
    
    If jj = 0 Then
        'Como no se contabilizan los beneficios no hace falta que calcule nada
        Cad = ""
     Else
        'Vya obteneer la cuenta de gastos bancarios
        Cad = RecuperaValor(vRemesa, 3)  'cta contable
        Cad = DevuelveDesdeBD("ctagastos", "ctabancaria", "codmacta", Cad, "T")
        If Cad = "" Then
            'NO esta configurada
            'Veo si esta en parametros
            'ctabenbanc
            Cad = DevuelveDesdeBD("ctabenbanc", "paramtesor", "codigo", "1", "N")
        End If
        If Cad = "" Then
            MsgBox "No esta configurada la gastos  bancarios", vbExclamation
            Exit Function
        End If
    End If
    
    If RealizarDevolucionRemesa(CDate(RecuperaValor(vRemesa, 4)), jj > 0, Cad, vRemesa, ValoresDevolucionRemesa) Then
        RealizarDevolucion = True
        Screen.MousePointer = vbHourglass
        frmActualizar2.OpcionActualizar = 20
        frmActualizar2.Show vbModal
        Screen.MousePointer = vbDefault
    End If
End Function


Private Sub RecorreBuscandoRecibo(ByRef Recibos As Collection, EsMensajeNoEncontrados As Boolean, EsSepa As Boolean)
    If EsSepa Then
        RecorreBuscandoReciboSEPA Recibos, EsMensajeNoEncontrados
    Else
        RecorreBuscandoRecibo2 Recibos, EsMensajeNoEncontrados
    End If
End Sub



Private Sub RecorreBuscandoRecibo2(ByRef Recibos As Collection, EsMensajeNoEncontrados As Boolean)
Dim B As Boolean

Dim EsFormatoAntiguoDevolucion As Boolean

    'Formato antiguo:A020500021
    'En el nuevo : X 00045771 >> serie(2)=X  factura(7)=4577    vto(1)=1
    EsFormatoAntiguoDevolucion = Dir(App.Path & "\DevRecAnt.dat") <> ""
    

    If EsMensajeNoEncontrados Then
            Cad = ""
            For jj = Recibos.Count To 1 Step -1
                'Ejemplo 0047080000004708
                '        251205A020500021
                '        $$$$$$ fecha                       6
                '              $ Serie                      1
                '               $$$$$$$$  Facutra           8
                '                       $  Vencimiento      1
                'La fecha
                ImporteQueda = CCur(Val(Mid(Recibos(jj), 1, 10)) / 100)
                Cad = Cad & jj & ".-Fecha: "
                Cad = Cad & Mid(Recibos(jj), 11, 2) & "/" & Mid(Recibos(jj), 13, 2) & "/20" & Mid(Recibos(jj), 15, 2)
                Cad = Cad & " Vto: " & Mid(Recibos(jj), 17, 1) & "/" & Mid(Recibos(jj), 18, 8) & "-" & Mid(Recibos(jj), 26, 1) & "   Importe: " & Format(ImporteQueda, FormatoImporte) & vbCrLf
            Next jj
            Cad = "Recibos no encontrados que vienen del fichero." & vbCrLf & vbCrLf & Cad
            MsgBox Cad, vbExclamation
            ImporteQueda = 0
    Else
        
        For jj = Recibos.Count To 1 Step -1
            'Ejemplo          0047080000004708
            '       0000001234251205A020500021
            '          ...$$$$    Importe                        10
            '                 $$$$$$ fecha                       6
            '                       $ Serie                      1
            '                        $$$$$$$$  Facutra           8
            '                                $  Vencimiento      1
            'La fecha
            Cad = Mid(Recibos(jj), 11, 2) & "/" & Mid(Recibos(jj), 13, 2) & "/20" & Mid(Recibos(jj), 15, 2)
            'Octubre 2011
            'If Not IsNumeric(Mid(Recibos(jj), 27, 1)) Then
               
            'SEPT 2013
            If Not EsFormatoAntiguoDevolucion Then
                'Alzira. Estaba mal formateado el numfac.
               B = EstaElRecibo(Mid(Recibos(jj), 17, 2), Mid(Recibos(jj), 19, 7), Cad, Mid(Recibos(jj), 26, 1))
            Else
               B = EstaElRecibo(Mid(Recibos(jj), 17, 2), Mid(Recibos(jj), 20, 7), Cad, Mid(Recibos(jj), 27, 1))
            End If
            If B Then Recibos.Remove jj
        Next jj
                
    End If
    
End Sub






Private Function EstaElRecibo(Serie As String, Fac As String, fec As String, Venci As String) As Boolean
Dim J As Integer

        'Itm.Text = miRsAux!NUmSerie
        'Itm.SubItems(1) = Format(miRsAux!codfaccl, "0000000000")
        'Itm.SubItems(2) = miRsAux!numorden
        'Itm.Tag = miRsAux!fecfaccl
        

    EstaElRecibo = False
    With ListView22
        For J = 1 To .ListItems.Count
            If Mid(.ListItems(J).Text, 1, 2) = Trim(Serie) Then
                'Misma serie
                If Val(.ListItems(J).SubItems(1)) = Val(Fac) And .ListItems(J).SubItems(2) = Venci And .ListItems(J).Tag = fec Then
                        'Este es el recibo
                        .ListItems(J).Checked = True
                        ImporteQueda = ImporteQueda + ImporteFormateado(.ListItems(J).SubItems(5))
                        EstaElRecibo = True
                        Exit For
                End If
            End If
        Next J
    
    
        'Nov 2012
        If Not EstaElRecibo Then
            'Pruebo solo con el numero de vto y que la primera letra d serie sea como la del vto (pueden ser dos)
            'Ademas meto el numero de vto dentro del fac
            For J = 1 To .ListItems.Count
                If Mid(.ListItems(J).Text, 1, 1) = Mid(Trim(Serie), 1, 1) Then
                        'Misma serie
                        If Val(.ListItems(J).SubItems(1)) = Val(Fac & Venci) And .ListItems(J).Tag = fec Then
                                'Este es el recibo
                                .ListItems(J).Checked = True
                                ImporteQueda = ImporteQueda + ImporteFormateado(.ListItems(J).SubItems(5))
                                EstaElRecibo = True
                                Exit For
                        End If
                End If
            Next
        End If
    End With
End Function


'Esta recibo SEPA
Private Sub RecorreBuscandoReciboSEPA(ByRef Recibos As Collection, EsMensajeNoEncontrados As Boolean)
Dim B As Boolean



    

    If EsMensajeNoEncontrados Then
            Cad = ""
            For jj = Recibos.Count To 1 Step -1
                'M  0330047820131201001   430000061
                'SER FACTU    FEC   VTO
                
                'ImporteQueda = CCur(Val(Mid(Recibos(jj), 1, 10)) / 100)
                Cad = Cad & jj & ".-Fecha: "
                Cad = Cad & Mid(Recibos(jj), 18, 2) & "/" & Mid(Recibos(jj), 16, 2) & "/" & Mid(Recibos(jj), 12, 4)
                Cad = Cad & " Vto: " & Mid(Recibos(jj), 1, 3) & "/" & Mid(Recibos(jj), 4, 8) & "-" & Mid(Recibos(jj), 20, 3) & vbCrLf
            Next jj
            Cad = "Recibos no encontrados que vienen del fichero." & vbCrLf & vbCrLf & Cad
            MsgBox Cad, vbExclamation
            ImporteQueda = 0
    Else
        
        For jj = Recibos.Count To 1 Step -1
            'M  0330047820131201001   430000061
            'SER FACTU    FEC   VTO
            Cad = Mid(Recibos(jj), 18, 2) & "/" & Mid(Recibos(jj), 16, 2) & "/" & Mid(Recibos(jj), 12, 4)
            
            
            B = EstaElReciboSEPA(Trim(Mid(Recibos(jj), 1, 3)), Mid(Recibos(jj), 4, 8), Cad, Mid(Recibos(jj), 20, 3))

            If B Then Recibos.Remove jj
        Next jj
                
    End If
    
End Sub



Private Function EstaElReciboSEPA(Serie As String, Fac As String, fec As String, Venci As String) As Boolean
Dim J As Integer

        'Itm.Text = miRsAux!NUmSerie
        'Itm.SubItems(1) = Format(miRsAux!codfaccl, "0000000000")
        'Itm.SubItems(2) = miRsAux!numorden
        'Itm.Tag = miRsAux!fecfaccl
        

    EstaElReciboSEPA = False
    With ListView22
        For J = 1 To .ListItems.Count
            If Trim(.ListItems(J).Text) = Trim(Serie) Then
                'Misma serie
                If Val(.ListItems(J).SubItems(1)) = Val(Fac) And Val(.ListItems(J).SubItems(2)) = Venci And .ListItems(J).Tag = fec Then
                        'Este es el recibo
                        .ListItems(J).Checked = True
                        ImporteQueda = ImporteQueda + ImporteFormateado(.ListItems(J).SubItems(5))
                        EstaElReciboSEPA = True
                        Exit For
                End If
            End If
        Next J
    
    End With
End Function




Private Function ComprobarRiesgoBanco() As Boolean
Dim C As Collection

    On Error GoTo EComprobarRiesgoBanco
    ComprobarRiesgoBanco = False
    'Insertaremos strings empipados con codmacta|nommacta|riesgo|importearemesarahora|importeYaremesadoEnScobro|
    Set C = New Collection
    
    Set miRsAux = New ADODB.Recordset
    'En tmpcierre tengo para cada banco, cuanto dinero quiero remesar
    'Co Lo cual recorreremos y comprobaremos con maximo riesgo remesa
    Cad = "Select * from tmpcierre1 where codusu = " & vUsu.Codigo & " ORDER BY cta"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
            
        ValoresDevolucionRemesa = DevuelveDesdeBD("remesamaximo", "ctabancaria", "codmacta", miRsAux!Cta, "T")
        
        'Si no he puesto el riesgo NO compruebo nada
        If ValoresDevolucionRemesa <> "" Then
            Cad = miRsAux!Cta & "|" & miRsAux!nomcta & "|" & ValoresDevolucionRemesa & "|"
            'A remesar ahora
            Cad = Cad & CStr(miRsAux!acumperd) & "|"
            C.Add Cad
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If C.Count > 0 Then
        If CompruebaImportesRemesa(C) Then ComprobarRiesgoBanco = True
    Else
        ComprobarRiesgoBanco = True
    End If

    
EComprobarRiesgoBanco:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    ValoresDevolucionRemesa = ""
    Set C = Nothing
    Set miRsAux = Nothing
End Function




Private Function ComprobarRiesgoAsignandoVto() As Boolean
Dim C As Collection
Dim I1 As Currency
    
    On Error GoTo EComprobarRiesgoAsignandoVto
    ComprobarRiesgoAsignandoVto = False
    Set miRsAux = New ADODB.Recordset
    Set C = New Collection
    Cad = "Select * from tmpcierre1 where codusu = " & vUsu.Codigo
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        I1 = 0
        
        'Vemos el riesgo de la remesa
        ValoresDevolucionRemesa = DevuelveDesdeBD("remesariesgo", "ctabancaria", "codmacta", miRsAux!Cta, "T")
        
        'Si no he puesto el riesgo NO compruebo nada
        If ValoresDevolucionRemesa <> "" Then
        
        
            For jj = 1 To ListView1.ListItems.Count
               ' WHERE numserie='A' AND codfaccl=99588 AND fecfaccl='2003-12-01' AND numorden=4"
               With ListView1.ListItems(jj)
                    'Si el subitem es del banco
                    If .SubItems(7) = miRsAux!Cta Then
                        I1 = I1 + CCur(.SubItems(5))
                    End If
               End With
            Next
        End If
        If I1 > 0 Then
            Cad = miRsAux!Cta & "|" & miRsAux!nomcta & "|" & ValoresDevolucionRemesa & "|"
            'A remesar ahora
            Cad = Cad & CStr(I1) & "|"
            C.Add Cad
        End If
        
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If C.Count > 0 Then
        If CompruebaImportesRemesa(C) Then ComprobarRiesgoAsignandoVto = True
    Else
        ComprobarRiesgoAsignandoVto = True
    End If
    
EComprobarRiesgoAsignandoVto:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set C = Nothing
End Function






'A partir de un col que llevara el banco
'
' A esta funcion la llamaremos cargando los bancos y los importes a remesar
' La primera vez que lo llamamaos es desde "dividiendo remesa".Ahi lee el importe desde la tabla
' La segunda lo leera ya con los vtos seleccionados
Private Function CompruebaImportesRemesa(ByRef Bancos As Collection) As Boolean
Dim C2 As Collection
Dim I1 As Currency

    On Error GoTo ECompruebaImportesRemesa
    CompruebaImportesRemesa = False


    'Para cada banco voy a ver cuantos vtos tengo para este tipo remesa
    'Tenemos que ver todos los vencimientos que sean de tipo de pago talon o pagare, que la cta de pago sea
    'la del banco en question y ver cuanto llevamos remesado
    Set C2 = New Collection
    For jj = 1 To Bancos.Count
        ValoresDevolucionRemesa = RecuperaValor(Bancos(jj), 1)
        Cad = "select sum(impcobro) FROM scobro,sforpa WHERE scobro.codforpa = sforpa.codforpa AND "
        Cad = Cad & "siturem>'B' AND siturem < 'Z'"
        Cad = Cad & " and ctabanc2='" & ValoresDevolucionRemesa & "' AND tiporem = 1  " '1: efectos
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        ValoresDevolucionRemesa = "0"
        If Not miRsAux.EOF Then
            'Le sumo lo que llevamos en esta remesa (los k estan check) a los vtos ya remesados Y nO eleminidados
            ValoresDevolucionRemesa = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        ValoresDevolucionRemesa = ValoresDevolucionRemesa & "|"
        Cad = Bancos(jj) & ValoresDevolucionRemesa  'Añado el riesgo YA remesado
        C2.Add Cad
    Next
    
    Cad = ""
    For jj = 1 To C2.Count
        'Para cada banco veo los importes
        ValoresDevolucionRemesa = RecuperaValor(C2(jj), 4)
        I1 = CCur(ValoresDevolucionRemesa)   'En esta remesa
        
        ValoresDevolucionRemesa = RecuperaValor(C2(jj), 5)
        I1 = I1 + CCur(ValoresDevolucionRemesa)   '+ YA remesado
        
        ValoresDevolucionRemesa = RecuperaValor(C2(jj), 3)
        If I1 > CCur(ValoresDevolucionRemesa) Then
            '*Importe mayor que el riesgo **
            ValoresDevolucionRemesa = "Importe riesgo: " & Format(ValoresDevolucionRemesa, FormatoImporte)
            ValoresDevolucionRemesa = "Cuenta: " & RecuperaValor(C2(jj), 1) & " - " & RecuperaValor(C2(jj), 2) & vbCrLf & ValoresDevolucionRemesa & vbCrLf
           ' ValoresDevolucionRemesa = ValoresDevolucionRemesa & "Importe riesgo: " & Format(ValoresDevolucionRemesa, FormatoImporte) & vbCrLf
            ValoresDevolucionRemesa = ValoresDevolucionRemesa & "Total : " & Format(I1, FormatoImporte) & vbCrLf
            ValoresDevolucionRemesa = ValoresDevolucionRemesa & "Ya remesado  : " & Format(RecuperaValor(C2(jj), 5), FormatoImporte) & vbCrLf
            ValoresDevolucionRemesa = ValoresDevolucionRemesa & "Esta remesa : " & Format(RecuperaValor(C2(jj), 4), FormatoImporte)
            If Cad <> "" Then Cad = Cad & vbCrLf & vbCrLf
            Cad = Cad & ValoresDevolucionRemesa & vbCrLf
        End If
    Next
    
    If Cad <> "" Then
        'Superado el riesgo. Hago pregunta
        ValoresDevolucionRemesa = String(60, "*") & vbCrLf & vbCrLf
        Cad = ValoresDevolucionRemesa & Cad & ValoresDevolucionRemesa & vbCrLf & vbCrLf
        Cad = Cad & "¿Continuar?"
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then Cad = ""
    End If
    If Cad = "" Then CompruebaImportesRemesa = True

    


ECompruebaImportesRemesa:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set C2 = Nothing
    Cad = ""
End Function




