VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPuntear 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punteo de extractos"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17700
   Icon            =   "frmPuntear.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   17700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDatos 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   885
      Left            =   120
      TabIndex        =   23
      Top             =   780
      Width           =   17505
      Begin VB.CommandButton cmdVer 
         Height          =   375
         Left            =   13200
         Picture         =   "frmPuntear.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ver datos"
         Top             =   210
         Width           =   375
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
         Left            =   1260
         TabIndex        =   0
         Text            =   "Text4"
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox txtDesCta 
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
         Left            =   2910
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   180
         Width           =   4275
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
         Left            =   9120
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   210
         Width           =   1305
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
         Left            =   11850
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   210
         Width           =   1275
      End
      Begin VB.CheckBox chkSin 
         Caption         =   "Incluir sólo apuntes sin punteo"
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
         Left            =   13980
         TabIndex        =   5
         Top             =   600
         Width           =   3435
      End
      Begin VB.CheckBox chkImp 
         Caption         =   "Ordenar apuntes por importe"
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
         Left            =   13980
         TabIndex        =   4
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Cuenta"
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
         Left            =   90
         TabIndex        =   27
         Top             =   210
         Width           =   795
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Left            =   990
         Picture         =   "frmPuntear.frx":0A0E
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha inicio"
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
         Left            =   7560
         TabIndex        =   26
         Top             =   240
         Width           =   1275
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   11610
         Picture         =   "frmPuntear.frx":7260
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha fin"
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
         Index           =   1
         Left            =   10500
         TabIndex        =   25
         Top             =   270
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   8850
         Picture         =   "frmPuntear.frx":72EB
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   21
      Top             =   30
      Width           =   3555
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   22
         Top             =   180
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cuenta Anterior"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cuenta Siguiente"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Saldos"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Importes del Punteo"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar marcas de punteado"
               Object.Tag             =   "0"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Punteo automático por importes"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FramePorImportes 
      Height          =   735
      Left            =   150
      TabIndex        =   16
      Top             =   8850
      Width           =   8715
      Begin VB.CommandButton cmdPorIMportes 
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
         Height          =   435
         Index           =   1
         Left            =   5940
         TabIndex        =   19
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton cmdPorIMportes 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
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
         Index           =   0
         Left            =   7230
         TabIndex        =   18
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Leyendo datos"
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
         Height          =   255
         Index           =   1
         Left            =   3900
         TabIndex        =   20
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Punteo automático por importes"
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
         TabIndex        =   17
         Top             =   300
         Width           =   3615
      End
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
      Left            =   10020
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   9090
      Width           =   2265
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
      Index           =   1
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   9090
      Width           =   2265
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
      Index           =   2
      Left            =   15300
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   9090
      Width           =   2265
   End
   Begin MSComctlLib.ListView lwh 
      Height          =   6525
      Left            =   8880
      TabIndex        =   9
      Top             =   2160
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   11509
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   2645
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Asiento"
         Object.Width           =   1728
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Documento"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ampliación"
         Object.Width           =   4531
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Importe"
         Object.Width           =   2469
      EndProperty
   End
   Begin MSComctlLib.ListView lwD 
      Height          =   6525
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   11509
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   2645
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Asiento"
         Object.Width           =   1728
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Documento"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ampliación"
         Object.Width           =   4531
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Importe"
         Object.Width           =   2469
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   16890
      TabIndex        =   28
      Top             =   150
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
      Index           =   3
      Left            =   10950
      Picture         =   "frmPuntear.frx":7376
      ToolTipText     =   "Puntear al haber"
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   2
      Left            =   10500
      Picture         =   "frmPuntear.frx":74C0
      ToolTipText     =   "Quitar al haber"
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   1800
      Picture         =   "frmPuntear.frx":760A
      ToolTipText     =   "Puntear al Debe"
      Top             =   1770
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   1440
      Picture         =   "frmPuntear.frx":7754
      ToolTipText     =   "Quitar al Debe"
      Top             =   1770
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "DEBE"
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
      Index           =   2
      Left            =   10020
      TabIndex        =   15
      Top             =   8850
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "HABER"
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
      Index           =   3
      Left            =   12510
      TabIndex        =   14
      Top             =   8850
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "SALDO"
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
      Index           =   4
      Left            =   15300
      TabIndex        =   13
      Top             =   8850
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "HABER"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   8880
      TabIndex        =   8
      Top             =   1740
      Width           =   1200
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "DEBE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1740
      Width           =   915
   End
End
Attribute VB_Name = "frmPuntear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 304

Public EjerciciosCerrados As Boolean

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1

Dim Sql As String
Dim RC As String
Dim Rs As Recordset

Dim PrimeraSeleccion As Boolean
Dim ClickAnterior As Byte '0 Empezar 1.-Debe 2.-Haber
Dim I As Integer
Dim De As Currency
Dim Ha As Currency
    
Dim ModoPunteo As Byte
    '0- Punteo normal. El de toda la vida
    '1- Punteo automatico por importes
    
Dim CtasQueHaPunteado As String
    
'Con estas dos variables
Dim ContadorBus As Integer
Dim Checkear As Boolean

Dim ValorAnt As String

Private Sub chkImp_Click()
    If Text3(2).Text <> "" And Text3(0).Text <> "" And Text3(1).Text <> "" Then
        'ConfirmarDatos False
    End If
End Sub

Private Sub chkImp_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkSin_Click()
    If Text3(2).Text <> "" And Text3(0).Text <> "" And Text3(1).Text <> "" Then
        'ConfirmarDatos False
    End If
End Sub

Private Sub chkSin_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub ConfirmarDatos2()
    If Text3(2).Text = "" Then
        MsgBox "Introduzca una cuenta", vbExclamation
        PonleFoco Text3(2)
        Exit Sub
    End If
    If Text3(0).Text = "" Or Text3(1).Text = "" Then
        MsgBox "Introduce las fechas de consulta de extractos", vbExclamation
        Exit Sub
    End If
    If Text3(0).Text <> "" And Text3(1).Text <> "" Then
        If CDate(Text3(0).Text) > CDate(Text3(1).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Sub
        End If
    End If
    
    Sql = ""
    'Llegados aqui. Vemos la fecha y demas
    If Text3(0).Text <> "" Then
        Sql = " fechaent >= '" & Format(Text3(0).Text, FormatoFecha) & "'"
    End If
    
    If Text3(1).Text <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & " fechaent <= '" & Format(Text3(1).Text, FormatoFecha) & "'"
    End If
    Text3(0).Tag = Sql  'Para las fechas
    
    BloqueoManual False, "PUNTEO", ""
    If Not BloqueoManual(True, "PUNTEO", CStr(Abs(EjerciciosCerrados) & Text3(2).Text)) Then
         MsgBox "Imposible acceder a puntear la cuenta. Puede estar bloqueada", vbExclamation
         Exit Sub
    End If

    
    
    espera 0.1
    Me.Refresh
    DoEvents
    CargarDatos_

End Sub



Private Sub cmdPorIMportes_Click(Index As Integer)

    If Index = 1 And AlgunNodoPunteado Then
        If MsgBox("¿Actualizar el punteo en la base de datos?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
        

    

    If Index = 1 Then
        'Actualizamos la BD
        'Pongo ModoPunteo=0 para que ejecute el SQL
        ModoPunteo = 0
        
        'UPDATEAMOS EN LA BD
        'Y volveremos a cargar los datos
        For I = 1 To lwh.ListItems.Count
            If lwh.ListItems(I).Checked Then PunteaEnBD lwh.ListItems(I), False
        Next I
        
        For I = 1 To lwD.ListItems.Count
            If lwD.ListItems(I).Checked Then PunteaEnBD lwD.ListItems(I), True
        Next I
        
        
    Else
        'Quit la seleccion
        For I = 1 To lwD.ListItems.Count
            If lwD.ListItems(I).Checked Then lwD.ListItems(I).Checked = False
        Next I
        For I = 1 To lwh.ListItems.Count
            If lwh.ListItems(I).Checked Then lwh.ListItems(I).Checked = False
        Next I
    End If
    'Limpiamos campos
    
    Text2(0).Text = ""
    Text2(1).Text = ""
    Text2(2).Text = ""
    De = 0: Ha = 0
    


    'Quitamos las posibles marcas
    PonerModoPunteo False
    
    If Index = 1 Then CargarDatos_
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub






Private Sub cmdVer_Click()
    ConfirmarDatos2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Icon = frmppal.Icon

'    Me.framePregunta.Visible = True
    Limpiar Me
    PrimeraSeleccion = True
    Caption = "Punteo de extractos"
    'La toolbar
    With Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 7
        .Buttons(2).Image = 8
        .Buttons(4).Image = 45 '22
        .Buttons(5).Image = 23
        .Buttons(6).Image = 14
        .Buttons(7).Image = 10
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    
    FramePorImportes.visible = False
    FrameDatos.Enabled = True
    FrameBotonGnral.Enabled = False
    
    I = 0
    Text3(0).Text = Format(DateAdd("yyyy", I, vParam.fechaini), "dd/mm/yyyy")
    If Not vParam.FecEjerAct Then I = I + 1
    Text3(1).Text = Format(DateAdd("yyyy", I, vParam.fechafin), "dd/mm/yyyy")
    
    
    CtasQueHaPunteado = ""   'Parar cuando haga el unload
    

    PonerTool False
    
    PonleFoco Text3(2)
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerTool(Activa As Boolean)
    With Toolbar1
        .Buttons(1).Enabled = Activa
        .Buttons(2).Enabled = Activa
        .Buttons(4).Enabled = Activa
        .Buttons(5).Enabled = Activa
        .Buttons(6).Enabled = Activa
        .Buttons(7).Enabled = Activa
    End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
    BloqueoManual False, "PUNTEO", Text3(2).Text
'    VerLogPunteado
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CByte(Image1(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub




Private Sub Image1_Click(Index As Integer)
    Set frmC = New frmCal
    Image1(0).Tag = Index
    If Text3(Index).Text <> "" Then
        frmC.Fecha = CDate(Text3(Index).Text)
    Else
        frmC.Fecha = Now
    End If
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub Image2_Click()
    BloqueoManual False, "PUNTEO", Text3(2).Text
    Me.Text3(2).Text = Text3(2).Text
    Me.txtDesCta.Text = txtDesCta.Text
    PrimeraSeleccion = True
    PonleFoco Text3(2)
End Sub



Private Sub Image4_Click()
    OtraCuenta True
End Sub

Private Sub imgCheck_Click(Index As Integer)

    If Text3(2).Text = "" Then Exit Sub
    
    If Index = 0 Or Index = 1 Then
        If lwD.ListItems.Count = 0 Then Exit Sub
    End If
    If Index = 2 Or Index = 3 Then
        If lwh.ListItems.Count = 0 Then Exit Sub
    End If
    '
    
    If (Index Mod 2) = 0 Then
        RC = "quitar punteos de lo apuntes"
    Else
        RC = "puntear los apuntes"
    End If
    
    Sql = "Seguro que desea " & RC
    If Index > 1 Then
        RC = "HABER"
    Else
        RC = "DEBE"
    End If
    Sql = Sql & " al " & RC & "?"
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    'HA DICHO SI
    
    If Index < 2 Then
        'PUNTEAMOS o DESPUNTEAMOS EL DEBE
        '---------------------------------
        Checkear = True
        If Index = 1 Then Checkear = False
        For I = 1 To lwD.ListItems.Count
            If lwD.ListItems(I).Checked = Checkear Then
                    
                lwD.ListItems(I).Checked = Not lwD.ListItems(I).Checked
                PunteaEnBD lwD.ListItems(I), True
            End If
        Next I
        If Index = 0 Then De = 0
    Else
    
        'PUNTEAMOS o DESPUNTEAMOS EL HABER
        '---------------------------------
        Checkear = True
        If Index = 3 Then Checkear = False
        For I = 1 To lwh.ListItems.Count
            If lwh.ListItems(I).Checked = Checkear Then
                lwh.ListItems(I).Checked = Not lwh.ListItems(I).Checked
                PunteaEnBD lwh.ListItems(I), False
            End If
        Next I
        If Index = 2 Then Ha = 0
    End If
    ContadorBus = 0
    
    If De - Ha <> 0 Then
        Text2(2).Text = Format(De - Ha, FormatoImporte)
    Else
        Text2(2).Text = ""
    End If
    
    
    
End Sub

Private Sub imgCuentas_Click()
    Sql = ""
    Set frmCta = New frmColCtas
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 3  'NUEVO
    frmCta.Show vbModal
    Set frmCta = Nothing
    If Sql <> "" Then
        Text3(2).Text = RecuperaValor(Sql, 1)
        txtDesCta.Text = RecuperaValor(Sql, 2)
        Text3_LostFocus 2
    End If
End Sub



Private Sub lwD_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Screen.MousePointer = vbHourglass
    Set lwD.SelectedItem = Item
    'Ponemos a true o a false
    PunteaEnBD Item, True
    'Comprobamos primero que esta checkeado
    If Item.Checked = False Then
        PrimeraSeleccion = True
        ClickAnterior = 0
    Else
        If ClickAnterior <> 1 Then
            If PrimeraSeleccion Then
                BusquedaEnHaber
                PrimeraSeleccion = False
                ClickAnterior = 1
            Else
                PrimeraSeleccion = True
                ClickAnterior = 0
            End If
        Else
            PrimeraSeleccion = True
        End If
    
    End If
    Screen.MousePointer = vbDefault

End Sub



Private Sub lwh_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Screen.MousePointer = vbHourglass
    Set lwh.SelectedItem = Item
    'Ponemos a true o a false
    PunteaEnBD Item, False
    'Comprobamos primero que esta checkeado
    If Item.Checked = False Then
        PrimeraSeleccion = True
        ClickAnterior = 0
    Else
        If ClickAnterior <> 2 Then
            If PrimeraSeleccion Then
                BusquedaEnDebe
                PrimeraSeleccion = False
                ClickAnterior = 2
            Else
                PrimeraSeleccion = True
                ClickAnterior = 0
            End If
        Else
            PrimeraSeleccion = True
        End If
    
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub





Private Sub Text3_GotFocus(Index As Integer)
    ValorAnt = Text3(Index).Text

    With Text3(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    If Index = 2 Then
        BloqueoManual False, "PUNTEO", Text3(2).Text
    End If
End Sub


Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

'++
Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYFecha KeyAscii, 0
            Case 1:  KEYFecha KeyAscii, 1
            Case 2:  KEYCuentas KeyAscii, 0 'cta contable
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    Image1_Click (Indice)
End Sub


Private Sub KEYCuentas(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgCuentas_Click
End Sub
'++



Private Sub OtraCuenta(mas As Boolean)
Dim B As Boolean
    Screen.MousePointer = vbHourglass
    
    txtDesCta.Text = "OBTENER CUENTA"
    txtDesCta.Refresh
    Screen.MousePointer = vbHourglass
    B = ObtenerCuenta(mas)
    Screen.MousePointer = vbDefault
    If B Then
'
'
'        'Ya tenemos la cuenta
'        If Not BloqueoManual(True, "PUNTEO", CStr(Abs(EjerciciosCerrados) & Text3(2).Text)) Then
'            MsgBox "Imposible acceder a puntear la cuenta. Puede que este bloqueada", vbExclamation
'            Image2_Click
'            Exit Sub
'        End If
        
        'CargarDatos2
        ConfirmarDatos2
    Else
        txtDesCta.Text = ""
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Text3_LostFocus(Index As Integer)

    If Text3(Index).Text = ValorAnt Then Exit Sub

    Select Case Index
        Case 0, 1
            Text3(Index).Text = Trim(Text3(Index).Text)
            If Text3(Index).Text = "" Then Exit Sub
            If Not EsFechaOK(Text3(Index)) Then
                MsgBox "Fecha incorrecta: " & Text3(Index).Text, vbExclamation
                Text3(Index).Text = ""
                PonleFoco Text3(Index)
                Exit Sub
            End If
            'ConfirmarDatos False

        Case 2
            RC = Trim(Text3(Index).Text)
            PonerTool Text3(2).Text <> ""
            If RC = "" Then
                lwD.ListItems.Clear
                lwh.ListItems.Clear
                DoEvents
                txtDesCta.Text = ""
                
                MsgBox "Introduzca una cuenta", vbExclamation
                PonleFoco Text3(2)
                Exit Sub
                
            End If
            If CuentaCorrectaUltimoNivel(RC, Sql) Then
                Text3(Index).Text = RC
                txtDesCta.Text = Sql
                FrameBotonGnral.Enabled = True
                'ConfirmarDatos True
            Else
                MsgBox Sql, vbExclamation
                Text3(Index).Text = ""
                
                lwD.ListItems.Clear
                lwh.ListItems.Clear
                DoEvents
                
                txtDesCta.Text = ""
                PonleFoco Text3(Index)
                Exit Sub
            End If
            PonerTool Text3(2).Text <> ""
            
    End Select
    
    
End Sub





Private Sub CargarDatos_()
        Label5.Caption = "CARGA"
        Label6.Caption = "CARGA"
        Me.Refresh
        Screen.MousePointer = vbHourglass
        CargarDatos
        Screen.MousePointer = vbDefault
        Label5.Caption = "DEBE"
        Label6.Caption = "HABER"
        Label5.Refresh
        Label6.Refresh
End Sub

Private Sub CargarDatos()
Dim ItmX As ListItem
On Error GoTo ECargarDatos

    'Deberiamos bloquear la cuenta en punteos, es decir
    'en alguna tabla poner que se esta punteando la cuenta X


    'Limpiamos listview
    lwD.ListItems.Clear
    lwh.ListItems.Clear
    DoEvents
    
    lwD.ColumnHeaders(3).Text = "Documento"
    lwh.ColumnHeaders(4).Text = "Ampliación"
    
    
    'Resetamos importes punteados
    De = 0
    Ha = 0
    Text2(0).Text = "": Text2(1).Text = "": Text2(2).Text = ""
    
    Set Rs = New ADODB.Recordset
    
    RC = "SELECT numdiari, linliapu, fechaent, numasien, ampconce"
        
    RC = RC & ",timporteD, timporteH,punteada,numdocum"
    RC = RC & " FROM hlinapu"
    
    RC = RC & " WHERE "
    RC = RC & " codmacta ='" & Me.Text3(2).Text & "' AND "
    RC = RC & Text3(0).Tag
    'Si solo mostramos los sin puntear
    If chkSin.Value = 1 Then RC = RC & " AND punteada =0 "
    'Si esta marcado ordenar por importe o no
    If Me.chkImp.Value = 1 Then
        RC = RC & " ORDER BY timported desc,timporteh desc "
    Else
        RC = RC & " ORDER BY fechaent"
    End If
    
    Rs.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 0
    While Not Rs.EOF
        
        If IsNull(Rs!timported) Then
            'Va al haber
            Set ItmX = lwh.ListItems.Add()
            ItmX.SubItems(4) = Format(Rs!timporteH, FormatoImporte)
        Else
            'AL DEBE
            Set ItmX = lwD.ListItems.Add()
            ItmX.SubItems(4) = Format(Rs!timported, FormatoImporte)
        End If
        ItmX.Text = Format(Rs!FechaEnt, "dd/mm/yyyy")
        
        ItmX.SubItems(1) = Rs!NumAsien
        ItmX.SubItems(2) = DBLet(Rs!Numdocum, "T")
        ItmX.SubItems(3) = Rs!Ampconce
        ItmX.ListSubItems(3).ToolTipText = Rs!Ampconce
        
        'En el tag, para actualizaciones i demas pondremos
        'Separado por pipes los valores de numdiari y linliapu
        'claves de la tabla hlinapu
        ItmX.Tag = Rs!NumDiari & "|" & Rs!Linliapu & "|"
        
        'El check
        ItmX.Checked = (Rs!punteada = 1)
        
        
        
        'Siguiente
        Rs.MoveNext
        I = I + 1
        'Por si refrescamos
        If I > 3000 Then
            I = 0
            DoEvents
        End If
    Wend
    Rs.Close

    Exit Sub
ECargarDatos:
        MuestraError Err.Number, "Cargando datos", Err.Description
        Set Rs = Nothing
End Sub






Private Sub BusquedaEnHaber()
    ContadorBus = 1
    Checkear = False
    Do
        I = 1
        While I <= lwh.ListItems.Count
            'Comprobamos k no esta chekeado
            If Not lwh.ListItems(I).Checked Then
                'K tiene el mismo importe
                If lwD.SelectedItem.SubItems(4) = lwh.ListItems(I).SubItems(4) Then
                    'Comprobamos k tienen el mismo DOCUM
                    'Si no es el mismo, pero es la segunda busqueda, tb aceptamos
                    If ContadorBus > 1 Then
                        Checkear = True
                    Else
                        Checkear = (lwD.SelectedItem.SubItems(2) = lwh.ListItems(I).SubItems(2))
                    End If
                
                    If Checkear Then
                        'Tiene el mismo importe y no esta chequeado
                        Set lwh.SelectedItem = lwh.ListItems(I)
                        lwh.SelectedItem.EnsureVisible
                        lwh.SetFocus
                        Beep
                        Exit Sub
                    End If
                End If
            End If
            I = I + 1
        Wend
        ContadorBus = ContadorBus + 1
        Loop Until ContadorBus > 2
End Sub



Private Sub BusquedaEnDebe()
    ContadorBus = 1
    Checkear = False
    Do
        I = 1
        While I <= lwD.ListItems.Count
            If lwh.SelectedItem.SubItems(4) = lwD.ListItems(I).SubItems(4) Then
                'Lo hemos encontrado. Comprobamos que no esta chequeado
                If Not lwD.ListItems(I).Checked Then
                    'Tiene el mismo importe y no esta chequeado
                    'Comprobamos k el docum es el mismo
                    'Si no es el mismo, pero es la segunda busqueda, tb aceptamos
                    If ContadorBus > 1 Then

                        Checkear = True
                    Else
                        Checkear = (lwh.SelectedItem.SubItems(2) = lwD.ListItems(I).SubItems(2))
                    End If
                    If Checkear Then
                        Set lwD.SelectedItem = lwD.ListItems(I)
                        lwD.SelectedItem.EnsureVisible
                        lwD.SetFocus
                        Beep
                        Exit Sub
                    End If
                End If
            End If
            I = I + 1
        Wend
        ContadorBus = ContadorBus + 1
    Loop Until ContadorBus > 2
End Sub



Private Sub PunteaEnBD(ByRef IT As ListItem, EnDEBE As Boolean)
Dim Sng As Currency
On Error GoTo EPuntea
    
        
    
    Sql = "UPDATE hlinapu"
    Sql = Sql & " SET "
    If IT.Checked Then
        RC = "1"
        Sng = 1
        Else
        RC = "0"
        Sng = -1
    End If
    Sng = Sng * CCur(IT.SubItems(4))
    If EnDEBE Then
        De = De + Sng
    Else
        Ha = Ha + Sng
    End If




    Sql = Sql & " punteada = " & RC
    Sql = Sql & " WHERE fechaent='" & Format(IT.Text, FormatoFecha) & "'"
    Sql = Sql & " AND numasien=" & IT.SubItems(1) & " AND numdiari ="
    RC = RecuperaValor(IT.Tag, 1)
    Sql = Sql & RC & " AND linliapu ="
    RC = RecuperaValor(IT.Tag, 2)
    Sql = Sql & RC & ";"
    If ModoPunteo = 0 Then
        Conn.Execute Sql
        InsertarCtaCadenaPunteados
    End If
    
    'Ponemos los importes
    If De <> 0 Then
        Text2(0).Text = Format(De, FormatoImporte)
        Else
        Text2(0).Text = ""
    End If
    If Ha <> 0 Then
        Text2(1).Text = Format(Ha, FormatoImporte)
        Else
        Text2(1).Text = ""
    End If
    Sng = De - Ha
    If Sng <> 0 Then
        Text2(2).Text = Format(Sng, FormatoImporte)
        Else
        Text2(2).Text = ""
    End If
    
    
    Exit Sub
EPuntea:
    MuestraError Err.Number, "Accediendo BD para puntear", Err.Description
End Sub



Private Function ObtenerCuenta(Siguiente As Boolean) As Boolean

    Sql = "select distinct codmacta from hlinapu"
    
    Sql = Sql & " WHERE  fechaent >= '" & Format(Text3(0).Text, FormatoFecha) & "'"
    Sql = Sql & " AND fechaent <= '" & Format(Text3(1).Text, FormatoFecha) & "'"
    Sql = Sql & " AND codmacta "
    If Siguiente Then
        Sql = Sql & ">"
    Else
        Sql = Sql & "<"
    End If
    Sql = Sql & " '" & Text3(2).Text & "'"
    
    If chkSin.Value = 1 Then Sql = Sql & " AND punteada =0 "
    
    Sql = Sql & "  ORDER BY codmacta"
    If Siguiente Then
        Sql = Sql & " ASC"
    Else
        Sql = Sql & " DESC"
    End If
    
    'Para optimizar la velocidad
    Sql = Sql & " limit 0, 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        Sql = "No se ha obtenido la cuenta "
        If Siguiente Then
            Sql = Sql & " siguiente."
        Else
            Sql = Sql & " anterior."
        End If
        MsgBox Sql, vbExclamation
        ObtenerCuenta = False
    Else
        Text3(2).Text = Rs!codmacta
        txtDesCta.Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Rs!codmacta, "T")
        ObtenerCuenta = True
    End If
    Rs.Close
    Set Rs = Nothing
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1
            OtraCuenta False
        Case 2
            OtraCuenta True
        Case 4
            ' en historicoCalculamos saldo. Lleva ya el sql montado
            SaldoHistorico Text3(2).Text, Text3(0).Tag, txtDesCta.Text, False 'EjerciciosCerrados
        Case 5
            Screen.MousePointer = vbHourglass
            HazSumas
            Screen.MousePointer = vbDefault
        Case 6
        
            Screen.MousePointer = vbHourglass
            DesmarcaTodo
            Screen.MousePointer = vbDefault
        
        
        Case 7
            'Comprobamos si hay algun ITEM seleccionado
            If AlgunNodoPunteado Then
                MsgBox "Existen lineas punteadas.  Debe seleccionar solo 'Sin puntear'", vbInformation
                Exit Sub
            End If
            
            Screen.MousePointer = vbHourglass
            
            
            'Punteo automatico por importes
            PonerModoPunteo True
            'Refrescamos el form
            Me.Refresh
            PuntearImportesAutomaticamente
            
            Screen.MousePointer = vbDefault
            DoEvents
    End Select


End Sub


Private Sub HazSumas()
Dim Im As Currency
Dim PuntD As Currency
Dim PuntH As Currency
Dim d As Currency
Dim H As Currency
    On Error GoTo EHazSumas
    d = 0
    H = 0
    PuntD = 0: PuntH = 0
    'Recorremos el debe
    With lwD
        If .ListItems.Count > 0 Then
            For I = 1 To .ListItems.Count
                Im = CCur(ImporteFormateado(.ListItems(I).SubItems(4)))
                If Not .ListItems(I).Checked Then
                 
                    d = d + Im
                    
                Else
                    PuntD = PuntD + Im
                End If
            Next I
        End If
    End With
    
    
    With lwh
        If .ListItems.Count > 0 Then
            For I = 1 To .ListItems.Count
                Im = CCur(ImporteFormateado(.ListItems(I).SubItems(4)))
                If Not .ListItems(I).Checked Then
                    H = H + Im
                Else
                    PuntH = PuntH + Im
                End If
            Next I
        End If
    End With
    
    
    Sql = Format(PuntD, FormatoImporte) & "|" & Format(d, FormatoImporte) & "|" & Format(PuntD + d, FormatoImporte) & "|"
    Sql = Sql & Format(PuntH, FormatoImporte) & "|" & Format(H, FormatoImporte) & "|" & Format(PuntH + H, FormatoImporte) & "|"
    'Las diferencias
    Sql = Sql & Format(PuntD - PuntH, FormatoImporte) & "|" & Format(d - H, FormatoImporte) & "|" & Format((PuntD - PuntH) + (d - H), FormatoImporte) & "|"
    
    frmMensajes.Parametros = Sql
    frmMensajes.Opcion = 18
    frmMensajes.Show vbModal
    
    Exit Sub
EHazSumas:
    
    MuestraError Err.Number, "Realizando sumas Debe/haber"
End Sub


Private Sub DesmarcaTodo()

    Sql = "Va a desmarcar todos los punteos para: " & vbCrLf & vbCrLf
    Sql = Sql & "Cuenta: " & Text3(2).Text & " - " & txtDesCta.Text & vbCrLf
    Sql = Sql & "Fecha inicio: " & Text3(0).Text & vbCrLf
    Sql = Sql & "Fecha fin:     " & Text3(1).Text & vbCrLf & vbCrLf & vbCrLf
    Sql = Sql & "          ¿Desea continuar?"
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Sql = "UPDATE hlinapu"
    Sql = Sql & " SET punteada=0 WHERE codmacta = '" & Text3(2).Text & "'"
    Sql = Sql & " AND fechaent>= '" & Format(Text3(0).Text, FormatoFecha) & "'"
    Sql = Sql & " AND fechaent<= '" & Format(Text3(1).Text, FormatoFecha) & "'"
    Conn.Execute Sql
    InsertarCtaCadenaPunteados
    CargarDatos
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerModoPunteo(ModoImporte As Boolean)

    ModoPunteo = 0
    If ModoImporte Then
        ModoPunteo = 1
    Else
        ModoPunteo = 0
    End If
        
    FramePorImportes.visible = ModoImporte
    FrameDatos.Enabled = Not ModoImporte
    
    With Toolbar1
        .Buttons(1).Enabled = Not ModoImporte
        .Buttons(2).Enabled = Not ModoImporte
        .Buttons(4).Enabled = Not ModoImporte
        .Buttons(5).Enabled = Not ModoImporte
        .Buttons(6).Enabled = Not ModoImporte
        .Buttons(7).Enabled = Not ModoImporte
    End With
    
    
    For I = 0 To 3
        imgCheck(I).visible = Not ModoImporte
    Next I
End Sub



Private Function AlgunNodoPunteado() As Boolean

    AlgunNodoPunteado = True
    
    For I = 1 To lwD.ListItems.Count
        If lwD.ListItems(I).Checked Then Exit Function
    Next I
    For I = 1 To lwh.ListItems.Count
        If lwh.ListItems(I).Checked Then Exit Function
    Next I
    'Si llega aqui es que NO hay ninguno punteado
    AlgunNodoPunteado = False
End Function


Private Sub PuntearImportesAutomaticamente()
Dim J As Integer
Dim T1 As Single


    T1 = Timer - 1
    For I = 1 To lwD.ListItems.Count
        'Label
        Sql = lwD.ListItems(I).SubItems(4) 'Cargo el importe
        
        If Timer - T1 > 1 Then
            Me.Label7(1).visible = Not Me.Label7(1).visible
            If Me.Label7(1).visible Then Me.Label7(1).Refresh
            T1 = Timer
        End If
        
        For J = 1 To lwh.ListItems.Count
            If Not lwh.ListItems(J).Checked Then
                RC = lwh.ListItems(J).SubItems(4)
                If Sql = RC Then
                    'EUREKA!!!!!! El mismo importe
                    lwD.ListItems(I).Checked = True
                    PunteaEnBD lwD.ListItems(I), True
                    lwh.ListItems(J).Checked = True
                    PunteaEnBD lwh.ListItems(J), False
                    'Nos salimos del for
                    Exit For
                End If
            End If
        Next J
    Next I
    
    Me.Label7(1).visible = False
End Sub

'-------------------------------------------------------
'-------------------------------------------------------
'Para el LOG de punteo de cuentas
'-------------------------------------------------------
'-------------------------------------------------------
Private Sub InsertarCtaCadenaPunteados()
Dim Aux As String

    Aux = Me.Text3(2).Text & "|"
    If InStr(1, CtasQueHaPunteado, Aux) = 0 Then CtasQueHaPunteado = CtasQueHaPunteado & Aux
        
End Sub


Private Sub VerLogPunteado()

    On Error GoTo Evl
    If CtasQueHaPunteado <> "" Then
        CtasQueHaPunteado = Replace(CtasQueHaPunteado, "|", " ")
        CtasQueHaPunteado = "Cuentas punteadas: " & CtasQueHaPunteado
        vLog.Insertar 17, vUsu, CtasQueHaPunteado
    End If
    
    Exit Sub
Evl:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub
