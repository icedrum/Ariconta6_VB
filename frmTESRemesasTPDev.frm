VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESRemesasTPDev 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "1"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15690
   Icon            =   "frmTESRemesasTPDev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   15690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5610
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameDevlucionRe 
      Height          =   9195
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   15315
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8970
         TabIndex        =   28
         Tag             =   "Importe|N|N|||reclama|importes|||"
         Top             =   8700
         Width           =   1815
      End
      Begin VB.Frame FrameConcepto 
         Caption         =   "Datos Contabilizaci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   300
         TabIndex        =   14
         Top             =   1800
         Width           =   14895
         Begin VB.ComboBox CmbDevol 
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
            ItemData        =   "frmTESRemesasTPDev.frx":000C
            Left            =   2340
            List            =   "frmTESRemesasTPDev.frx":001F
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Tag             =   "Ampliacion debe/CLIENTES|N|N|0||stipoformapago|ampdecli|||"
            Top             =   2010
            Width           =   4830
         End
         Begin VB.ComboBox Combo2 
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
            ItemData        =   "frmTESRemesasTPDev.frx":0093
            Left            =   9450
            List            =   "frmTESRemesasTPDev.frx":00A9
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Tag             =   "Ampliacion debe/CLIENTES|N|N|0||stipoformapago|ampdecli|||"
            Top             =   1500
            Width           =   2820
         End
         Begin VB.TextBox txtImporte 
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
            Left            =   5970
            TabIndex        =   2
            Text            =   "Text4"
            Top             =   930
            Width           =   1125
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
            Index           =   11
            Left            =   9450
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   600
            Width           =   1275
         End
         Begin VB.OptionButton optDevRem 
            Caption         =   "� x Recibo"
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
            Left            =   780
            TabIndex        =   21
            Top             =   810
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton optDevRem 
            Caption         =   "% x Recibo"
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
            Left            =   780
            TabIndex        =   20
            Top             =   1170
            Width           =   2205
         End
         Begin VB.TextBox txtImporte 
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
            Left            =   6360
            TabIndex        =   19
            Text            =   "Text4"
            Top             =   1500
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.OptionButton optDevRem 
            Caption         =   "% x  rec. con m�nimo"
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
            Left            =   780
            TabIndex        =   18
            Top             =   1545
            Width           =   2535
         End
         Begin VB.TextBox txtDConcpeto 
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
            Left            =   10050
            TabIndex        =   17
            Text            =   "Text9"
            Top             =   1050
            Width           =   4725
         End
         Begin VB.TextBox txtConcepto 
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
            Left            =   9450
            TabIndex        =   4
            Text            =   "Text10"
            Top             =   1050
            Width           =   525
         End
         Begin VB.CheckBox chkAgrupadevol2 
            Caption         =   "Agrupa apunte banco"
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
            Left            =   11970
            TabIndex        =   16
            Top             =   600
            Width           =   2475
         End
         Begin VB.TextBox txtImporte 
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
            Index           =   5
            Left            =   9450
            TabIndex        =   5
            Text            =   "Text4"
            Top             =   1950
            Width           =   1245
         End
         Begin VB.Label Label7 
            Caption         =   "Ampliaci�n"
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
            Height          =   345
            Index           =   1
            Left            =   7410
            TabIndex        =   34
            Top             =   1560
            Width           =   1830
         End
         Begin VB.Label Label7 
            Caption         =   "Gastos Banco"
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
            Height          =   345
            Index           =   0
            Left            =   7380
            TabIndex        =   33
            Top             =   2010
            Width           =   1740
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Motivo devoluci�n"
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
            Index           =   50
            Left            =   420
            TabIndex        =   31
            Top             =   2070
            Width           =   1785
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "C�culo Gastos Devoluci�n Cliente"
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
            Index           =   0
            Left            =   390
            TabIndex        =   26
            Top             =   390
            Width           =   3630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Importe Gasto (�)"
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
            Left            =   3720
            TabIndex        =   25
            Top             =   930
            Width           =   1785
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   11
            Left            =   9180
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Devoluci�n"
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
            Index           =   4
            Left            =   7380
            TabIndex        =   24
            Top             =   600
            Width           =   1740
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Importe Minimo (�)"
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
            Index           =   8
            Left            =   3720
            TabIndex        =   23
            Top             =   1515
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Image imgConcepto 
            Height          =   240
            Index           =   1
            Left            =   9180
            Top             =   1110
            Width           =   240
         End
         Begin VB.Label Label7 
            Caption         =   "Concepto apunte"
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
            Height          =   345
            Index           =   9
            Left            =   7380
            TabIndex        =   22
            Top             =   1110
            Width           =   1740
         End
         Begin VB.Label lblAsiento 
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
            Left            =   2550
            TabIndex        =   15
            Top             =   1440
            Width           =   4095
         End
      End
      Begin VB.Frame FrameDevDesdeRemesa 
         Height          =   1185
         Left            =   270
         TabIndex        =   10
         Top             =   540
         Width           =   3585
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
            Index           =   5
            Left            =   990
            TabIndex        =   0
            Text            =   "Text3"
            Top             =   570
            Width           =   915
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
            Index           =   6
            Left            =   2430
            TabIndex        =   1
            Text            =   "Text3"
            Top             =   570
            Width           =   915
         End
         Begin VB.Image imgRem 
            Height          =   240
            Index           =   1
            Left            =   1050
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Remesa"
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
            Left            =   120
            TabIndex        =   13
            Top             =   210
            Width           =   885
         End
         Begin VB.Label Label6 
            Caption         =   "C�digo"
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
            Index           =   5
            Left            =   240
            TabIndex        =   12
            Top             =   585
            Width           =   705
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "A�o"
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
            Index           =   6
            Left            =   1830
            TabIndex        =   11
            Top             =   585
            Width           =   555
         End
      End
      Begin VB.CommandButton cmdCancelar 
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
         Height          =   375
         Index           =   9
         Left            =   13920
         TabIndex        =   7
         Top             =   8730
         Width           =   1215
      End
      Begin VB.CommandButton cmdDevolRem 
         Caption         =   "Devolucion"
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
         Left            =   12510
         TabIndex        =   6
         Top             =   8730
         Width           =   1335
      End
      Begin MSComctlLib.ListView lwCobros 
         Height          =   3915
         Left            =   300
         TabIndex        =   27
         Top             =   4710
         Width           =   14835
         _ExtentX        =   26167
         _ExtentY        =   6906
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
         NumItems        =   0
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
         Left            =   7740
         TabIndex        =   29
         Top             =   8700
         Width           =   1575
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   14430
         Picture         =   "frmTESRemesasTPDev.frx":0137
         ToolTipText     =   "Quitar al Debe"
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   14790
         Picture         =   "frmTESRemesasTPDev.frx":0281
         ToolTipText     =   "Puntear al Debe"
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "DEVOLUCION REMESA"
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
         Index           =   3
         Left            =   5100
         TabIndex        =   9
         Top             =   210
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmTESRemesasTPDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As Byte
    '
    '9.- Devolucion remesa
        
    '16.- Devolucion remesa desde fichero del banco
    
    '28 .- Devolucion remesa desde un vto
    
    
    
    
Public SubTipo As Byte

    'Para la opcion 22
    '   Remesas cancelacion cliente.
    '       1:  Efectos
    '       2: Talones pagares
    
'Febrero 2010
'Cuando pago proveedores con un talon, y le he indicado el numero
Public NumeroDocumento As String
    
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmRe As frmTESRemesas
Attribute frmRe.VB_VarHelpID = -1
Private WithEvents frmB As frmBasico
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1


Dim RS As ADODB.Recordset
Dim SQL As String
Dim I As Integer
Dim IT As ListItem  'Comun
Dim PrimeraVez As Boolean
Dim Cancelado As Boolean
Dim CuentasCC As String
Dim ImporteQueda As Currency

Dim vRemesa As String
Dim ValoresDevolucionRemesa As String
Dim ImporteRemesa As Currency
Dim vSql As String
Dim OpcionAnt As Integer

Dim Remesa As Long
Dim A�oRem As Long
Dim BancoRem As String



Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 21 Or Index = 25 Or Index = 31 Then CadenaDesdeOtroForm = "" 'ME garantizo =""
    If Index = 31 Then
        If MsgBox("�Cancelar el proceso?", vbQuestion + vbYesNo) = vbYes Then SubTipo = 0
    End If
    Unload Me
End Sub

Private Sub CargaCombo()

    CmbDevol.Clear
    
    Set RS = New ADODB.Recordset
    SQL = "select codigo, descripcion from usuarios.wdevolucion order by codigo"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
        CmbDevol.AddItem RS!Codigo & "-" & RS!Descripcion
        CmbDevol.ItemData(CmbDevol.NewIndex) = I
        I = I + 1
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing


End Sub

Private Sub cmdDevolRem_Click()
Dim Importe As Currency
Dim GastoDevolGral As Currency
Dim CadenaVencimiento As String
Dim MultiRemesaDevuelta As String
Dim TipoFicheroDevolucion As Byte

    If Text3(5).Text <> "" Then Opcion = 9
    
    
    SQL = ""
    
    If Text1(11).Text = "" Then SQL = "Ponga la fecha de abono"
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    'Fecha pertenece a ejercicios contbles
    If FechaCorrecta2(CDate(Text1(11).Text), True) > 1 Then Exit Sub
    
    
    If txtImporte(1).Text = "" Then
        MsgBox "Indique el gasto por recibo", vbExclamation
        Exit Sub
    End If
    '
    If Me.optDevRem(2).Value Then
        If (txtImporte(2).Text = "") Then
            MsgBox "Debe poner valores del  minimo", vbExclamation
            Exit Sub
        End If
        
    End If
    
    If txtImporte(1).Text <> "" Then
        'Hay gravamen por gastos
        'Bloqueariamos la opcion de modificar esa remesa
        Importe = TextoAimporte(txtImporte(1).Text)
        If Me.optDevRem(1).Value Or Me.optDevRem(2).Value Then
            'Porcentual. No puede ser superior al 100%
            If Importe > 100 Then
                MsgBox "Importe no puede ser superior al 100%", vbExclamation
                Exit Sub
            End If
        End If
        
    Else
        Importe = 0
    End If
    
    'Comprobamos los conceptos y ampliaciones
    SQL = ""
    If txtConcepto(1).Text <> "" Then
        If txtDConcpeto(1).Text = "" Then SQL = "Concepto"
    Else
        SQL = "Debe introducir un concepto. Revise."
    End If
    
    
    If SQL = "" Then
        If Combo2(0).ListIndex = -1 Then
            SQL = "Ampliacion concepto incorrecta"
        End If
    End If
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    'Nuevo Noviembre 2009
    GastoDevolGral = 0
    GastoDevolGral = TextoAimporte(txtImporte(5).Text)
    
    'Ahora miramos la remesa. En que sitaucion , y de que tipo es
    
    If Opcion = 9 Then
    
        SQL = "Select * from remesas where codigo =" & Text3(5).Text
        SQL = SQL & " AND anyo =" & Text3(6).Text
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RS.EOF Then
            SQL = "Ninguna remesa con esos valores."
            If Opcion = 16 Then SQL = SQL & "  Remesa: " & Text3(5).Text & " / " & Text3(6).Text
            MsgBox SQL, vbExclamation
            RS.Close
            Set RS = Nothing
            Exit Sub
        End If
        
        
        'Tiene valor
        If RS!Situacion = "A" Then
            MsgBox "Remesa abierta. Sin llevar al banco.", vbExclamation
            RS.Close
            Set RS = Nothing
            Exit Sub
        End If
        
        
        
        If Asc(RS!Situacion) < Asc("Q") Then
            MsgBox "Remesa sin contabilizar.", vbExclamation
            RS.Close
            Set RS = Nothing
            Exit Sub
        End If
        
        
    
        SQL = RS!Codigo & "|" & RS!Anyo & "|" & RS!codmacta & "|" & Text1(11).Text & "|"
    Else
        SQL = Remesa & "|" & A�oRem & "|" & BancoRem & "|" & Text1(11).Text & "|"
    End If
    
    Importe = TextoAimporte(txtImporte(1).Text)   ''Levara el gasto por recibo
    If Me.optDevRem(1).Value Or Me.optDevRem(2).Value Then SQL = SQL & "%"
    SQL = SQL & "|"
    If Me.optDevRem(2).Value Then SQL = SQL & TextoAimporte(txtImporte(2).Text)
    SQL = SQL & "|"
    
    
    'SQL llevara hasta ahora
    '        remes    cta ban  fec contb tipo gasto el 1: si tiene valor es el minimo por recibo
    ' Ej:    1|2009|572000005|20/11/2009|%|1|
    
    
    'Si contabilizamos el gasto, o pro contra vendra como factura bancaria desde otro lugar(norma34 p.e.)
    If GastoDevolGral = 0 Then
        'NO HAY GASTO
        SQL = SQL & "0|"
    Else
        SQL = SQL & CStr(GastoDevolGral) & "|"
        If ComprobarCero(txtImporte(5).Text) <> 0 Then
            'Voy a contabi�izar los gastos.
            'Vere si tiene CC
            If vParam.autocoste Then
                If DevuelveDesdeBD("codccost", "bancos", "codmacta", RS!codmacta, "T") = "" Then
                    MsgBox "Va a contabilizar los gastos pero no esta configurado el Centro de coste para el banco: " & RS!codmacta, vbExclamation
                    RS.Close
                    Set RS = Nothing
                    Exit Sub
                End If
            End If
        End If
    End If
    
    'Depues del gasto
    'A�adire el fichero, si es autmatico
    SQL = SQL & "|"
    'Nov 2012. En las devoluciones puede ser que el fichero traiga mas de una devolucion
    SQL = SQL & "|"
    

    
    
    'Bloqueamos la devolucion
    BloqueoManual True, "Devolrem", vUsu.Codigo
    'Hacemos la devolucion
    vRemesa = SQL
    ImporteRemesa = Importe
    
    
    SQL = txtConcepto(1).Text & "|" & Combo2(0).ListIndex & "|"
    'y el banco
    'Agrupa el apunte del banco
    SQL = SQL & Abs(chkAgrupadevol2.Value) & "|"
    
    vSql = CadenaVencimiento
    
    DevolverRemesa

    'Desbloqueamos
    BloqueoManual False, "Devolrem", vUsu.Codigo

End Sub

Private Sub DevolverRemesa()
Dim cad As String
Dim jj As Integer
Dim Aux As String

    cad = ""
    For jj = 1 To Me.lwCobros.ListItems.Count
        If lwCobros.ListItems(jj).Checked Then
            cad = cad & "1"
        End If
    Next jj
    If cad = "" Then
        MsgBox "Seleccione los efectos devueltos", vbExclamation
        Exit Sub
    End If
    cad = Len(cad) & " efecto(s)."
    
    'Llegado aqui hago la pregunta
    cad = "Va a realizar la devoluci�n de " & cad & vbCrLf
    If Text1(4).Text <> "" Then
        cad = cad & vbCrLf & "Importe total de la devoluci�n: "
        cad = cad & Text1(4).Text & "�" & vbCrLf & vbCrLf
    End If
    
    Aux = RecuperaValor(vRemesa, 5)
    If optDevRem(1).Value Then
        Aux = "Porcentaje por recibo: " & txtImporte(1) & "%" & vbCrLf
        If txtImporte(2) <> "" Then
            Aux = Aux & "Gasto m�nimo: " & txtImporte(2) & " �" & vbCrLf
        End If
    Else
        Aux = "Gasto por recibo: " & txtImporte(1) & " �" & vbCrLf
    End If
    
    cad = cad & Aux & vbCrLf
    
    'Gasto tramitacion devolucion
    Aux = txtImporte(5)
    If Aux <> "" Then
        Aux = "Gasto bancario : " & Aux & "�" & vbCrLf
        cad = cad & vbCrLf & Aux
    End If
    
    cad = cad & vbCrLf & "�Desea continuar?"
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    If Not RealizarDevolucion Then Exit Sub

    Unload Me

End Sub

Private Function RealizarDevolucion() As Boolean
Dim IncPorcentaje As Boolean
Dim Gasto As Currency
Dim Minimo As Currency
Dim cad As String
Dim jj As Long
Dim CtaBan As String

    RealizarDevolucion = False
    'Tipo de aumento del gasto de devolucion
    cad = RecuperaValor(vRemesa, 5)
    If optDevRem(1).Value Then
        'Porcentual
        IncPorcentaje = True
        Minimo = 0
        cad = txtImporte(2).Text 'RecuperaValor(vRemesa, 6)
        If cad <> "" Then Minimo = cad
    Else
        IncPorcentaje = False
    End If
        
    
    vSql = "DELETE FROM tmpfaclin WHERE codusu =" & vUsu.Codigo
    Conn.Execute vSql
    '                                               numero        serie     vto
    vSql = "INSERT INTO tmpfaclin (codusu, codigo, Numfac, Fecha, numserie, NIF,  "
    vSql = vSql & "Imponible,  ImpIVA,total,cta,cliente,ctabase) VALUES (" & vUsu.Codigo & ","
    For jj = 1 To lwCobros.ListItems.Count
        If Me.lwCobros.ListItems(jj).Checked Then
                                        'cdofaccl
            cad = jj & "," & Val(lwCobros.ListItems(jj).SubItems(1)) & ",'"
                                    'fecfaccl                                                   SERIE
            cad = cad & Format(lwCobros.ListItems(jj).Tag, FormatoFecha) & "','" & lwCobros.ListItems(jj).Text
                                    'numvencimiento numorden
            cad = cad & "'," & Val(lwCobros.ListItems(jj).SubItems(2)) & ","
            ImporteQueda = ImporteFormateado(lwCobros.ListItems(jj).SubItems(5))
            cad = cad & TransformaComasPuntos(CStr(ImporteQueda)) & ","
            
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
            cad = cad & TransformaComasPuntos(CStr(Gasto)) & ","
            cad = cad & TransformaComasPuntos(CStr(ImporteQueda)) & ",'"
            'Cuenta cliente, y banco
            cad = cad & lwCobros.ListItems(jj).SubItems(3) & "','"
            cad = cad & RecuperaValor(vRemesa, 3) & "','"
            If Opcion = 16 Then
                cad = cad & Mid(lwCobros.ListItems(jj).SubItems(11), 1, 4)
            Else
                cad = cad & Mid(CmbDevol.Text, 1, 4)
            End If
            cad = cad & "')"
            cad = vSql & cad
            If Not Ejecuta(cad) Then Exit Function
            
            CtaBan = RecuperaValor(vRemesa, 3)

            
        End If
    Next jj
    
    
    'OK. Ya tengo grabada la temporal con los recibos que devuelvo. Ahora
    'hare:
    '       - generar un asiento con los datos k devuelvo
    '       - marcar los cobros como devueltos, a�adirle el gasto y insertar en la
    '           tabla de hco de devueltos
    
    jj = Val(txtImporte(5).Text) 'Val(RecuperaValor(vRemesa, 7))
    
    If jj = 0 Then
        'Como no se contabilizan los beneficios no hace falta que calcule nada
        cad = ""
     Else
        'Vya obteneer la cuenta de gastos bancarios
        cad = RecuperaValor(vRemesa, 3)  'cta contable
        cad = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", cad, "T")
        If cad = "" Then
            'NO esta configurada
            'Veo si esta en parametros
            'ctabenbanc
            cad = DevuelveDesdeBD("ctabenbanc", "paramtesor", "codigo", "1", "N")
        End If
        If cad = "" Then
            MsgBox "No esta configurada la gastos  bancarios", vbExclamation
            Exit Function
        End If
    End If
    
    ValoresDevolucionRemesa = txtConcepto(1).Text & "|" & Combo2(0).ListIndex & "|"
    
    If Opcion = 9 Then
        vRemesa = Text3(5).Text & "|" & Text3(6).Text & "|" & BancoRem & "|" & Text1(11).Text & "|"
    Else
        vRemesa = Remesa & "|" & A�oRem & "|" & BancoRem & "|" & Text1(11).Text & "|"
    End If
    
    If optDevRem(1).Value Then
        vRemesa = vRemesa & "%|"
    Else
        vRemesa = vRemesa & "|"
    End If
    
    vRemesa = vRemesa & txtImporte(2).Text & "|" & txtImporte(5).Text & "||"
    
    Select Case Opcion
        Case 9
            vRemesa = vRemesa & "Remesa: " & Text3(5).Text & "/" & Text3(6).Text
    End Select
    vRemesa = vRemesa & "|2|"
    
    Dim CodDev As String
    CodDev = ""
    If CmbDevol.ListIndex <> -1 Then CodDev = Mid(CmbDevol.List(CmbDevol.ListIndex), 1, 4)
    
    If RealizarDevolucionRemesa(CDate(Text1(11)), jj > 0, CtaBan, vRemesa, ValoresDevolucionRemesa) Then
        RealizarDevolucion = True
        Screen.MousePointer = vbHourglass
        Screen.MousePointer = vbDefault
    End If
End Function




Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
            Case 9
                PonerFocoLw lwCobros
            Case 16, 28
                PonerFoco Text3(5)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
    Limpiar Me
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    
    
    'Cago los iconos
    CargaImagenesAyudas Me.Image1, 2
    CargaImagenesAyudas imgRem, 1, "Seleccionar remesa"
    CargaImagenesAyudas imgConcepto, 1, "Concepto"
    


    Select Case Opcion
    Case 9, 16, 28
        If SubTipo = 1 Then
            Caption = "EFECTOS"
        Else
            Caption = "TALONES / PAGARES"
        End If
        FrameDevlucionRe.Visible = True
        Caption = "Devolucion remesa (" & UCase(Caption) & ")"
        W = FrameDevlucionRe.Width
        H = FrameDevlucionRe.Height
        Text1(11).Text = Format(Now, "dd/mm/yyyy")
        txtImporte(1).Text = 0
        
        'FALTA####
        'El combo no es igual que el de los tipos de pago.
        
        'Ofertamos el haber de la forma de pago recibo bancario
        SQL = DevuelveDesdeBD("amphacli", "tipofpago", "tipoformapago", "4")
        If SQL <> "" Then Combo2(0).ListIndex = CInt(SQL)
            
    End Select
    
    OpcionAnt = 0
    CargaCabecera
    
    If NumeroDocumento <> "" Then
        Text3(5).Text = RecuperaValor(NumeroDocumento, 1)
        Text3(6).Text = RecuperaValor(NumeroDocumento, 2)
        Text3_LostFocus (5)
    End If
    
    CargaCombo
    
    Remesa = 0
    A�oRem = 0
    
    Me.Height = H + 560
    Me.Width = W + 90
    
End Sub

Private Sub CargaCabecera()
    
    
    If OpcionAnt = Opcion Then Exit Sub
    
    
    lwCobros.ColumnHeaders.Clear
    
    If Opcion = 9 Or Opcion = 28 Then
        lwCobros.ColumnHeaders.Add , , "Serie", 720
        lwCobros.ColumnHeaders.Add , , "Factura", 1140
        lwCobros.ColumnHeaders.Add , , "Vto", 650
        lwCobros.ColumnHeaders.Add , , "Cuenta", 1500
        lwCobros.ColumnHeaders.Add , , "Cliente", 3800
        lwCobros.ColumnHeaders.Add , , "Importe", 1950, 1
        lwCobros.ColumnHeaders.Add , , "FechaOrden", 0
        lwCobros.ColumnHeaders.Add , , "ImporteOrden", 0
        lwCobros.ColumnHeaders.Add , , "Remesa", 1000, 1
        lwCobros.ColumnHeaders.Add , , "A�o", 800
        lwCobros.ColumnHeaders.Add , , "Banco", 1500
    Else
        ' en el caso de devolucion desde fichero mostramos el codigo de devolucion
        lwCobros.ColumnHeaders.Add , , "Serie", 720
        lwCobros.ColumnHeaders.Add , , "Factura", 1140
        lwCobros.ColumnHeaders.Add , , "Vto", 650
        lwCobros.ColumnHeaders.Add , , "Cuenta", 1500
        lwCobros.ColumnHeaders.Add , , "Cliente", 3800
        lwCobros.ColumnHeaders.Add , , "Importe", 1950, 1
        
        lwCobros.ColumnHeaders.Add , , "FechaOrden", 0
        lwCobros.ColumnHeaders.Add , , "ImporteOrden", 0
        lwCobros.ColumnHeaders.Add , , "Remesa", 0, 1
        lwCobros.ColumnHeaders.Add , , "A�o", 0
        lwCobros.ColumnHeaders.Add , , "Banco", 0
        
        lwCobros.ColumnHeaders.Add , , "Devoluci�n", 4000, 0
        
    
    End If

    OpcionAnt = Opcion

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set RS = Nothing
    Set miRsAux = Nothing
    
    NumeroDocumento = "" 'Para reestrablecerlo siempre
End Sub



Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(Image1(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtConcepto(1).Text = RecuperaValor(CadenaSeleccion, 1)
        txtConcepto_LostFocus 1
    End If
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text1(Index).Text <> "" Then frmC.Fecha = CDate(Text1(Index).Text)
    Image1(0).Tag = Index
    frmC.Show vbModal
    Set frmC = Nothing
    If Text1(Index).Text <> "" Then PonerFoco Text1(Index)
End Sub


Private Sub PonerFoco(ByRef o As Object)
    On Error Resume Next
    o.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ObtenerFoco(ByRef T As TextBox)
    T.SelStart = 0
    T.SelLength = Len(T.Text)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub imgCheck_Click(Index As Integer)

    If Index < 2 Then
        'Selecciona forma de pago
        For I = 1 To Me.lwCobros.ListItems.Count
            If Remesa <> 0 And A�oRem <> 0 Then
                If Index = 1 And (Remesa <> lwCobros.ListItems(I).SubItems(8) Or A�oRem <> lwCobros.ListItems(I).SubItems(9)) Then
                    MsgBox "No pueden ser cobros de distintas remesas.", vbExclamation
                    lwCobros.ListItems(I).Checked = False
                    Exit Sub
                End If
            Else
                If Index = 1 Then
                    Remesa = lwCobros.ListItems(I).SubItems(8)
                    A�oRem = lwCobros.ListItems(I).SubItems(9)
                    BancoRem = lwCobros.ListItems(I).SubItems(10)
                End If
            End If
            Me.lwCobros.ListItems(I).Checked = Index = 1
        Next
    End If
    
    CalcularTotal
End Sub

Private Sub imgConcepto_Click(Index As Integer)
  
    Set frmCon = New frmConceptos
    frmCon.DatosADevolverBusqueda = "0|"
    frmCon.Show vbModal
    Set frmCon = Nothing
    
End Sub


Private Sub imgRem_Click(Index As Integer)
    I = Index
    Set frmRe = New frmTESRemesas
    frmRe.Tipo = SubTipo  'Para abrir efectos o talonesypagares
    frmRe.DatosADevolverBusqueda = "1|"
    frmRe.Show vbModal
    Set frmRe = Nothing
    'Por si ha puesto los datos
    CamposRemesaAbono
    
End Sub

Private Sub lwCobros_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim I As Currency

    If Remesa <> 0 And A�oRem <> 0 Then
        If Item.Checked And (Remesa <> Item.SubItems(8) Or A�oRem <> Item.SubItems(9)) Then
            MsgBox "No pueden ser cobros de distintas remesas.", vbExclamation
            Item.Checked = False
            Exit Sub
        End If
    Else
        If Item.Checked Then
            Remesa = Item.SubItems(8)
            A�oRem = Item.SubItems(9)
            BancoRem = Item.SubItems(10)
        End If
    End If

    CalcularTotal


End Sub


Private Sub CalcularTotal()
Dim I As Integer
Dim Total As Currency

    Total = 0
    For I = 1 To Me.lwCobros.ListItems.Count
        If Me.lwCobros.ListItems(I).Checked Then
            Total = Total + Me.lwCobros.ListItems(I).SubItems(5)
        End If
    Next I
    
    If Total <> 0 Then
        Me.Text1(4).Text = Format(Total, FormatoImporte)
    Else
        Text1(4).Text = ""
    End If
    
End Sub


Private Sub optDevRem_Click(Index As Integer)
    txtImporte(2).Visible = Index = 2
    Label4(8).Visible = Index = 2
    If Index <> 0 Then
        Label4(3).Caption = "% aplicado"
    Else
        Label4(3).Caption = "Gastos (�)"
    End If
End Sub

Private Sub optDevRem_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then Exit Sub
    
    If Not EsFechaOK(Text1(Index)) Then
        MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
        Text1(Index).Text = ""
        PonerFoco Text1(Index)
    End If
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    With Text3(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Text3_LostFocus(Index As Integer)
    With Text3(Index)
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        
        If Not IsNumeric(.Text) Then
            MsgBox "Campo debe ser num�rico: " & .Text, vbExclamation
            .Text = ""
            PonerFoco Text3(Index)
        Else
            Opcion = 9
            LimpiarLin Me, "FrameDevDesdeFichero"
            LimpiarLin Me, "FrameDevDesdeVto"
            
            If Text3(5).Text <> "" And Text3(6).Text <> "" Then
                If RemesaCorrecta Then
                    CargaList
                Else
                    Text3(5).Text = ""
                    Text3(6).Text = ""
                End If
            End If
        End If
        
        'Para que vaya a la tabla y traiga datos remesa
'        If Index = 3 Or Index = 4 Then CamposRemesaAbono
    End With
End Sub

Private Function RemesaCorrecta() As Boolean
        
    On Error GoTo eRemesaCorrecta
        
     RemesaCorrecta = False
        
        SQL = "Select * from remesas where codigo =" & Text3(5).Text
        SQL = SQL & " AND anyo =" & Text3(6).Text
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RS.EOF Then
            SQL = "Ninguna remesa con esos valores."
            If Opcion = 16 Then SQL = SQL & "  Remesa: " & Text3(5).Text & " / " & Text3(6).Text
            MsgBox SQL, vbExclamation
            RS.Close
            Set RS = Nothing
            Exit Function
        End If
        
        
        'Tiene valor
        If RS!Situacion = "A" Then
            MsgBox "Remesa abierta. Sin llevar al banco.", vbExclamation
            RS.Close
            Set RS = Nothing
            Exit Function
        End If
        
        
        
        If Asc(RS!Situacion) < Asc("Q") Then
            MsgBox "Remesa sin contabilizar.", vbExclamation
            RS.Close
            Set RS = Nothing
            Exit Function
        End If
        RemesaCorrecta = True
eRemesaCorrecta:
    

End Function

Private Sub CargarValores()
Dim Importe As Currency
Dim GastoDevolGral As Currency
Dim CadenaVencimiento As String
Dim MultiRemesaDevuelta As String
Dim TipoFicheroDevolucion As Byte
    
    MultiRemesaDevuelta = ""
'    CadenaVencimiento = ""
    Select Case Opcion
        Case 9
            SQL = "Select * from remesas where codigo =" & Text3(5).Text
            SQL = SQL & " AND anyo =" & Text3(6).Text
            SQL = SQL & " AND situacion = 'Q'"
        
            
    End Select
    
    
    
    Select Case Opcion
        Case 9
            SQL = "Select * from remesas where codigo =" & Text3(5).Text
            SQL = SQL & " AND anyo =" & Text3(6).Text
            SQL = SQL & " AND situacion = 'Q'"
    End Select
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then

        If Opcion = 9 Then
            SQL = RS!Codigo & "|" & RS!Anyo & "|" & RS!codmacta & "|" & Text1(11).Text & "|"
        Else
            SQL = "|||" & Text1(11).Text & "|"
        End If
        
        
        
'        Importe = TextoAimporte(txtImporte(1).Text)   ''Levara el gasto por recibo
        If Me.optDevRem(1).Value Or Me.optDevRem(2).Value Then SQL = SQL & "%"
        SQL = SQL & "|"
        If Me.optDevRem(2).Value Then SQL = SQL & TextoAimporte(txtImporte(2).Text)
        SQL = SQL & "|"
        
        
        'SQL llevara hasta ahora
        '        remes    cta ban  fec contb tipo gasto el 1: si tiene valor es el minimo por recibo
        ' Ej:    1|2009|572000005|20/11/2009|%|1|
        
        
        'Si contabilizamos el gasto, o pro contra vendra como factura bancaria desde otro lugar(norma34 p.e.)
        If GastoDevolGral = 0 Then
            'NO HAY GASTO
            SQL = SQL & "0|"
        Else
            SQL = SQL & CStr(GastoDevolGral) & "|"
            If ComprobarCero(txtImporte(5).Text) <> 0 Then
                'Voy a contabi�izar los gastos.
                'Vere si tiene CC
                If vParam.autocoste Then
                    If DevuelveDesdeBD("codccost", "bancos", "codmacta", RS!codmacta, "T") = "" Then
                        MsgBox "Va a contabilizar los gastos pero no esta configurado el Centro de coste para el banco: " & RS!codmacta, vbExclamation
                        RS.Close
                        Set RS = Nothing
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        'Depues del gasto
        'A�adire el fichero, si es autmatico
        SQL = SQL & "|"
        'Nov 2012. En las devoluciones puede ser que el fichero traiga mas de una devolucion
        SQL = SQL & "|"
        
        vRemesa = SQL
    End If


End Sub


Private Sub CargaList()
Dim Itm As ListItem
Dim Col As Collection
Dim EfectoSerie As String
Dim EfectoFra As String
Dim EfectoVto As String
Dim EltoItm  As ListItem
Dim EsSepa As Boolean
Dim cad As String
Dim jj As Long

Dim TipoFicheroDevolucion As Byte

    CargaCabecera
    
    Set lwCobros.SmallIcons = frmPpal.imgListComun16
    
    lwCobros.ListItems.Clear
    
    vSql = ""
    Select Case Opcion
        Case 9
            vSql = " AND codrem =" & DBSet(Text3(5).Text, "N")
            vSql = vSql & " AND anyorem =" & DBSet(Text3(6).Text, "N")
        
    End Select
    
    
    If Opcion <> 16 Then
        vSql = "Select cobros.* from cobros where (1=1)" & vSql
        
        vSql = vSql & " ORDER BY numserie,numfactu"
        Set miRsAux = New ADODB.Recordset
        lwCobros.ListItems.Clear
        miRsAux.Open vSql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        jj = 1
        While Not miRsAux.EOF
            Set Itm = lwCobros.ListItems.Add(, "C" & jj)
            Itm.Text = miRsAux!NUmSerie
            
            Itm.SubItems(1) = Format(DBLet(miRsAux!NumFactu, "N"), "0000000")
            Itm.SubItems(2) = miRsAux!numorden
            Itm.SubItems(3) = miRsAux!codmacta
            Itm.SubItems(4) = miRsAux!nomclien
            ImporteQueda = DBLet(miRsAux!Gastos, "N")
            'No lo pongo con el importe de gastos pq pudiera ser k habiendo sido devuelto, no quiera
            ' cobrarle gastos
            If DBLet(miRsAux!Devuelto, "N") = 1 Then
                Itm.SmallIcon = 42
            End If
            ImporteQueda = ImporteQueda + miRsAux!ImpVenci
            Itm.SubItems(5) = Format(ImporteQueda, FormatoImporte)
            
            'Para la ordenacion
            'Por si ordena por fecha
            'ItmX.SubItems(6) = Format(RS!fecfaccl, "yyyymmdd")
            'Por si ordena por importe
            Itm.SubItems(7) = Format(miRsAux!ImpVenci * 100, "0000000000")
            
                    
            'remesas
            Itm.SubItems(8) = miRsAux!CodRem
            Itm.SubItems(9) = miRsAux!AnyoRem
            Itm.SubItems(10) = DevuelveValor("select codmacta from remesas where codigo = " & DBSet(miRsAux!CodRem, "N") & " and anyo = " & DBSet(miRsAux!AnyoRem, "N"))
            
            
            'En el tag meto la fecha factura
            Itm.Tag = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        
            
            jj = jj + 1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
    
        Me.Refresh
        Screen.MousePointer = vbDefault
    
    End If
    
    ' si viene de fichero no dejamos marcar ni desmarcar
    lwCobros.Enabled = (Opcion <> 16)
    imgCheck(0).Enabled = (Opcion <> 16)
    imgCheck(1).Enabled = (Opcion <> 16)
    CmbDevol.Enabled = (Opcion <> 16)
    
End Sub


Private Sub RecorreBuscandoRecibo(ByRef Recibos As Collection, EsMensajeNoEncontrados As Boolean, EsSepa As Boolean)
    If EsSepa Then
        RecorreBuscandoReciboSEPA Recibos, EsMensajeNoEncontrados
    Else
        RecorreBuscandoRecibo2 Recibos, EsMensajeNoEncontrados
    End If
End Sub



Private Sub PonerVtosRemesa(vSql As String, Modificar As Boolean)
Dim IT
Dim ImporteTot As Currency
Dim cad As String
Dim Importe As Currency


    lwCobros.ListItems.Clear
    If Not Modificar Then Text1(4).Text = ""
    
    ImporteTot = 0
    
'    Set Me.lwCobros.SmallIcons = frmPpal.ImgListviews
    Set lwCobros.SmallIcons = frmPpal.imgListComun16
    Set miRsAux = New ADODB.Recordset
    
    cad = "Select cobros.*,nomforpa " & vSql
    cad = cad & " ORDER BY fecvenci"
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCobros.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!NumFactu, "0000000")
        IT.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!numorden
        IT.SubItems(4) = miRsAux!FecVenci
        IT.SubItems(5) = miRsAux!nomforpa
    
        IT.Checked = False
    
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!ImpVenci
        
        'Si ya he cobrado algo
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        
        IT.SubItems(6) = Format(Importe, FormatoImporte)
        
        ImporteTot = ImporteTot + Importe

        IT.Tag = Abs(Importe)  'siempre valor absoluto
            
        If DBLet(miRsAux!Devuelto, "N") = 1 Then
            IT.SmallIcon = 42
        End If
            
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    Text1(4).Text = Format(ImporteTot, "###,###,##0.00")
    

End Sub


Private Sub txtConcepto_GotFocus(Index As Integer)
    ObtenerFoco txtConcepto(Index)
End Sub

Private Sub txtConcepto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtConcepto_LostFocus(Index As Integer)
Dim SQL As String

    'Lost focus
    txtConcepto(Index).Text = Trim(txtConcepto(Index).Text)
    SQL = ""
    I = 0
    If txtConcepto(Index).Text <> "" Then
        If Not IsNumeric(txtConcepto(Index).Text) Then
            MsgBox "Campo num�rico", vbExclamation
            I = 1
        Else
            
            SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtConcepto(Index).Text, "N")
            If SQL = "" Then
                MsgBox "Concepto no existe", vbExclamation
                I = 1
            End If
        End If
    End If
    Me.txtDConcpeto(Index).Text = SQL
    If I = 1 Then
        txtConcepto(Index).Text = ""
        PonerFoco txtConcepto(Index)
    Else
        SQL = "select ampdecli from tipofpago where tipoformapago = 4"
        I = DevuelveValor(SQL)
        PosicionarCombo Me.Combo2(0), I
    End If
End Sub



Private Sub txtImporte_GotFocus(Index As Integer)
    With txtImporte(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
 Dim Valor
        txtImporte(Index).Text = Trim(txtImporte(Index))
        If txtImporte(Index).Text = "" Then Exit Sub
        

        If Not EsNumerico(txtImporte(Index).Text) Then
            txtImporte(Index).Text = ""
            Exit Sub
        End If
    
        
        If Index = 6 Or Index = 7 Then
           
            If InStr(1, txtImporte(Index).Text, ",") > 0 Then
                Valor = ImporteFormateado(txtImporte(Index).Text)
            Else
                Valor = CCur(TransformaPuntosComas(txtImporte(Index).Text))
            End If
            txtImporte(Index).Text = Format(Valor, FormatoImporte)
        End If
        
End Sub





Private Sub CamposRemesaAbono()
       
   
   
   If Text3(3) <> "" And Text3(4).Text <> "" Then
        
        Set RS = New ADODB.Recordset
        SQL = "select importe,nommacta from remesas,cuentas where remesas.codmacta=cuentas.codmacta "
        SQL = SQL & " and anyo=" & Text3(4).Text & " and codigo=" & Text3(3).Text
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
        End If
        RS.Close
        Set RS = Nothing
    End If
    
End Sub



Private Sub EliminarEnRecepcionDocumentos()
Dim CtaPte As Boolean
Dim J As Integer
Dim CualesEliminar As String
On Error GoTo EEliminarEnRecepcionDocumentos

    'Comprobaremos si hay datos
    
        'Si no lleva cuenta puente, no hace falta que este contabilizada
        'Es decir. Solo mirare contabilizados si llevo ctapuente
        CuentasCC = ""
        CualesEliminar = ""
        J = 0
        For I = 0 To 1
            ' contatalonpte
            SQL = "pagarecta"
            If I = 1 Then SQL = "contatalonpte"
            CtaPte = (DevuelveDesdeBD(SQL, "paramtesor", "codigo", "1") = "1")
            
            'Repetiremos el proceso dos veces
            SQL = "Select * from scarecepdoc where fechavto<='" & Format(Text1(17).Text, FormatoFecha) & "'"
            SQL = SQL & " AND   talon = " & I
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                    'Si lleva cta puente habra que ver si esta contbilizada
                    J = 0
                    If CtaPte Then
                        If Val(RS!Contabilizada) = 0 Then
                            'Veo si tiene lineas. S
                            SQL = DevuelveDesdeBD("count(*)", "slirecepdoc", "id", CStr(RS!Codigo))
                            If SQL = "" Then SQL = "0"
                            If Val(SQL) > 0 Then
                                CuentasCC = CuentasCC & RS!Codigo & " - No contabilizada" & vbCrLf
                                J = 1
                            End If
                        End If
                    End If
                    If J = 0 Then
                        'Si va benee
                        If Val(DBLet(RS!llevadobanco, "N")) = 0 Then
                            SQL = DevuelveDesdeBD("count(*)", "slirecepdoc", "id", CStr(RS!Codigo))
                            If SQL = "" Then SQL = "0"
                            If Val(SQL) > 0 Then
                                CuentasCC = CuentasCC & RS!Codigo & " - Sin llevar a banco" & vbCrLf
                                J = 1
                            End If
                    
                        End If
                    End If
                    'Esta la borraremos
                    If J = 0 Then CualesEliminar = CualesEliminar & ", " & RS!Codigo
                    
                    RS.MoveNext
            Wend
            RS.Close
            
            
            
        Next I
        
        

        
        If CualesEliminar = "" Then
            'No borraremos ninguna
            If CuentasCC <> "" Then
                CuentasCC = "No se puede eliminar de la recepcion de documentos los siguientes registros: " & vbCrLf & vbCrLf & CuentasCC
                MsgBox CuentasCC, vbExclamation
                
            End If
            Exit Sub
        End If
            
        
        
        'Si k hay para borrar
        CualesEliminar = Mid(CualesEliminar, 2)
        J = 1
        SQL = "X"
        Do
            I = InStr(J, CualesEliminar, ",")
            If I > 0 Then
                J = I + 1
                SQL = SQL & "X"
            End If
        Loop Until I = 0
        
        SQL = "Va a eliminar " & Len(SQL) & " registros de la recepcion de documentos." & vbCrLf & vbCrLf & vbCrLf
        If CuentasCC <> "" Then CuentasCC = "No se puede eliminar de la recepcion de documentos los siguientes registros: " & vbCrLf & vbCrLf & CuentasCC
        SQL = SQL & vbCrLf & CuentasCC
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
            SQL = "DELETE from slirecepdoc where id in (" & CualesEliminar & ")"
            Conn.Execute SQL
            
            SQL = "DELETE from scarecepdoc where codigo in (" & CualesEliminar & ")"
            Conn.Execute SQL
    
        End If

    Exit Sub
EEliminarEnRecepcionDocumentos:
    MuestraError Err.Number, Err.Description
End Sub

'Esta recibo SEPA
Private Sub RecorreBuscandoReciboSEPA(ByRef Recibos As Collection, EsMensajeNoEncontrados As Boolean)
Dim B As Boolean
Dim cad As String
Dim jj As Integer


    If EsMensajeNoEncontrados Then
            cad = ""
            For jj = Recibos.Count To 1 Step -1
                'M  0330047820131201001   430000061
                'SER FACTU    FEC   VTO
                
                'ImporteQueda = CCur(Val(Mid(Recibos(jj), 1, 10)) / 100)
                cad = cad & jj & ".-Fecha: "
                cad = cad & Mid(Recibos(jj), 18, 2) & "/" & Mid(Recibos(jj), 16, 2) & "/" & Mid(Recibos(jj), 12, 4)
                cad = cad & " Vto: " & Mid(Recibos(jj), 1, 3) & "/" & Mid(Recibos(jj), 4, 8) & "-" & Mid(Recibos(jj), 20, 3) & vbCrLf
            Next jj
            cad = "Recibos no encontrados que vienen del fichero." & vbCrLf & vbCrLf & cad
            MsgBox cad, vbExclamation
            ImporteQueda = 0
    Else
        
        For jj = Recibos.Count To 1 Step -1
            'M  0330047820131201001   430000061
            'SER FACTU    FEC   VTO
            cad = Mid(Recibos(jj), 18, 2) & "/" & Mid(Recibos(jj), 16, 2) & "/" & Mid(Recibos(jj), 12, 4)
            
            
            B = EstaElReciboSEPA(Trim(Mid(Recibos(jj), 1, 3)), Mid(Recibos(jj), 4, 8), cad, Mid(Recibos(jj), 20, 3))

            If B Then Recibos.Remove jj
        Next jj
                
    End If
    
End Sub



Private Sub RecorreBuscandoRecibo2(ByRef Recibos As Collection, EsMensajeNoEncontrados As Boolean)
Dim B As Boolean

Dim EsFormatoAntiguoDevolucion As Boolean
Dim cad As String
Dim jj As Long

    'Formato antiguo:A020500021
    'En el nuevo : X 00045771 >> serie(2)=X  factura(7)=4577    vto(1)=1
    EsFormatoAntiguoDevolucion = Dir(App.Path & "\DevRecAnt.dat") <> ""
    

    If EsMensajeNoEncontrados Then
            cad = ""
            For jj = Recibos.Count To 1 Step -1
                'Ejemplo 0047080000004708
                '        251205A020500021
                '        $$$$$$ fecha                       6
                '              $ Serie                      1
                '               $$$$$$$$  Facutra           8
                '                       $  Vencimiento      1
                'La fecha
                ImporteQueda = CCur(Val(Mid(Recibos(jj), 1, 10)) / 100)
                cad = cad & jj & ".-Fecha: "
                cad = cad & Mid(Recibos(jj), 11, 2) & "/" & Mid(Recibos(jj), 13, 2) & "/20" & Mid(Recibos(jj), 15, 2)
                cad = cad & " Vto: " & Mid(Recibos(jj), 17, 1) & "/" & Mid(Recibos(jj), 18, 8) & "-" & Mid(Recibos(jj), 26, 1) & "   Importe: " & Format(ImporteQueda, FormatoImporte) & vbCrLf
            Next jj
            cad = "Recibos no encontrados que vienen del fichero." & vbCrLf & vbCrLf & cad
            MsgBox cad, vbExclamation
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
            cad = Mid(Recibos(jj), 11, 2) & "/" & Mid(Recibos(jj), 13, 2) & "/20" & Mid(Recibos(jj), 15, 2)
            'Octubre 2011
            'If Not IsNumeric(Mid(Recibos(jj), 27, 1)) Then
               
            'SEPT 2013
            If Not EsFormatoAntiguoDevolucion Then
                'Alzira. Estaba mal formateado el numfac.
               B = EstaElRecibo(Mid(Recibos(jj), 17, 2), Mid(Recibos(jj), 19, 7), cad, Mid(Recibos(jj), 26, 1))
            Else
               B = EstaElRecibo(Mid(Recibos(jj), 17, 2), Mid(Recibos(jj), 20, 7), cad, Mid(Recibos(jj), 27, 1))
            End If
            If B Then Recibos.Remove jj
        Next jj
                
    End If
    
End Sub


Private Function EstaElReciboSEPA(Serie As String, Fac As String, Fec As String, Venci As String) As Boolean
Dim J As Integer

    EstaElReciboSEPA = False
    With lwCobros
        For J = 1 To .ListItems.Count
            If Trim(.ListItems(J).Text) = Trim(Serie) Then
                'Misma serie
                If Val(.ListItems(J).SubItems(1)) = Val(Fac) And Val(.ListItems(J).SubItems(2)) = Venci And .ListItems(J).Tag = Fec Then
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


Private Function EstaElRecibo(Serie As String, Fac As String, Fec As String, Venci As String) As Boolean
Dim J As Integer

    EstaElRecibo = False
    With lwCobros
        For J = 1 To .ListItems.Count
            If Mid(.ListItems(J).Text, 1, 2) = Trim(Serie) Then
                'Misma serie
                If Val(.ListItems(J).SubItems(1)) = Val(Fac) And .ListItems(J).SubItems(2) = Venci And .ListItems(J).Tag = Fec Then
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
                        If Val(.ListItems(J).SubItems(1)) = Val(Fac & Venci) And .ListItems(J).Tag = Fec Then
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






