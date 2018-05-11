VERSION 5.00
Begin VB.Form frmAseguradas 
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameAseg_Bas 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtfecha 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   23
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtfecha 
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   22
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   21
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   20
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   2640
         Width           =   2715
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   2280
         Width           =   2715
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   5040
         TabIndex        =   17
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdAsegBascios 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   5040
         Width           =   975
      End
      Begin VB.Frame FrOrdenAseg1 
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   5895
         Begin VB.OptionButton optAsegBasic 
            Caption         =   "Póliza"
            Height          =   255
            Index           =   2
            Left            =   4320
            TabIndex        =   14
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optAsegBasic 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   13
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton optAsegBasic 
            Caption         =   "Cuenta"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ordenar por"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   41
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.Frame FrameASeg2 
         Height          =   855
         Left            =   1560
         TabIndex        =   8
         Top             =   3120
         Width           =   4575
         Begin VB.OptionButton optFecgaASig 
            Caption         =   "Fecha factura"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optFecgaASig 
            Caption         =   "Fecha vencimiento"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   9
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame FrameForpa 
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   4080
         Width           =   5775
         Begin VB.OptionButton optFP 
            Caption         =   "Descripción forma pago"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   7
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optFP 
            Caption         =   "Descripción tipo pago"
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   2655
         End
      End
      Begin VB.Frame FrameAsegAvisos 
         Caption         =   "Avisos"
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
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   4080
         Visible         =   0   'False
         Width           =   6015
         Begin VB.OptionButton optAsegAvisos 
            Caption         =   "Falta de pago"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optAsegAvisos 
            Caption         =   "Prorroga"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optAsegAvisos 
            Caption         =   "Siniestro"
            Height          =   255
            Index           =   2
            Left            =   4320
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "ccc"
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
         Index           =   11
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   5970
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1440
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha solicitud"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   39
         Left            =   240
         TabIndex        =   28
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4200
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   38
         Left            =   840
         TabIndex        =   26
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   39
         Left            =   840
         TabIndex        =   25
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   40
         Left            =   240
         TabIndex        =   24
         Top             =   2040
         Width           =   645
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   1440
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   0
         Left            =   1440
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   19
         Left            =   3600
         TabIndex        =   27
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   18
         Left            =   840
         TabIndex        =   29
         Top             =   1440
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmAseguradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Byte
        '1 .-       Datos basicos
        '2 .-       Listado facturacion
        '3 .-       Impagados
        '4 .-        Listado efectos
        '5 .-        ASEGURADOS.  Listados avisos falta pago, avisos prorroga, aviso siniestro
Public Parametros As String

Dim PrimeraVez  As Boolean
Dim SQL As String
Dim Cad As String

Private Sub PonerFrameVisible(ByRef Fr As frame)
    Fr.top = 30
    Fr.Left = 30
    Fr.Visible = True
    Me.Height = Fr.Height + 540
    Me.Width = Fr.Width + 120
End Sub




Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    
    PrimeraVez = True
    Limpiar Me
    Me.Icon = frmppal.Icon
    For I = 0 To imgFecha.Count - 1
        Me.imgFecha(I).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next I
    For I = 0 To imgCta.Count - 1
        Me.imgCta(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next I
    
    
    
    
    
    Select Case Opcion
    Case 0, 1, 2, 3
        'Operaciones aseguradas
        '       Datos basicos
        '       Listado facturacion
        '       Impagados
        '       Listado efectos
        '       ASEGURADOS.  Listados avisos falta pago, avisos prorroga, aviso siniestro
        optAsegBasic(2).Visible = True 'Ordenar por poliza
        FrOrdenAseg1.Visible = True
        FrameASeg2.Visible = False
        FrameForpa.Visible = False
        FrameAsegAvisos.Visible = False
        Select Case Opcion
        Case 0, 1, 2, 3, 4
            Case 0
                '       Datos basicos
                SQL = "Fecha solicitud"
                Cad = "Datos básicos operaciones aseguradas"
                
            Case 1
                '       Listado facturacion
                SQL = "Fecha"
                Cad = "List. facturacion oper. aseguradas"
                FrOrdenAseg1.Visible = False
                FrameASeg2.Visible = True
                FrameForpa.Visible = True
            Case 2
                '       Listado impagados asegurados
                SQL = "Fecha aviso"
                Cad = "Impagados en operaciones aseguradas"
                
            Case 3
                optAsegBasic(2).Visible = False
                SQL = "Fecha vto"
                Cad = "Listado efectos operaciones aseguradas"
                
            Case 4
                FrameAsegAvisos.Visible = True
               
                SQL = "Fecha aviso falta pago"
                Cad = "Listados avisos aseguradoras"
                optAsegAvisos(0).Value = True
            End Select
            
            
            Label4(39).Caption = SQL
            Label2(11).Caption = Cad
            Caption = "Listado"
            
            PonerFrameVisible FrameAseg_Bas
    
    
    
        End Select
    
    
End Sub
