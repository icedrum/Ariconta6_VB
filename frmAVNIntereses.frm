VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAVNIntereses 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cálculo de intereses"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7650
   Icon            =   "frmAVNIntereses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   7650
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
      Height          =   8850
      Left            =   0
      TabIndex        =   10
      Top             =   45
      Width           =   7410
      Begin VB.CommandButton cmdAccion 
         Caption         =   "&Imprimir"
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
         Left            =   225
         TabIndex        =   8
         Top             =   8190
         Width           =   1335
      End
      Begin VB.Frame FrameTipoSalida 
         Caption         =   "Tipo de salida"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   225
         TabIndex        =   21
         Top             =   5355
         Width           =   6915
         Begin VB.OptionButton optTipoSal 
            Caption         =   "Impresora"
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
            TabIndex        =   31
            Top             =   720
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optTipoSal 
            Caption         =   "Archivo csv"
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
            Left            =   240
            TabIndex        =   30
            Top             =   1200
            Width           =   1515
         End
         Begin VB.OptionButton optTipoSal 
            Caption         =   "PDF"
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
            Left            =   240
            TabIndex        =   29
            Top             =   1680
            Width           =   975
         End
         Begin VB.OptionButton optTipoSal 
            Caption         =   "eMail"
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
            Left            =   240
            TabIndex        =   28
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtTipoSalida 
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
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   720
            Width           =   3345
         End
         Begin VB.TextBox txtTipoSalida 
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
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   1200
            Width           =   4665
         End
         Begin VB.TextBox txtTipoSalida 
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
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   1680
            Width           =   4665
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   0
            Left            =   6450
            TabIndex        =   24
            Top             =   1200
            Width           =   255
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   1
            Left            =   6450
            TabIndex        =   23
            Top             =   1680
            Width           =   255
         End
         Begin VB.CommandButton PushButtonImpr 
            Caption         =   "Propiedades"
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
            Left            =   5190
            TabIndex        =   22
            Top             =   720
            Width           =   1515
         End
      End
      Begin VB.Frame FrameConcepto 
         Caption         =   "Selección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4965
         Left            =   225
         TabIndex        =   11
         Top             =   270
         Width           =   6915
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
            Index           =   6
            Left            =   1305
            MaxLength       =   6
            TabIndex        =   6
            Top             =   4410
            Width           =   830
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
            Left            =   1305
            MaxLength       =   30
            TabIndex        =   5
            Top             =   3645
            Width           =   4410
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
            Left            =   1290
            MaxLength       =   10
            TabIndex        =   2
            Top             =   2040
            Width           =   1305
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
            Left            =   1290
            MaxLength       =   10
            TabIndex        =   3
            Top             =   2400
            Width           =   1305
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   4
            Left            =   1290
            MaxLength       =   10
            TabIndex        =   4
            Top             =   3240
            Width           =   1305
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
            Index           =   1
            Left            =   2190
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "Text5"
            Top             =   1320
            Width           =   4170
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
            Index           =   0
            Left            =   2190
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "Text5"
            Top             =   945
            Width           =   4170
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
            Left            =   1290
            MaxLength       =   6
            TabIndex        =   1
            Top             =   1320
            Width           =   830
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
            Left            =   1290
            MaxLength       =   6
            TabIndex        =   0
            Top             =   945
            Width           =   830
         End
         Begin VB.Label Label4 
            Caption         =   "% Retención"
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
            Index           =   10
            Left            =   225
            TabIndex        =   33
            Top             =   4140
            Width           =   2940
         End
         Begin VB.Label Label4 
            Caption         =   "Concepto"
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
            Index           =   1
            Left            =   225
            TabIndex        =   32
            Top             =   3690
            Width           =   2940
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   945
            Picture         =   "frmAVNIntereses.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   3285
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   3
            Left            =   975
            Picture         =   "frmAVNIntereses.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   975
            Picture         =   "frmAVNIntereses.frx":0122
            ToolTipText     =   "Buscar fecha"
            Top             =   2040
            Width           =   240
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
            Left            =   225
            TabIndex        =   20
            Top             =   2400
            Width           =   600
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
            Left            =   225
            TabIndex        =   19
            Top             =   2040
            Width           =   600
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Liquidación"
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
            Left            =   225
            TabIndex        =   18
            Top             =   1710
            Width           =   2760
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha de Movimiento"
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
            Index           =   0
            Left            =   225
            TabIndex        =   17
            Top             =   2955
            Width           =   2940
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   990
            MouseIcon       =   "frmAVNIntereses.frx":01AD
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cliente"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1005
            MouseIcon       =   "frmAVNIntereses.frx":02FF
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cliente"
            Top             =   945
            Width           =   240
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
            Left            =   225
            TabIndex        =   16
            Top             =   1320
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
            Index           =   13
            Left            =   225
            TabIndex        =   15
            Top             =   945
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Código Avnics"
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
            Height          =   375
            Index           =   8
            Left            =   180
            TabIndex        =   12
            Top             =   540
            Width           =   3120
         End
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
         Left            =   5805
         TabIndex        =   9
         Top             =   8235
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
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
         Index           =   1
         Left            =   4320
         TabIndex        =   7
         Top             =   8235
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAVNIntereses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Integer

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmavn As frmAVNAvnics 'Avnics
Attribute frmavn.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub





Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String

Dim cDesde As String
Dim cHasta As String
Dim cadWhere As String
Dim Cad As String
Dim NReg As Long


    MontaSQL = False
    
    If Not DatosOk Then Exit Function
    
    
    If Not PonerDesdeHasta("avnic.fechalta", "F", Me.txtCodigo(2), Me.txtCodigo(2), Me.txtCodigo(3), Me.txtCodigo(3), "pDHfechaVto=""") Then Exit Function
    If Not PonerDesdeHasta("avnic.codavnic", "COD", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHavnics=""") Then Exit Function
    
    cadParam = cadParam & "pFechaMov= """ & txtCodigo(4).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pReten= " & DBSet(txtCodigo(6).Text, "N") & "|"
    numParam = numParam + 1
    
    
    AnyadirAFormula cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo
    
    ' el avnic no tiene que haber sido cancelado en el ejercicio
    cadWhere = "where anoejerc = year(" & DBSet(txtCodigo(4).Text, "F") & ") and codialta <> 2 "
    
    If txtCodigo(0).Text <> "" Then cadWhere = cadWhere & " and avnic.codavnic >= " & DBSet(txtCodigo(0).Text, "N")
    If txtCodigo(1).Text <> "" Then cadWhere = cadWhere & " and avnic.codavnic <= " & DBSet(txtCodigo(1).Text, "N")
    cadWhere = cadWhere & " and (( (1 = 1) "
    If txtCodigo(2).Text <> "" Then cadWhere = cadWhere & " and avnic.fechalta >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then cadWhere = cadWhere & " and avnic.fechalta <= " & DBSet(txtCodigo(3).Text, "F")
    cadWhere = cadWhere & " ) or ( (1 = 1) "
    If txtCodigo(2).Text <> "" Then cadWhere = cadWhere & " and avnic.fechalta <= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(2).Text <> "" Then cadWhere = cadWhere & " and avnic.fechavto >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then cadWhere = cadWhere & " and avnic.fechavto <= " & DBSet(txtCodigo(3).Text, "F")
    cadWhere = cadWhere & " ) or ( (1 = 1) "
    If txtCodigo(2).Text <> "" Then cadWhere = cadWhere & " and avnic.fechalta <= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then cadWhere = cadWhere & " and avnic.fechavto > " & DBSet(txtCodigo(3).Text, "F")
    cadWhere = cadWhere & "))"
    
  
    Cad = "select count(*) from avnic  " & cadWhere
    NReg = TotalRegistros(Cad)
    If NReg <> 0 Then
       If CargarTablaIntermedia(cadWhere) Then
            MontaSQL = True
       End If
       
    End If
    
End Function


Private Function DatosOk() As Boolean
Dim Sql As String
Dim b As Boolean

    b = True
    If txtCodigo(4).Text = "" Then
        MsgBoxA "La fecha de movimiento debe de tener un valor. Reintroduzca.", vbExclamation
        b = False
        PonFoco txtCodigo(4)
    End If
    DatosOk = b
    
End Function


Private Sub cmdAccion_Click(Index As Integer)
Dim Cad As String

    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("tmpinformes", "codusu = " & DBSet(vUsu.Codigo, "N")) Then Exit Sub
    
    If optTipoSal(1).Value Then
        'EXPORTAR A CSV
        AccionesCSV
    
    Else
        'Tanto a pdf,imprimiir, preevisualizar como email van COntral Crystal
    
        If optTipoSal(2).Value Or optTipoSal(3).Value Then
            ExportarPDF = True 'generaremos el pdf
        Else
            ExportarPDF = False
        End If
        SoloImprimir = False
        If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
        
        AccionesCrystal
    End If
    
    Cad = "¿ Impresión correcta para actualizar ?"
    If MsgBoxA(Cad, vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
        If ActualizarTablas Then
            MsgBoxA "Proceso realizado correctamente", vbInformation
            cmdCancel_Click
        End If
    End If

End Sub

Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    vMostrarTree = False
    conSubRPT = False
        
'    cadParam = cadParam & "pTitulo=""" & txtTitulo(0).Text & """|"
'    cadParam = cadParam & "pFecha=""" & txtFecha(2).Text & """|"
'    cadParam = cadParam & "pTotalAsiento=" & chkTotalAsiento.Value & "|"
'
'    numParam = numParam + 3
    
    cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
    
    indRPT = "1416-01" '"rCalculInt.rpt"

    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Sub AccionesCSV()
Dim Sql As String

    'Monto el SQL
    Sql = "Select codavnic AS Código,avnic.nombrper as ApeNombre, importe2 as Importe, fecha1 as  deFecha,fecha2 as aFecha, datediff(fecha2,fecha1) as Dias,importe1 as Intereses, round(importe1*" & DBSet(txtCodigo(6), "N") & "*0.01,2) as Retencion, importe1 - round(importe1*" & DBSet(txtCodigo(6), "N") & "*0.01,2) as ImporteNeto "
    Sql = Sql & " From tmpinformes inner join avnic on tmpinformes.codigo1 = avnic.codavnic and avnic.anoejerc=tmpinformes.campo1 "
    Sql = Sql & " WHERE tmpinformes.codusu = " & DBSet(vUsu.Codigo, "N")
    
'    If cadselect <> "" Then SQL = SQL & " AND " & cadselect
    
    Sql = Sql & " ORDER BY 1 "
        
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonFoco txtCodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    Limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmppal.imgIcoForms.ListImages(1).Picture
     
     
    FrameCobros.visible = False
     
    '###Descomentar
'    CommitConexion
    Select Case OpcionListado
        Case 0 ' calculo de intereses
            FrameCobrosVisible True, h, w
            indFrame = 5
            Tabla = "avnic"
                    
            txtCodigo(4).Text = Format(Now, "dd/mm/yyyy")
'            txtCodigo(6).Text = Format(vParamAplic.Porcrete, "##0.00")
            
            PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
            ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
        
    End Select
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmAvn_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
    Dim obj As Object

    indCodigo = Index
    Select Case Index
        Case 0
            indCodigo = 6
    End Select
    
    'FECHA
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtCodigo(indCodigo).Text <> "" Then frmC.Fecha = CDate(txtCodigo(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    PonFoco txtCodigo(indCodigo)
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'AVNICS
            AbrirFrmAvnics (Index)
        
    End Select
    PonFoco txtCodigo(indCodigo)
End Sub

Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub

Private Sub PushButton2_Click(Index As Integer)
    'FILTROS
    If Index = 0 Then
         frmppal.cd1.Filter = "*.csv|*.csv"
         
    Else
        frmppal.cd1.Filter = "*.pdf|*.pdf"
    End If
    frmppal.cd1.InitDir = App.Path & "\Exportar" 'PathSalida
    frmppal.cd1.FilterIndex = 1
    frmppal.cd1.ShowSave
    If frmppal.cd1.FileTitle <> "" Then
        If Dir(frmppal.cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        txtTipoSalida(Index + 1).Text = frmppal.cd1.FileName
    End If

End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'avnics desde
            Case 1: KEYBusqueda KeyAscii, 1 'avnics hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            
            Case 4: KEYFecha KeyAscii, 0 'fecha de movimiento
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 0, 1 'AVNICS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "avnic", "nombrper", "codavnic", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
        
        Case 2, 3, 4 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 6 'RETENCION
              PonerFormatoDecimal txtCodigo(Index), 4
              
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para los cobros a clientes por fecha vencimiento
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.top = -90
        Me.FrameCobros.Left = 0
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Sub AbrirFrmAvnics(indice As Integer)
    indCodigo = indice
    Set frmavn = New frmAVNAvnics
    frmavn.DatosADevolverBusqueda = "0|4|"
    frmavn.DeConsulta = True
    frmavn.CodigoActual = txtCodigo(indCodigo)
    frmavn.Show vbModal
    Set frmavn = Nothing
End Sub
 
Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub


Public Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    cadFormula = ""
    cadSelect = ""
    cadParam = "|"
    numParam = 0
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False
    
    If AñadireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub


Private Function CargarTablaIntermedia(cadWhere As String) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Dim AntCodavnic As Long
Dim ActCodavnic As Long
Dim v_fecha As String
Dim DesFec As String
Dim HasFec As String
Dim DifDia As Long
Dim Pasado As Boolean
Dim Cad As String
Dim Importe As Currency
Dim Sql2 As String
Dim Sql3 As String

    On Error GoTo eCargarTablaIntermedia

    CargarTablaIntermedia = False
    
    If Not BorrarTablaIntermedia Then Exit Function

    Set RS = New ADODB.Recordset
    
    Sql = "select codavnic, importes, porcinte, anoejerc, fechalta, fechavto from avnic " & cadWhere
    Sql = Sql & " order by codavnic"
    
    RS.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    AntCodavnic = DBLet(RS!codavnic, "N")
    ActCodavnic = AntCodavnic
    Pasado = False
    While Not RS.EOF
        AntCodavnic = ActCodavnic
        ActCodavnic = RS!codavnic
        
        If ActCodavnic <> AntCodavnic Then Pasado = False
        
        If Not Pasado Then
                ' obtenemos la maxima fehca de movimiento
                Set Rs3 = New ADODB.Recordset
                
                Cad = "select max(fechamov) from movim where codavnic = " & DBSet(RS!codavnic, "N")
                Rs3.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                v_fecha = ""
                If Not Rs3.EOF Then v_fecha = DBLet(Rs3.Fields(0).Value, "F")
                Set Rs3 = Nothing
                
                If v_fecha = "" Then DesFec = DBLet(RS!fechalta, "F")
                If v_fecha <> "" Then DesFec = v_fecha
    
                HasFec = DBLet(RS!fechavto, "F")
                If txtCodigo(3).Text <> "" Then
                    If DBLet(RS!fechavto, "F") < CDate(txtCodigo(3).Text) Then
                        HasFec = DBLet(RS!fechavto, "F")
                    Else
                        HasFec = txtCodigo(3).Text
                    End If
                End If
                DifDia = CDate(HasFec) - CDate(DesFec)
                If DifDia < 0 Then DifDia = 0
                
                If DBLet(RS!Importes, "N") <> 0 Then
                    Importe = Round2(DBLet(RS!Importes, "N") * DBLet(RS!Porcinte, "N") * 0.01 * DifDia / 365, 2)
                    Sql2 = "insert into tmpinformes (codusu, codigo1, campo1, fecha1, fecha2, importe1, importe2, nombre1) values (" & DBSet(vUsu.Codigo, "N") & ","
                    Sql2 = Sql2 & DBSet(RS!codavnic, "N") & "," & DBSet(RS!anoejerc, "N") & "," & DBSet(DesFec, "F") & "," & DBSet(HasFec, "F") & ","
                    Sql2 = Sql2 & DBSet(Importe, "N") & ","
                    
                    Set Rs2 = New ADODB.Recordset
                    Sql3 = "select nombrper, importes from avnic where codavnic = " & DBSet(RS!codavnic, "N") & " and anoejerc = " & Year(CDate(txtCodigo(4).Text))
                    Rs2.Open Sql3, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                    
                    If Rs2.EOF Then
                        Sql2 = Sql2 & "0, null)"
                    Else
                        Sql2 = Sql2 & DBSet(Rs2!Importes, "N") & "," & DBSet(Rs2!nombrper, "T") & ")"
                    End If
                
                    Conn.Execute Sql2
                End If
                
                Pasado = True
        End If
        
        RS.MoveNext
    Wend
    CargarTablaIntermedia = True
    Exit Function
    
eCargarTablaIntermedia:
    MuestraError Err.Number, "Error cargando la tabla intermedia. Llame a soporte.", Err.Description
End Function

Private Function BorrarTablaIntermedia() As Boolean
Dim Sql As String
    
    On Error GoTo eBorrarTablaIntermedia
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    BorrarTablaIntermedia = True
    Exit Function

eBorrarTablaIntermedia:
    BorrarTablaIntermedia = False

End Function

Private Function ActualizarTablas() As Boolean
Dim Sql As String
Dim Rs3 As ADODB.Recordset
Dim Cad As String
Dim v_import As Currency
Dim t_import As Currency

    On Error GoTo eActualizarTablas

    ActualizarTablas = False

    Conn.BeginTrans


    Set Rs3 = New ADODB.Recordset
    Cad = "select codigo1, campo1, importe1 from tmpinformes where codusu = " & vUsu.Codigo
    Rs3.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs3.EOF
        v_import = 0
        t_import = 0
        v_import = Round2(Rs3!Importe1 * CCur(ComprobarCero(txtCodigo(6).Text)) * 0.01, 2)
        t_import = Rs3!Importe1 - v_import
        
        If t_import <> 0 Then
            Sql = ""
            Sql = DevuelveDesdeBDNew(cConta, "movim", "codavnic", "codavnic", Rs3!codigo1, "N", , "fechamov", txtCodigo(4), "F", "anoejerc", Year(CDate(txtCodigo(4).Text)), "N")
            If Sql = "" Then
                Sql = "insert into movim (codavnic, fechamov, concepto, timporte, intconta, anoejerc, timport1, timport2) "
                Sql = Sql & "values (" & DBSet(Rs3!codigo1, "N") & "," & DBSet(txtCodigo(4).Text, "F") & "," & DBSet(txtCodigo(5).Text, "T") & ","
                Sql = Sql & DBSet(t_import, "N") & ",0," & DBSet(Year(CDate(txtCodigo(4).Text)), "N") & ","
                Sql = Sql & DBSet(Rs3!Importe1, "N") & "," & DBSet(v_import, "N") & ")"
                
                Conn.Execute Sql
            End If
            Sql = "update avnic set imporper = imporper + " & DBSet(Rs3!Importe1, "N") & ","
            Sql = Sql & " imporret = imporret + " & DBSet(v_import, "N")
            Sql = Sql & " where codavnic = " & DBSet(Rs3!codigo1, "N") & " and anoejerc = " & DBSet(Year(CDate(txtCodigo(4).Text)), "N")
            
            Conn.Execute Sql
                
        End If
    
        Rs3.MoveNext
    Wend
    Set Rs3 = Nothing
    
    Conn.CommitTrans
    ActualizarTablas = True
    Exit Function
    
eActualizarTablas:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la actualizacion de datos: " & Err.Description
        Conn.RollbackTrans
    End If
End Function


