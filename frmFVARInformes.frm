VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFVARInformes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Conceptos"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   11145
   Icon            =   "frmFVARInformes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameErrores 
      Height          =   5505
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   8835
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
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
         Left            =   7335
         TabIndex        =   24
         Top             =   4830
         Width           =   1065
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4155
         Left            =   210
         TabIndex        =   25
         Top             =   495
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin VB.Label Label5 
         Caption         =   "Errores de Comprobación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   225
         TabIndex        =   27
         Top             =   180
         Width           =   3585
      End
      Begin VB.Label Label1 
         Caption         =   "Label2"
         Height          =   345
         Index           =   7
         Left            =   450
         TabIndex        =   26
         Top             =   1470
         Width           =   3555
      End
   End
   Begin VB.Frame FrameCuentasContables 
      Height          =   5790
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   10020
      Begin VB.CommandButton cmdAcepCtasContables 
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
         Left            =   7095
         TabIndex        =   34
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdCanCtasContables 
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
         Left            =   8535
         TabIndex        =   33
         Top             =   5160
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   4245
         Left            =   225
         TabIndex        =   32
         Top             =   675
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7488
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
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
      Begin VB.Label Label3 
         Caption         =   "Cuentas Contables"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   375
         Index           =   0
         Left            =   270
         TabIndex        =   35
         Top             =   270
         Width           =   5145
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   600
         Picture         =   "frmFVARInformes.frx":000C
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   240
         Picture         =   "frmFVARInformes.frx":0156
         Top             =   5160
         Width           =   240
      End
   End
   Begin VB.Frame FrameCobros 
      Height          =   5925
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   10875
      Begin VB.Frame frameConceptoDer 
         Caption         =   "Ordenación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   7245
         TabIndex        =   28
         Top             =   270
         Width           =   3285
         Begin VB.OptionButton optVarios 
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
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton optVarios 
            Caption         =   "Descripción"
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
            Index           =   1
            Left            =   1440
            TabIndex        =   29
            Top             =   840
            Width           =   1455
         End
      End
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
         TabIndex        =   3
         Top             =   5175
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
         TabIndex        =   12
         Top             =   2340
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   1680
            Width           =   4665
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   0
            Left            =   6450
            TabIndex        =   15
            Top             =   1200
            Width           =   255
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   1
            Left            =   6450
            TabIndex        =   14
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
            TabIndex        =   13
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
         Height          =   2085
         Left            =   225
         TabIndex        =   6
         Top             =   270
         Width           =   6915
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
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "Text5"
            Top             =   1320
            Width           =   4440
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
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "Text5"
            Top             =   945
            Width           =   4440
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
            Left            =   1200
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
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   0
            Top             =   945
            Width           =   830
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   900
            MouseIcon       =   "frmFVARInformes.frx":02A0
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar concepto"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   915
            MouseIcon       =   "frmFVARInformes.frx":03F2
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar concepto"
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
            TabIndex        =   11
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
            TabIndex        =   10
            Top             =   945
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Código"
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
            TabIndex        =   7
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
         Left            =   9225
         TabIndex        =   4
         Top             =   5175
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
         Left            =   7740
         TabIndex        =   2
         Top             =   5175
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6750
      Top             =   3105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmFVARInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Integer
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmFVARCon As frmFVARConceptos 'Conceptos
Attribute frmFVARCon.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

Dim IndCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
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

    MontaSQL = False
    
    If Not PonerDesdeHasta("fvarconceptos.codconce", "COD", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHavnics=""") Then Exit Function
    
    
    MontaSQL = True
    
End Function






Private Sub cmdAccion_Click(Index As Integer)

    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("fvarconceptos", cadselect) Then Exit Sub
    
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

End Sub

Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
        
    vMostrarTree = False
    conSubRPT = False
        
    If optVarios(1).Value Then
        cadParam = cadParam & "pOrden={fvarconceptos.nomconce}|"
    Else
        cadParam = cadParam & "pOrden={fvarconceptos.codconce}|"
    End If
    numParam = numParam + 1
        
        
        
    indRPT = "420-00" '"rInfAvnics.rpt"

    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 1
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Sub AccionesCSV()
Dim Sql As String

    'Monto el SQL
    Sql = "Select codconce AS Código, nomconce as Descripcion, fvarconceptos.codmacta as CtaContable, nommacta as Descripcion,  "
    Sql = Sql & " tipoiva TipoIva, porceiva PorIva, codccost CentroCoste"
    Sql = Sql & " from (fvarconceptos inner join cuentas on fvarconceptos.codmacta = cuentas.codmacta) inner join tiposiva on fvarconceptos.tipoiva = tiposiva.codigiva "
    
    If cadselect <> "" Then Sql = Sql & " WHERE " & cadselect
    
    If optVarios(1).Value Then
        Sql = Sql & " ORDER BY 2 "
    Else
        Sql = Sql & " ORDER BY 1 "
    End If
    
        
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub cmdAcepCtasContables_Click()
Dim Cadena As String
    'Cargo las variedades marcadas
    Cadena = ""
    For NumRegElim = 1 To ListView6.ListItems.Count
        If ListView6.ListItems(NumRegElim).Checked Then
             Cadena = Cadena & "'" & Trim(ListView6.ListItems(NumRegElim).Text) & "',"
        End If
    Next NumRegElim
    ' quitamos la ultima coma
    If Cadena <> "" Then
        Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
    End If
    
    RaiseEvent DatoSeleccionado(Cadena)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCanCtasContables_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 0
                optVarios(0).Value = True

                PonFoco txtCodigo(0)
            Case 1
            
        
        End Select
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
     
     
    FrameCobros.visible = False
    Me.FrameErrores.visible = False
    Me.FrameCuentasContables.visible = False
    
    '###Descomentar
'    CommitConexion
    Select Case OpcionListado
        Case 0 ' listado de conceptos
            FrameCobrosVisible True, H, W
            indFrame = 5
            tabla = "fvarconceptos"
                
            PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
            ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
            
        Case 1 ' cuentas contables
            PonerFrameCuentasVisible True, H, W
            CargarListaCuentas
        
        Case 2 ' visualizacion de errores de la contabilizacion
            PonerFrameErroresVisible True, H, W
            CargarListaErrComprobacion
            Me.Caption = "Errores de Comprobacion: "
            PonerFocoBtn Me.CmdSalir
        
    End Select
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(IndCodigo).Text = Format(vFecha, "dd/MM/yyyy")
End Sub


Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
    Dim Obj As Object

    IndCodigo = Index
    Select Case Index
        Case 0
            IndCodigo = 6
    End Select
    
    'FECHA
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtCodigo(IndCodigo).Text <> "" Then frmC.Fecha = CDate(txtCodigo(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    PonFoco txtCodigo(IndCodigo)
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub frmFVARCon_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'Facturas varias conceptos
            AbrirFrmFVARConceptos (Index)
        
    End Select
    PonFoco txtCodigo(IndCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
'        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
'        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim B As Boolean
Dim TotalArray As Integer

    Select Case Index
       Case 2, 3
            'En el listview7
            B = (Index = 2)
            For TotalArray = 1 To ListView6.ListItems.Count
                ListView6.ListItems(TotalArray).Checked = B
                If (TotalArray Mod 50) = 0 Then DoEvents
            Next TotalArray
    End Select
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

Private Sub PushButtonImpr_Click()
    frmppal.cd1.ShowPrinter
    PonerDatosPorDefectoImpresion Me, True
End Sub

Private Sub txtcodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtcodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtcodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'concepto desde
            Case 1: KEYBusqueda KeyAscii, 1 'concepto hasta
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

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0, 1 'CONCEPTOS DE FACTURAS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "fvarconceptos", "nomconce", "codconce", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
              
  End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los cobros a clientes por fecha vencimiento
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.top = -90
        Me.FrameCobros.Left = 0
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub


Private Sub AbrirFrmFVARConceptos(Indice As Integer)
    IndCodigo = Indice
    Set frmFVARCon = New frmFVARConceptos
    frmFVARCon.DatosADevolverBusqueda = "0|1|"
    frmFVARCon.DeConsulta = True
    frmFVARCon.CodigoActual = txtCodigo(IndCodigo)
    frmFVARCon.Show vbModal
    Set frmFVARCon = Nothing
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
    
    Unload Me
End Sub


Public Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    cadFormula = ""
    cadselect = ""
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

Private Sub CargarListaErrComprobacion()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarListErrComprobacion

    Sql = " SELECT  * "
    Sql = Sql & " FROM tmperrcomprob "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        'Los encabezados
        ListView2.ColumnHeaders.Clear

        ListView2.ColumnHeaders.Add , , "Error en cuentas contables", 6000
        
    
        While Not Rs.EOF
            Set ItmX = ListView2.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = Rs.Fields(0).Value
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarListErrComprobacion:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub


Private Sub CargarListaCuentas()
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarListErrComprobacion

    Sql = " SELECT  * "
    Sql = Sql & " FROM cuentas "
    Sql = Sql & " where apudirec = 'S'"
    If CadTag <> "" Then Sql = Sql & CadTag
    Sql = Sql & " order by codmacta"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        'Los encabezados
        ListView6.ColumnHeaders.Clear

        ListView6.ColumnHeaders.Add , , "Cuenta", 1800.0631 '1000
        ListView6.ColumnHeaders.Add , , "Nombre", 6500 ' 2200.2522 , 1

        While Not Rs.EOF
            Set ItmX = ListView6.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = Rs.Fields(0).Value
            ItmX.SubItems(1) = Rs.Fields(1).Value
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarListErrComprobacion:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub PonerFrameErroresVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameErrores.visible = visible
    If visible = True Then
        Me.FrameErrores.top = -90
        Me.FrameErrores.Left = 0
        W = Me.FrameErrores.Width
        H = Me.FrameErrores.Height
    End If
    
End Sub

Private Sub PonerFrameCuentasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCuentasContables.visible = visible
    If visible = True Then
        Me.FrameCuentasContables.top = -90
        Me.FrameCuentasContables.Left = 0
        W = Me.FrameCuentasContables.Width
        H = Me.FrameCuentasContables.Height
    End If
    
End Sub



