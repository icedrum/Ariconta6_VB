VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESTransferDev 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "1"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15465
   Icon            =   "frmTESTransferDev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDevlucionRe 
      Height          =   8835
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
         TabIndex        =   24
         Tag             =   "Importe|N|N|||reclama|importes|||"
         Top             =   8220
         Width           =   1815
      End
      Begin VB.Frame FrameConcepto 
         Caption         =   "Datos Contabilización"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   14895
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
            ItemData        =   "frmTESTransferDev.frx":000C
            Left            =   10890
            List            =   "frmTESTransferDev.frx":0022
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Tag             =   "Ampliacion debe/CLIENTES|N|N|0||stipoformapago|ampdecli|||"
            Top             =   540
            Width           =   2820
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
            Left            =   2490
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   480
            Width           =   1275
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
            Left            =   3090
            TabIndex        =   20
            Text            =   "Text9"
            Top             =   930
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
            Left            =   2490
            TabIndex        =   4
            Text            =   "Text10"
            Top             =   930
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
            Left            =   4320
            TabIndex        =   19
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
            Left            =   10890
            TabIndex        =   5
            Text            =   "Text4"
            Top             =   990
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label7 
            Caption         =   "Ampliación"
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
            Left            =   8850
            TabIndex        =   28
            Top             =   600
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
            Left            =   8820
            TabIndex        =   27
            Top             =   1050
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   11
            Left            =   2220
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Devolución"
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
            Left            =   420
            TabIndex        =   22
            Top             =   480
            Width           =   1740
         End
         Begin VB.Image imgConcepto 
            Height          =   240
            Index           =   1
            Left            =   2220
            Top             =   990
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
            Left            =   420
            TabIndex        =   21
            Top             =   990
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
            Left            =   10080
            TabIndex        =   18
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame FrameDevDesdeRemesa 
         Height          =   1185
         Left            =   270
         TabIndex        =   13
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
            Left            =   1680
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Transferencia"
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
            TabIndex        =   16
            Top             =   210
            Width           =   1470
         End
         Begin VB.Label Label6 
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
            Index           =   5
            Left            =   240
            TabIndex        =   15
            Top             =   585
            Width           =   705
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Año"
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
            TabIndex        =   14
            Top             =   585
            Width           =   555
         End
      End
      Begin VB.Frame FrameDevDesdeVto 
         Height          =   1215
         Left            =   3990
         TabIndex        =   10
         Top             =   540
         Width           =   11085
         Begin VB.TextBox txtDCtaNormal 
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
            Index           =   11
            Left            =   3000
            TabIndex        =   11
            Text            =   "Text9"
            Top             =   570
            Width           =   6525
         End
         Begin VB.TextBox txtCtaNormal 
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
            Left            =   660
            TabIndex        =   2
            Text            =   "Text9"
            Top             =   570
            Width           =   1935
         End
         Begin VB.Image imgCtaNorma 
            Height          =   240
            Index           =   11
            Left            =   1050
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta"
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
            Height          =   195
            Index           =   37
            Left            =   180
            TabIndex        =   12
            Top             =   210
            Width           =   825
         End
      End
      Begin VB.CommandButton cmdCancelar 
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
         Height          =   375
         Index           =   9
         Left            =   13920
         TabIndex        =   7
         Top             =   8160
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
         Left            =   12480
         TabIndex        =   6
         Top             =   8160
         Width           =   1335
      End
      Begin MSComctlLib.ListView lwCobros 
         Height          =   3915
         Left            =   300
         TabIndex        =   23
         Top             =   3990
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
         TabIndex        =   25
         Top             =   8220
         Width           =   1575
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   14430
         Picture         =   "frmTESTransferDev.frx":00B0
         ToolTipText     =   "Quitar al Debe"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   14790
         Picture         =   "frmTESTransferDev.frx":01FA
         ToolTipText     =   "Puntear al Debe"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   3
         Left            =   3420
         TabIndex        =   9
         Top             =   210
         Width           =   8535
      End
   End
End
Attribute VB_Name = "frmTESTransferDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    
Public Cobros As Boolean

Public Numtrans     As Integer
Public Anotrans As Integer
    
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1
'Private WithEvents frmRe As frmTESRemesas
Private WithEvents frmB As frmBasico
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmBas2 As frmBasico2
Attribute frmBas2.VB_VarHelpID = -1

Dim Rs As ADODB.Recordset
Dim SQL As String
Dim i As Integer
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


Dim BancoRem As String



Private Sub cmdCancelar_Click(Index As Integer)
 
    Unload Me
End Sub


Private Sub cmdDevolRem_Click()
Dim Importe As Currency
Dim GastoDevolGral As Currency
Dim CadenaVencimiento As String
Dim MultiRemesaDevuelta As String
Dim TipoFicheroDevolucion As Byte

'    If Text8.Text <> "" Then Opcion = 16
'    If Text3(5).Text <> "" Then Opcion = 9
'    If txtCtaNormal(11).Text <> "" Then Opcion = 28
'
'
    SQL = ""

    If Text1(11).Text = "" Then SQL = "Ponga la fecha de abono"
    If txtConcepto(1).Text = "" Then SQL = SQL & vbCrLf & "Concepto apunte"
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    'Fecha pertenece a ejercicios contbles
    If FechaCorrecta2(CDate(Text1(11).Text), True) > 1 Then Exit Sub
    
    
   
   Importe = 0
   
     If Combo2(0).ListIndex = -1 Then
         SQL = "Ampliacion concepto incorrecta"
     End If

    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    'Nuevo Noviembre 2009
    GastoDevolGral = 0
    GastoDevolGral = TextoAimporte(txtImporte(5).Text)
    
    
    If Me.lwCobros.ListItems.Count = 0 Then Exit Sub
    
    'Ahora miramos la trasnferencia
    
    If Me.txtCtaNormal(11).Text = "" Then
        
            If Text3(5).Text = "" Or Text3(6).Text = "" Then
                MsgBox "Codigo de transferencia", vbExclamation
                PonerFoco Text3(5)
                Exit Sub
            End If
       
        
        
        
        SQL = "Select * from transferencias where codigo =" & Text3(5).Text
        SQL = SQL & " AND anyo =" & Text3(6).Text
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Rs.EOF Then
            SQL = "Ninguna remesa con esos valores."
            
            MsgBox SQL, vbExclamation
            Rs.Close
            Set Rs = Nothing
            Exit Sub
        End If
        
        
        'Tiene valor
        If Rs!Situacion = "A" Then
            MsgBox "Remesa abierta. Sin llevar al banco.", vbExclamation
            Rs.Close
            Set Rs = Nothing
            Exit Sub
        End If
        
        
        
        If Asc(Rs!Situacion) < Asc("Q") Then
            MsgBox "Remesa sin contabilizar.", vbExclamation
            Rs.Close
            Set Rs = Nothing
            Exit Sub
        End If
        SQL = Rs!Codigo & "|" & Rs!Anyo & "|" & Rs!codmacta & "|" & Text1(11).Text & "|"
    Else
    
        MultiRemesaDevuelta = ""
         For i = 1 To Me.lwCobros.ListItems.Count
            If lwCobros.ListItems(i).Checked Then
                If MultiRemesaDevuelta <> "" Then
                    If MultiRemesaDevuelta <> Me.lwCobros.ListItems(i).SubItems(11) Then
                        MsgBox "Distintos bancos de transferencias", vbExclamation
                        Exit Sub
                    End If
                Else
                    MultiRemesaDevuelta = Me.lwCobros.ListItems(i).SubItems(11)
                End If
                
            End If
        Next i
        
        SQL = "||" & MultiRemesaDevuelta & "|"
        MultiRemesaDevuelta = ""
    End If
        
    
    

    
    
    
    
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
            'Voy a contabiñizar los gastos.
            'Vere si tiene CC
            If vParam.autocoste Then
                If DevuelveDesdeBD("codccost", "bancos", "codmacta", Rs!codmacta, "T") = "" Then
                    MsgBox "Va a contabilizar los gastos pero no esta configurado el Centro de coste para el banco: " & Rs!codmacta, vbExclamation
                    Rs.Close
                    Set Rs = Nothing
                    Exit Sub
                End If
            End If
        End If
    End If
    
    
    SQL = SQL & "|"
    

    
    'Bloqueamos la devolucion
    BloqueoManual True, "Devoltrans", vUsu.Codigo
    'Hacemos la devolucion
    vRemesa = SQL
    ImporteRemesa = Importe
    
    
    SQL = txtConcepto(1).Text & "|" & Combo2(0).ListIndex & "|"
    'y el banco
    'Agrupa el apunte del banco
    SQL = SQL & Abs(chkAgrupadevol2.Value) & "|"
    
    vSql = CadenaVencimiento
    
    DevolverVtosTransferencia

    'Desbloqueamos
    BloqueoManual False, "Devoltrans", vUsu.Codigo

End Sub

Private Sub DevolverVtosTransferencia()
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
    
    
    If RecuperaValor(vRemesa, 3) = "" Then
        MsgBox "Error obteniendo banco devolucion(I)", vbExclamation
        Exit Sub
    End If
    
    
    
    'Llegado aqui hago la pregunta
    cad = "Va a realizar la devolución de " & cad & vbCrLf
    If Text1(4).Text <> "" Then
        cad = cad & vbCrLf & "Importe total de la devolución: "
        cad = cad & Text1(4).Text & "€" & vbCrLf & vbCrLf
    End If
    
    
    cad = cad & Aux & vbCrLf
    
    'Gasto tramitacion devolucion
    Aux = txtImporte(5)
    If Aux <> "" Then
        Aux = "Gasto bancario : " & Aux & "€" & vbCrLf
        cad = cad & vbCrLf & Aux
    End If
    
    cad = cad & vbCrLf & "¿Desea continuar?"
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    If Not RealizarDevolucionTransferencia Then Exit Sub

    'Unload Me
    MsgBox "Proceso realizado correctamente", vbInformation
End Sub

Private Function RealizarDevolucionTransferencia() As Boolean
Dim Gasto As Currency
Dim cad As String
Dim jj As Long
Dim CtaBan As String

    RealizarDevolucionTransferencia = False
   
    
    vSql = "DELETE FROM tmpfaclin WHERE codusu =" & vUsu.Codigo
    Conn.Execute vSql
    
    
        '                                               numero        serie     vto
        vSql = "INSERT INTO tmpfaclin (codusu, codigo, "
        vSql = vSql & IIf(Cobros, "NumFac ", "nomserie") & ", Fecha, NUmSerie, NIF,"
        vSql = vSql & "Imponible,  ImpIVA,total,cta,cliente,ctabase" & IIf(Cobros, "", ",numfac")
        vSql = vSql & ") VALUES (" & vUsu.Codigo & ","
        
    BancoRem = ""
    CtaBan = ""
    For jj = 1 To lwCobros.ListItems.Count
        If Me.lwCobros.ListItems(jj).Checked Then
                                        'cdofaccl
            If Cobros Then
                cad = jj & "," & Val(lwCobros.ListItems(jj).SubItems(1)) & ",'"
            Else
                cad = jj & "," & DBSet(lwCobros.ListItems(jj).SubItems(1), "T") & ",'"
            End If
            
                                    'fecfaccl                                                   SERIE
            cad = cad & Format(lwCobros.ListItems(jj).Tag, FormatoFecha) & "','" & lwCobros.ListItems(jj).Text
                                    'numvencimiento numorden
            cad = cad & "'," & Val(lwCobros.ListItems(jj).SubItems(2)) & ","
            ImporteQueda = ImporteFormateado(lwCobros.ListItems(jj).SubItems(6))
            If Cobros Then ImporteQueda = ImporteQueda * -1
            cad = cad & TransformaComasPuntos(CStr(ImporteQueda)) & ","
            
            
            cad = cad & TransformaComasPuntos(CStr(Gasto)) & ","
            cad = cad & TransformaComasPuntos(CStr(ImporteQueda)) & ",'"
            'Cuenta cliente, y banco
            cad = cad & lwCobros.ListItems(jj).SubItems(4) & "','"
            cad = cad & RecuperaValor(vRemesa, 3) & "','"
          
          
            cad = cad & lwCobros.ListItems(jj).SubItems(11) & "'"
          
            
            If Not Cobros Then cad = cad & "," & jj
            cad = cad & ")"
            cad = vSql & cad
            If Not Ejecuta(cad) Then Exit Function
            
            If BancoRem <> "" Then
                If lwCobros.ListItems(jj).SubItems(11) <> BancoRem Then
                    MsgBox "Devoluciones a dos bancos distintos", vbExclamation
                    Exit Function
                End If
            Else
                BancoRem = lwCobros.ListItems(jj).SubItems(11)
            End If
            

            
        End If
    Next jj
    
    
    'OK. Ya tengo grabada la temporal con los recibos que devuelvo. Ahora
    'hare:
    '       - generar un asiento con los datos k devuelvo
    '       - marcar los cobros como devueltos, añadirle el gasto y insertar en la
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
    
    
    vRemesa = Text3(5).Text & "|" & Text3(6).Text & "|" & BancoRem & "|" & Text1(11).Text & "|"
    
  
    
   
    If Me.txtCtaNormal(11).Text = "" Then
        vRemesa = vRemesa & "Dev. transferencia: " & Text3(5).Text & "/" & Text3(6).Text
    Else
        vRemesa = vRemesa & "Dev transf de Cta : " & txtCtaNormal(11).Text & "  " & Me.txtDCtaNormal(11).Text
    End If
    vRemesa = vRemesa & "|1|"
    
    Dim CodDev As String
    CodDev = ""
    jj = 0 'No hay gastos banco
    
    If RealizarLaDevolucionTransferencia(Cobros, CDate(Text1(11)), jj > 0, CtaBan, vRemesa, ValoresDevolucionRemesa) Then
    
        RealizarDevolucionTransferencia = True
        
        For jj = lwCobros.ListItems.Count To 1 Step -1
            If lwCobros.ListItems(jj).Checked Then lwCobros.ListItems.Remove jj
        Next jj
        
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
        If Numtrans > 0 Then
            
                PonerFocoLw lwCobros
        Else
                PonerFoco Text3(5)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
    Limpiar Me
    PrimeraVez = True
    Me.Icon = frmppal.Icon
    
    
    'Cago los iconos
    CargaImagenesAyudas Me.imgCtaNorma, 1, "Seleccionar cuenta"
    CargaImagenesAyudas Me.Image1, 2
    CargaImagenesAyudas imgRem, 1, "Seleccionar transferencia"
    CargaImagenesAyudas imgConcepto, 1, "Concepto"
   
    
     If Cobros Then
         SQL = "COBROS"
         Label5(3).ForeColor = &H800000
     Else
         SQL = "PAGOS"
         Label5(3).ForeColor = &H80&
     End If
     Label5(3).Caption = "Devolución transferencia " & LCase(SQL)
     FrameDevlucionRe.visible = True
     Caption = "Devolucion transferencia (" & UCase(SQL) & ")"
     W = FrameDevlucionRe.Width
     H = FrameDevlucionRe.Height
     Text1(11).Text = Format(Now, "dd/mm/yyyy")
     
     
     'Ofertamos el haber de la forma de pago recibo bancario
     SQL = DevuelveDesdeBD("amphacli", "tipofpago", "tipoformapago", "4")
     If SQL <> "" Then Combo2(0).ListIndex = CInt(SQL)
         

    
    OpcionAnt = 0
    CargaCabecera
    
    If Numtrans > 0 Then
        Text3(5).Text = Numtrans
        Text3(6).Text = Anotrans
        'Text3_LostFocus (5)  --> Si no lo hace dos veces
    End If
    
    
    
    Me.Height = H + 560
    Me.Width = W + 90
    
End Sub

Private Sub CargaCabecera()
    
    
    'If OpcionAnt = Opcion Then Exit Sub
    
    
    lwCobros.ColumnHeaders.Clear
    
    
        ' en el caso de devolucion desde fichero mostramos el codigo de devolucion
        lwCobros.ColumnHeaders.Add , , "Serie", 1020
        lwCobros.ColumnHeaders.Add , , "Factura", 2140
        lwCobros.ColumnHeaders.Add , , "Vto", 650
        lwCobros.ColumnHeaders.Add , , "F. Vto", 1350
        lwCobros.ColumnHeaders.Add , , "Cuenta", 1900
        lwCobros.ColumnHeaders.Add , , "Nombre", 4200
        lwCobros.ColumnHeaders.Add , , "Importe", 1450, 1
        
        lwCobros.ColumnHeaders.Add , , "FraOrden", 0
        lwCobros.ColumnHeaders.Add , , "ImporteOrden", 0
        lwCobros.ColumnHeaders.Add , , "Remesa", 0, 1
        lwCobros.ColumnHeaders.Add , , "Año", 0
        lwCobros.ColumnHeaders.Add , , "Banco", 0
        
        'lwCobros.ColumnHeaders.Add , , "Devolución", 3800, 0
        
        
    

    
    lwCobros.SortKey = 7
    lwCobros.SortOrder = lvwAscending
    lwCobros.Sorted = True
    
    OpcionAnt = Cobros

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set Rs = Nothing
    Set miRsAux = Nothing
    
    
End Sub



Private Sub frmBas2_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(Image1(11).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
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
    Image1(11).Tag = Index
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

    
        'Selecciona forma de pago
        For i = 1 To Me.lwCobros.ListItems.Count
    
            Me.lwCobros.ListItems(i).Checked = Index = 1
        Next
    
    
    CalcularTotal
End Sub

Private Sub imgConcepto_Click(Index As Integer)
  
    Set frmCon = New frmConceptos
    frmCon.DatosADevolverBusqueda = "0|"
    frmCon.Show vbModal
    Set frmCon = Nothing
    
End Sub

Private Sub imgCtaNorma_Click(Index As Integer)

        If Index <> 6 Then

               Set frmCCtas = New frmColCtas
               SQL = ""
               frmCCtas.DatosADevolverBusqueda = "0"
               frmCCtas.Show vbModal
               
               Set frmCCtas = Nothing
               If SQL <> "" Then
                   txtCtaNormal(Index).Text = SQL
                   txtCtaNormal_LostFocus Index
               End If
            
        End If
            
End Sub


Private Sub imgRem_Click(Index As Integer)
'    i = Index
'    Set frmRe = New frmTESRemesas
'    frmRe.Tipo = SubTipo  'Para abrir efectos o talonesypagares
'    frmRe.DatosADevolverBusqueda = "1|"
'    frmRe.Show vbModal
'    Set frmRe = Nothing
    'Por si ha puesto los datos
    
    Set frmBas2 = New frmBasico2
    SQL = ""
    AyudaTrasnferencia frmBas2, , "tipotrans = " & IIf(Cobros, 1, 0) & " AND subtipo = 0"
    
    Set frmBas2 = Nothing

    If SQL <> "" Then
        
        Text3(5).Text = RecuperaValor(SQL, 1)
        Text3(6).Text = RecuperaValor(SQL, 2)
        Text3_LostFocus 6
        
    End If
    
    
End Sub

Private Sub lwCobros_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'Reordenar
    
        
        i = ColumnHeader.Index
        If ColumnHeader.Index = 1 Then i = 7
        If ColumnHeader.Index = 6 Then i = 8
        i = i - 1
    
    
    
    If lwCobros.SortKey = i Then
        If lwCobros.SortOrder = lvwAscending Then
            lwCobros.SortOrder = lvwDescending
        Else
            lwCobros.SortOrder = lvwAscending
        End If
    Else
        
        lwCobros.SortOrder = lvwAscending
        lwCobros.SortKey = i
    End If
End Sub



Private Sub CalcularTotal()
Dim i As Integer
Dim Total As Currency

    Total = 0
    For i = 1 To Me.lwCobros.ListItems.Count
        If Me.lwCobros.ListItems(i).Checked Then
            Total = Total + Me.lwCobros.ListItems(i).SubItems(6)
        End If
    Next i
    
    If Total <> 0 Then
        Me.Text1(4).Text = Format(Total, FormatoImporte)
    Else
        Text1(4).Text = ""
    End If
    
End Sub


Private Sub lwCobros_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    CalcularTotal
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
            MsgBox "Campo debe ser numérico: " & .Text, vbExclamation
            .Text = ""
            PonerFoco Text3(Index)
        Else
            
            LimpiarLin Me, "FrameDevDesdeFichero"
            LimpiarLin Me, "FrameDevDesdeVto"
            
            If Text3(5).Text <> "" And Text3(6).Text <> "" Then
                If RemesaCorrecta Then
                    CargaList
                Else
                    Text3(5).Text = ""
                    Text3(6).Text = ""
                    Me.lwCobros.ListItems.Clear
                End If
            Else
                Me.lwCobros.ListItems.Clear
            End If
        End If
        
        'Para que vaya a la tabla y traiga datos remesa
'        If Index = 3 Or Index = 4 Then CamposRemesaAbono
    End With
End Sub

Private Function RemesaCorrecta() As Boolean
        
    On Error GoTo eRemesaCorrecta
        
     RemesaCorrecta = False
        Numtrans = 0
        Anotrans = 0
        
        SQL = "Select * from transferencias where codigo =" & Text3(5).Text
        SQL = SQL & " AND anyo =" & Text3(6).Text
        SQL = SQL & " AND tipotrans = " & IIf(Cobros, 1, 0)
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Rs.EOF Then
            SQL = "Ninguna transferencia con esos valores."
            SQL = SQL & "  Trans: " & Text3(5).Text & " / " & Text3(6).Text
            MsgBox SQL, vbExclamation
            Rs.Close
            Set Rs = Nothing
            Exit Function
        End If
        
        
        'Tiene valor
        If Rs!Situacion = "A" Then
            MsgBox "Transferencia abierta. Sin llevar al banco.", vbExclamation
            Rs.Close
            Set Rs = Nothing
            Exit Function
        End If
        
        
        
        If Asc(Rs!Situacion) < Asc("Q") Then
            MsgBox "Transferencia sin contabilizar.", vbExclamation
            Rs.Close
            Set Rs = Nothing
            Exit Function
        End If
        RemesaCorrecta = True
        
        Numtrans = Text3(5).Text
        Anotrans = Text3(6).Text
        
        
eRemesaCorrecta:
    

End Function

Private Sub CargarValores()
Dim Importe As Currency
Dim GastoDevolGral As Currency
Dim CadenaVencimiento As String
    
    
    If Anotrans > 0 Then
            SQL = "Select * from remesas where codigo =" & Text3(5).Text
            SQL = SQL & " AND anyo =" & Text3(6).Text
            SQL = SQL & " AND situacion = 'Q'"
        
     Else
        
            'Desde la cuenta
            Set Rs = New ADODB.Recordset
            
            SQL = "situacion = 'Q' "
            If Me.txtCtaNormal(11).Text <> "" Then SQL = SQL & " AND codmacta='" & Me.txtCtaNormal(11).Text & "'"
            
            SQL = "Select codrem,anyorem,NUmSerie,numfactu,numorden from cobros where " & SQL
            Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Rs.EOF Then
                SQL = "Ninguna pertenece a ninguna remesa "
                MsgBox SQL, vbExclamation
                Rs.Close
                Exit Sub
            End If
            Rs.Close
            Set Rs = Nothing
            
    End If
    
    
    
    If Anotrans > 0 Then
        
            SQL = "Select * from remesas where codigo =" & Text3(5).Text
            SQL = SQL & " AND anyo =" & Text3(6).Text
            SQL = SQL & " AND situacion = 'Q'"
    Else
            SQL = "Select distinct remesas.* from remesas where situacion = 'Q' "
            SQL = SQL & " and (codigo, anyo) in (select codrem, anyorem from cobros where codmacta = " & DBSet(Me.txtCtaNormal(11).Text, "T") & ")"
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then

        
            SQL = Rs!Codigo & "|" & Rs!Anyo & "|" & Rs!codmacta & "|" & Text1(11).Text & "|"

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
                'Voy a contabiñizar los gastos.
                'Vere si tiene CC
                If vParam.autocoste Then
                    If DevuelveDesdeBD("codccost", "bancos", "codmacta", Rs!codmacta, "T") = "" Then
                        MsgBox "Va a contabilizar los gastos pero no esta configurado el Centro de coste para el banco: " & Rs!codmacta, vbExclamation
                        Rs.Close
                        Set Rs = Nothing
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        'Depues del gasto
        SQL = SQL & "|"
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

   
    
    Set lwCobros.SmallIcons = frmppal.imgListComun16
    
    lwCobros.ListItems.Clear
    
    vSql = ""
    
    If Anotrans > 0 Then
            If Cobros Then
                vSql = " AND transfer =" & DBSet(Text3(5).Text, "N")
                vSql = vSql & " AND anyorem =" & DBSet(Text3(6).Text, "N")
                                
            Else
                vSql = " AND nrodocum =" & DBSet(Text3(5).Text, "N")
                vSql = vSql & " AND anyodocum =" & DBSet(Text3(6).Text, "N") & " AND ctaconfirm is null "
           End If
    Else
            If Cobros Then
            
                vSql = " and (transfer,anyorem) in (select codigo, anyo from transferencias where tipotrans=1 and situacion IN ('Q','Y','Z') )"
           
                vSql = " AND codmacta = " & DBSet(txtCtaNormal(11).Text, "T") & vSql
    
            Else
                 vSql = " AND ctaconfirm is null  AND anyodocum>0 and nrodocum >0  "
                vSql = " AND codmacta = " & DBSet(txtCtaNormal(11).Text, "T") & vSql
            
            End If
    
    End If
  
   
        If Cobros Then
            vSql = " from cobros where true " & vSql
            vSql = vSql & " ORDER BY numserie,numfactu"
            
            vSql = "SELECT NUmSerie,right(concat('0000000',NumFactu),7) numfactu,numorden,codmacta,nomclien,Gastos , ImpVenci,anyorem  anyodocum    ,transfer nrodocum   ,FecFactu, FecVenci " & vSql
            
        Else
            vSql = " from pagos where true " & vSql
            vSql = vSql & " ORDER BY numserie,numfactu"
            
            vSql = "SELECT NUmSerie,NumFactu,numorden,codmacta,nomprove nomclien,0 Gastos , impefect ImpVenci,anyodocum   ,nrodocum   ,FecFactu, fecefect FecVenci" & vSql
            
            
        End If
        
        Set miRsAux = New ADODB.Recordset
        lwCobros.ListItems.Clear
        miRsAux.Open vSql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        jj = 1
        While Not miRsAux.EOF
            Set Itm = lwCobros.ListItems.Add(, "C" & jj)
            Itm.Text = miRsAux!NUmSerie
            
            Itm.SubItems(1) = DBLet(miRsAux!NumFactu, "T")
            Itm.SubItems(2) = miRsAux!numorden
            Itm.SubItems(3) = Format(miRsAux!FecVenci, "dd/mm/yyyy")
            Itm.SubItems(4) = miRsAux!codmacta
            Itm.SubItems(5) = miRsAux!nomclien
            ImporteQueda = DBLet(miRsAux!Gastos, "N")
            'No lo pongo con el importe de gastos pq pudiera ser k habiendo sido devuelto, no quiera
'            ' cobrarle gastos
'            If DBLet(miRsAux!Devuelto, "N") = 1 Then
'                Itm.SmallIcon = 42
'            End If
            ImporteQueda = ImporteQueda + miRsAux!ImpVenci
            Itm.SubItems(6) = Format(ImporteQueda, FormatoImporte)
            
            'Para la ordenacion
            'Por si ordena por ser-fac
            Itm.SubItems(7) = Mid(miRsAux!NUmSerie & "   ", 1, 3) & Itm.SubItems(1)
            'Por si ordena por importe
            Itm.SubItems(8) = Format(miRsAux!ImpVenci * 100, "0000000000")
            
                    
            'remesas
            Itm.SubItems(9) = miRsAux!nrodocum
            Itm.SubItems(10) = miRsAux!anyodocum
            Itm.SubItems(11) = DevuelveValor("select codmacta from transferencias where codigo = " & DBSet(miRsAux!nrodocum, "N") & " and anyo = " & DBSet(miRsAux!anyodocum, "N"))
            
            
            Itm.Tag = miRsAux!FecFactu
            
            jj = jj + 1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
    
        Me.Refresh
        Screen.MousePointer = vbDefault
    
   
    
    ' si viene de fichero no dejamos marcar ni desmarcar
    lwCobros.Enabled = True
    imgCheck(0).Enabled = True
    imgCheck(1).Enabled = True
    
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
    i = 0
    If txtConcepto(Index).Text <> "" Then
        If Not IsNumeric(txtConcepto(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            i = 1
        Else
            
            SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtConcepto(Index).Text, "N")
            If SQL = "" Then
                MsgBox "Concepto no existe", vbExclamation
                i = 1
            End If
        End If
    End If
    Me.txtDConcpeto(Index).Text = SQL
    If i = 1 Then
        txtConcepto(Index).Text = ""
        PonerFoco txtConcepto(Index)
    Else
        SQL = "select ampdecli from tipofpago where tipoformapago = 4"
        i = DevuelveValor(SQL)
        PosicionarCombo Me.Combo2(0), i
    End If
End Sub

Private Sub txtCtaNormal_GotFocus(Index As Integer)
    ObtenerFoco txtCtaNormal(Index)
End Sub
    
Private Sub txtCtaNormal_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCtaNormal_LostFocus(Index As Integer)
Dim DevfrmCCtas As String
       
        DevfrmCCtas = Trim(txtCtaNormal(Index).Text)
        i = 0
        If DevfrmCCtas <> "" Then
            If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
                
            Else
                MsgBox SQL, vbExclamation
                If Index < 3 Or Index = 9 Or Index = 10 Or Index = 11 Then
                    DevfrmCCtas = ""
                    SQL = ""
                End If
            End If
            i = 1
        Else
            SQL = ""
        End If
        
        txtCtaNormal(Index).Text = DevfrmCCtas
        txtDCtaNormal(Index).Text = SQL
        If DevfrmCCtas = "" And i = 1 Then
            PonerFoco txtCtaNormal(Index)
        End If
        VisibleCC
        'limpiamos los otros frames
        If txtCtaNormal(11).Text <> "" Then
            
            LimpiarLin Me, "FrameDevDesdeFichero"
            LimpiarLin Me, "FrameDevDesdeRemesa"
            Numtrans = 0
            Anotrans = 0
            CargaList
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
    
        
        If Index = 6 Or Index = 7 Or Index = 2 Then
           
            If InStr(1, txtImporte(Index).Text, ",") > 0 Then
                Valor = ImporteFormateado(txtImporte(Index).Text)
            Else
                Valor = CCur(TransformaPuntosComas(txtImporte(Index).Text))
            End If
            txtImporte(Index).Text = Format(Valor, FormatoImporte)
        End If
        
End Sub


Private Sub VisibleCC()
Dim B As Boolean

    B = False
    If vParam.autocoste Then
        If txtCtaNormal(11).Text <> "" Then
                SQL = "|" & Mid(txtCtaNormal(11).Text, 1, 1) & "|"
                If InStr(1, CuentasCC, SQL) > 0 Then B = True
        End If
    End If
End Sub



Private Sub LanzaBuscaGrid(Opcion As Integer)

End Sub




Private Sub PonerValoresPorDefectoDevilucionRemesa()
Dim FP As Ctipoformapago

    On Error GoTo EPonerValoresPorDefectoDevilucionRemesa
    
    
    Set FP = New Ctipoformapago
    FP.Leer vbTipoPagoRemesa
    Me.txtConcepto(1).Text = FP.condecli
    'Ampliaciones
    Combo2(0).ListIndex = FP.ampdecli
    
    'Que carge el concepto
    txtConcepto_LostFocus 1
    Set FP = Nothing
    Exit Sub
EPonerValoresPorDefectoDevilucionRemesa:
    MuestraError Err.Number, "PonerValoresPorDefectoDevilucionRemesa"
    Set FP = Nothing
End Sub




