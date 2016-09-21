VERSION 5.00
Begin VB.Form frmTESPreguntas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "frmTESPreguntas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAyudaContabilizacion 
      Height          =   8775
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8775
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Text            =   "frmTESPreguntas.frx":000C
         Top             =   6000
         Width           =   8295
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "frmTESPreguntas.frx":0012
         Top             =   4080
         Width           =   8295
      End
      Begin VB.TextBox Text1 
         Height          =   1815
         Left            =   3360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "frmTESPreguntas.frx":0018
         Top             =   720
         Width           =   5175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmTESPreguntas.frx":001E
         Left            =   240
         List            =   "frmTESPreguntas.frx":0031
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   720
         Width           =   2895
      End
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   1800
         TabIndex        =   7
         Top             =   2640
         Width           =   5175
         Begin VB.CheckBox chkPorFechaVenci 
            Caption         =   "Contab. fecha vto."
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   1815
         End
         Begin VB.CheckBox chkContrapar 
            Caption         =   "Agrupar apunte bancario"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   11
            Top             =   960
            Width           =   2055
         End
         Begin VB.CheckBox chkAsiento 
            Caption         =   "Asiento por pago"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   10
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox chkGenerico 
            Caption         =   "Cuenta genérica"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox chkVtoCuenta 
            Caption         =   "Agrupar vtos por cuenta"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   8
            Top             =   600
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   2
         Left            =   7440
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Ejemplo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Contabilizacion vtos"
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
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame frameTalon 
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtTalon 
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Text            =   "Cuenta"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.CommandButton cmdTalon 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   2
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Talón / pagare"
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
         Index           =   5
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmTESPreguntas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0:     Descripcion talon/pagare
    '1:     Cuenta para los cobros genericos
    '2:     Ayuda para las opciones de contabilizacion de los efectos
    
Public vTexto As String
Dim T As String

Private Sub chkAsiento_Click(Index As Integer)
    Incompatibilidad
    PonetTextoAsiento
End Sub

Private Sub chkContrapar_Click(Index As Integer)
    
    Incompatibilidad
    PonetTextoAsiento
    
End Sub
Private Sub Incompatibilidad()
    If chkContrapar(1).Value = 1 Then
        If chkAsiento(1).Value = 1 Then
            MsgBox "Es incompatible agrupar apunte bancario y asiento por pago", vbExclamation
        End If
   End If
    
End Sub

Private Sub chkGenerico_Click(Index As Integer)
    PonetTextoAsiento
End Sub

Private Sub chkPorFechaVenci_Click()
    PonetTextoAsiento
End Sub

Private Sub chkVtoCuenta_Click(Index As Integer)
PonetTextoAsiento
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdTalon_Click()
    CadenaDesdeOtroForm = txtTalon.Text
    Unload Me
End Sub



Private Sub Combo1_Click()
    PonerTexto Combo1.ListIndex
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
    Me.Icon = frmPpal.Icon
    frameTalon.Visible = False
    Me.FrameAyudaContabilizacion.Visible = False
    Select Case Opcion
    Case 0
        
        Caption = "Talón / Pagaré"
        
        frameTalon.Visible = True
        H = frameTalon.Height + 150
        W = frameTalon.Width
        txtTalon.Text = vTexto
    Case 2
        Caption = "Ayuda"
        H = Me.FrameAyudaContabilizacion.Height
        W = Me.FrameAyudaContabilizacion.Width
        FrameAyudaContabilizacion.Visible = True
        Text1.Text = ""
        PonerTxtoVencimientos
        PonetTextoAsiento
    End Select
    Me.cmdCancelar(Opcion).Cancel = True
    
    Me.Height = H + 360
    Me.Width = W + 90
End Sub


Private Function PonerTexto(QueOpcion As Integer)


    Select Case QueOpcion
    Case 0
        T = "Todos los apuntes bancarios se resumirán en uno por asiento. "
        T = T & "No detallará documento ni ampliación" & vbCrLf & vbCrLf
        T = T & "En caso de no marcarse esta opción por cada linea de vencimiento habrá" & vbCrLf
        T = T & "un  apunte a la cuenta bancaria seleccionada." & vbCrLf
    Case 1
        T = "Todos los vencimientos para una cuenta seran agrupados en una única linea de asiento." & vbCrLf
        T = T & "Si hay más de un vencimiento no detallará documento ni ampliacion" & vbCrLf
        T = T & "En caso de no marcarse esta opción por cada linea de vencimiento habrá una linea de apunte" & vbCrLf
    Case 2
        T = "Generará un asiento por cada pago/cobro que haga." & vbCrLf
        T = T & "En caso de no marcarse esta opción se creará un asiento" & vbCrLf
        T = T & "por todos los vencimientos en la fecha seleccionada" & vbCrLf
    Case 3
        T = "En lugar de la fecha de contabilización que se pide , utilizará la fecha del vencimiento." & vbCrLf
        T = T & "Comprobará que esta en ejercicio actual/siguiente" & vbCrLf
        T = T & "En caso de no ser asi utilza la fecha de contabilización." & vbCrLf
        T = T & "Generara un apunte por cada fecha (obviamente)." & vbCrLf
    Case 4
        T = "En esta opción se piede una cuenta de cobro/pago genérica." & vbCrLf
        T = T & "Sustituye la cuenta del cobro/pago por la solicitada pasando esta a ser la " & vbCrLf
        T = T & "cuenta de contabilización ..." & vbCrLf
    
    End Select
    T = vbCrLf & UCase(Me.Combo1.List(Combo1.ListIndex)) & vbCrLf & "-----------------------------------------------------" & vbCrLf & vbCrLf & T
    Text1.Text = T
End Function


Private Function PonerTxtoVencimientos()
    T = " Cta    Factura      Fec.Vto      Banco   Importe  GENERICA    FecContab." & vbCrLf
    T = T & "--------------------------------------------------------------------------" & vbCrLf
    T = T & " 43001   S/123     10/02/2007    57201     85.46     43099       31/03/2007  " & vbCrLf
    T = T & " 43001   S/251     11/03/2007    57201     49.12     43099       31/03/2007  " & vbCrLf
    T = T & " 43001   S/252     11/03/2007    57201     97.54     43099       31/03/2007  " & vbCrLf
    T = T & " 43002   A/301     11/03/2007    57201     13.73     43099       31/03/2007  " & vbCrLf
    Text2.Text = T
End Function


Private Function PonetTextoAsiento()
    T = " NºAsto    Fecha      Cta     Docum     Contrpar    Debe     Haber " & vbCrLf
    T = T & "--------------------------------------------------------------------------" & vbCrLf
    '   T = T & "" & vbCrLf
    
    
    If Me.chkAsiento(1).Value Then
        'Un asiento por pago
        'Cada pago ire con un asiento
        If Me.chkPorFechaVenci.Value = 1 Then
            'Por fecha vto
            If chkGenerico(1).Value = 1 Then
                FechaVtoVariosAsientoCtaGENERICO chkVtoCuenta(1).Value = 1
            Else
                FechaVtoVariosAsientoSinAgruparCta chkVtoCuenta(1).Value = 1
            End If
        Else
            'Fecha contabilizacion
            If chkGenerico(1).Value = 1 Then
                'Ceunta generica
                FechaContabVariosAsientoCtaGenerica chkVtoCuenta(1).Value = 1
                
            Else
                'cuenta vto
                FechaContabVariosAsientoSinAgruparCta chkVtoCuenta(1).Value = 1
            End If
        End If
    Else
        'Todos juntos(si puede). Dependera de la fecha
        
        
        If Me.chkPorFechaVenci.Value = 1 Then
            'FECHA VTO
            If chkGenerico(1).Value = 1 Then
                'Utilizamos la cuenta geerica
                If chkVtoCuenta(1).Value = 1 Then
                    FechaVtoUnAsientoAgruparCtaGENERICO
                Else
                    FechaVtoUnAsientoSinAgruparCtaGENERICO chkContrapar(1).Value = 0
                End If
            Else
                'La cta vto
                If chkVtoCuenta(1).Value = 1 Then
                    'AGrupo por los vtos
                    FechaVtoUnAsientonAgruparCta chkContrapar(1).Value = 0
                Else
                    FechaVtoUnAsientoSinAgruparCta chkContrapar(1).Value = 0
                End If
            End If 'generico
        Else
            If chkGenerico(1).Value = 1 Then
                'Cuenta GENERICA
                If chkVtoCuenta(1).Value = 1 Then
                    FechaContabUnAsientonAgruparCtaGenerico chkContrapar(1).Value = 0
                Else
                    FechaContabUnAsientoSinAgruparCtaGENERICO chkContrapar(1).Value = 0
                End If
            Else
                'Fecha contabilizacion
                If chkVtoCuenta(1).Value = 1 Then
                    'AGrupamos por cuenta
                    FechaContabUnAsientonAgruparCta chkContrapar(1).Value = 0
                Else
                    'Detallamos las cuentas
                    FechaContabUnAsientoSinAgruparCta chkContrapar(1).Value = 0
                
                End If
        
            End If 'de chkGenerico
        End If ' de If Me.chkPorFechaVenci.Value=1 Then
    End If
    
    
        
    
    
    Text3.Text = T
End Function


Private Sub FechaContabUnAsientoSinAgruparCta(DetallaBanco As Boolean)
    If DetallaBanco Then
        T = T & "   1    31/03/2007    43001   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43001       85.46" & vbCrLf
        
        T = T & "   ""     ""            43001   S/251     57201               49.12" & vbCrLf
        T = T & "   ""     ""            57201   S/251     43001       49.12" & vbCrLf
        
        T = T & "   ""     ""            43001   S/252     57201               97.54" & vbCrLf
        T = T & "   ""     ""            57201   S/252     43001       97.54" & vbCrLf
        
        T = T & "   ""     ""            43002   A/301     57201               13.73" & vbCrLf
        T = T & "   ""     ""            57201   A/301     43002       13.73" & vbCrLf
        
    Else
        'Todo el apunte del banco junto
        T = T & "   1    31/03/2007    43001   S/123     57201               85.46" & vbCrLf
        'T = T & "   ""     ""            57201   S/123     43001       85.46" & vbCrLf
        
        T = T & "   ""     ""            43001   S/251     57201               49.12" & vbCrLf
        'T = T & "   ""     ""            57201   S/251     43001       49.12" & vbCrLf
        
        T = T & "   ""     ""            43001   S/252     57201               97.54" & vbCrLf
        'T = T & "   ""     ""            57201   S/252     43001       97.54" & vbCrLf
        
        T = T & "   ""     ""            43002   A/301     57201               13.73" & vbCrLf
        'TOdo el banco junto
        T = T & "   ""     ""            57201                         245.85" & vbCrLf
    End If
End Sub

Private Sub FechaContabUnAsientonAgruparCta(DetallaBanco As Boolean)
    If DetallaBanco Then
        T = T & "   1    31/03/2007    43001   3Vtos     57201              232.12" & vbCrLf
        T = T & "   ""     ""            57201                        232.12" & vbCrLf
        
        T = T & "   ""     ""            43002   A/301     57201               13.73" & vbCrLf
        T = T & "   ""     ""            57201   A/301     43002       13.73" & vbCrLf
        
    Else
        'Todo el apunte del banco junto
        T = T & "   1    31/03/2007    43001   3Vtos     57201              232.12" & vbCrLf
        T = T & "   ""     ""            43002   A/301     57201               13.73" & vbCrLf
   
        'TOdo el banco junto
        T = T & "   ""     ""            57201                         245.85" & vbCrLf
    End If
End Sub

'Por fecha de contabilizacion, un asiento
Private Sub FechaContabUnAsientoSinAgruparCtaGENERICO(DetallaBanco As Boolean)
    If DetallaBanco Then
        T = T & "   1    31/03/2007    43099   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43099       85.46" & vbCrLf
        
        T = T & "   ""     ""            43099   S/251     57201               49.12" & vbCrLf
        T = T & "   ""     ""            57201   S/251     43099       49.12" & vbCrLf
        
        T = T & "   ""     ""            43099   S/252     57201               97.54" & vbCrLf
        T = T & "   ""     ""            57201   S/252     43099       97.54" & vbCrLf
        
        T = T & "   ""     ""            43099   A/301     57201               13.73" & vbCrLf
        T = T & "   ""     ""            57201   A/301     43099       13.73" & vbCrLf
        
    Else
        'Todo el apunte del banco junto
        T = T & "   1    31/03/2007    43099   S/123     57201               85.46" & vbCrLf
        'T = T & "   ""     ""            57201   S/123     43001       85.46" & vbCrLf
        
        T = T & "   ""     ""            43099   S/251     57201               49.12" & vbCrLf
        'T = T & "   ""     ""            57201   S/251     43001       49.12" & vbCrLf
        
        T = T & "   ""     ""            43099   S/252     57201               97.54" & vbCrLf
        'T = T & "   ""     ""            57201   S/252     43001       97.54" & vbCrLf
        
        T = T & "   ""     ""            43099   A/301     57201               13.73" & vbCrLf
        'TOdo el banco junto
        T = T & "   ""     ""            57201                         245.85" & vbCrLf
    End If
End Sub

Private Sub FechaContabUnAsientonAgruparCtaGenerico(DetallaBanco As Boolean)
    
        T = T & "   1    31/03/2007    43099   4Vtos     57201              245.85" & vbCrLf
        T = T & "   ""     ""            57201                         245.85" & vbCrLf
    
End Sub



Private Sub FechaVtoUnAsientoSinAgruparCta(DetallaBanco As Boolean)
    If DetallaBanco Then
        T = T & "   1    10/02/2007    43001   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43001       85.46" & vbCrLf
        
        T = T & "   2    11/03/2007    43001   S/251     57201               49.12" & vbCrLf
        T = T & "   2     ""            57201   S/251     43001       49.12" & vbCrLf
        
        T = T & "   2     ""            43001   S/252     57201               97.54" & vbCrLf
        T = T & "   2     ""            57201   S/252     43001       97.54" & vbCrLf
        
        T = T & "   2     ""            43002   A/301     57201               13.73" & vbCrLf
        T = T & "   2     ""            57201   A/301     43002       13.73" & vbCrLf
        
    Else
        'Todo el apunte del banco junto
        T = T & "   1    10/02/2007    43001   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43001       85.46" & vbCrLf
        
        T = T & "   2    11/03/2007    43001   S/251     57201               49.12" & vbCrLf
        'T = T & "   ""     ""            57201   S/251     43001       49.12" & vbCrLf
        
        T = T & "   2     ""            43001   S/252     57201               97.54" & vbCrLf
        'T = T & "   ""     ""            57201   S/252     43001       97.54" & vbCrLf
        
        T = T & "   2     ""            43002   A/301     57201               13.73" & vbCrLf
        'TOdo el banco junto
        T = T & "   ""     ""            57201                         160.39" & vbCrLf
    End If
End Sub

Private Sub FechaVtoUnAsientonAgruparCta(DetallaBanco As Boolean)
    If DetallaBanco Then
        T = T & "   1    10/02/2007    43001   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43001       85.46" & vbCrLf
        
        T = T & "   2     11/03/2007   43001   vto:2     57201               146.66" & vbCrLf
        T = T & "   2     ""            57201                        146.66" & vbCrLf
        
        T = T & "   2     ""            43002   A/301     57201               13.73" & vbCrLf
        T = T & "   2     ""            57201   A/301     43002       13.73" & vbCrLf
        
    Else
        'Todo el apunte del banco junto
        T = T & "   1    10/02/2007    43001   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43001       85.46" & vbCrLf
           
        T = T & "   2     11/03/2007   43001   vto:2     57201               146.66" & vbCrLf
        T = T & "   2     ""            43002   A/301     57201               13.73" & vbCrLf
        T = T & "   2     ""            57201   A/301     43002      160.39" & vbCrLf
        
    End If
End Sub


Private Sub FechaVtoUnAsientoSinAgruparCtaGENERICO(DetallaBanco As Boolean)
    If DetallaBanco Then
        T = T & "   1    10/02/2007    43099   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43099       85.46" & vbCrLf
        
        T = T & "   2    11/03/2007    43099   S/251     57201               49.12" & vbCrLf
        T = T & "   2     ""            57201   S/251     43099       49.12" & vbCrLf
        
        T = T & "   2     ""            43099   S/252     57201               97.54" & vbCrLf
        T = T & "   2     ""            57201   S/252     43099       97.54" & vbCrLf
        
        T = T & "   2     ""            43099   A/301     57201               13.73" & vbCrLf
        T = T & "   2     ""            57201   A/301     43099       13.73" & vbCrLf
        
    Else
        'Todo el apunte del banco junto
        T = T & "   1    10/02/2007    43099   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43099       85.46" & vbCrLf
        
        T = T & "   2    11/03/2007    43099   S/251     57201               49.12" & vbCrLf
        
        T = T & "   2     ""            43099   S/252     57201               97.54" & vbCrLf
         
        T = T & "   2     ""            43099   A/301     57201               13.73" & vbCrLf
         'TOdo el banco junto
        T = T & "   ""     ""            57201                         160.39" & vbCrLf
    End If
End Sub


Private Sub FechaVtoUnAsientoAgruparCtaGENERICO()

        T = T & "   1    10/02/2007    43099   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43099       85.46" & vbCrLf
        
        T = T & "   2    11/03/2007    43099   Vto:3     57201              160.39" & vbCrLf
        T = T & "   2     ""            57201             43099      160.39" & vbCrLf
        

End Sub




Private Sub FechaContabVariosAsientoSinAgruparCta(AgruparPorCuenta As Boolean)
   
   If AgruparPorCuenta Then
        T = T & "   1    31/03/2007    43001   3Vtos     57201              232.12" & vbCrLf
        T = T & "   ""     ""            57201                        232.12" & vbCrLf
        
        T = T & "   2    31/03/2007    43002   A/301     57201               13.73" & vbCrLf
        T = T & "   ""     ""            57201   A/301     43002       13.73" & vbCrLf
       
    
   Else
   
        T = T & "   1    31/03/2007    43001   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43001       85.46" & vbCrLf
        
        T = T & "   2    31/03/2007    43001   S/251     57201               49.12" & vbCrLf
        T = T & "   ""     ""            57201   S/251     43001       49.12" & vbCrLf
        
        T = T & "   3    31/03/2007    43001   S/252     57201               97.54" & vbCrLf
        T = T & "   ""     ""            57201   S/252     43001       97.54" & vbCrLf
        
        T = T & "   4    31/03/2007    43002   A/301     57201               13.73" & vbCrLf
        T = T & "   ""     ""            57201   A/301     43002       13.73" & vbCrLf
        
   End If
End Sub


Private Sub FechaContabVariosAsientoCtaGenerica(AgruparPorCuenta As Boolean)
   
   If AgruparPorCuenta Then
        T = T & "   1    31/03/2007    43099   4Vtos     57201              245.85" & vbCrLf
        T = T & "   ""     ""            57201                         245.85" & vbCrLf
           
    
   Else
   
        T = T & "   1    31/03/2007    43099   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43099       85.46" & vbCrLf
        
        T = T & "   2    31/03/2007    43099   S/251     57201               49.12" & vbCrLf
        T = T & "   ""     ""            57201   S/251     43099       49.12" & vbCrLf
        
        T = T & "   3    31/03/2007    43099   S/252     57201               97.54" & vbCrLf
        T = T & "   ""     ""            57201   S/252     43099       97.54" & vbCrLf
        
        T = T & "   4    31/03/2007    43099   A/301     57201               13.73" & vbCrLf
        T = T & "   ""     ""            57201   A/301     43099       13.73" & vbCrLf
        
   End If
End Sub



Private Sub FechaVtoVariosAsientoCtaGENERICO(agruparcuenta As Boolean)
    If agruparcuenta Then
    
    
        T = T & "   1    10/02/2007    43099   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43099       85.46" & vbCrLf
        
        T = T & "   2    11/03/2007    43099   Vto:3     57201              160.39" & vbCrLf
        T = T & "   2     ""             57201             43099      160.39" & vbCrLf
    
    
    Else
        T = T & "   1    10/02/2007    43099   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43099       85.46" & vbCrLf
        
        T = T & "   2    11/03/2007    43099   S/251     57201               49.12" & vbCrLf
        T = T & "   2     ""            57201   S/251     43099       49.12" & vbCrLf
        
        T = T & "   3    11/03/2007    43099   S/252     57201               97.54" & vbCrLf
        T = T & "   3     ""            57201   S/252     43099       97.54" & vbCrLf
        
        T = T & "   4    11/03/2007    43099   A/301     57201               13.73" & vbCrLf
        T = T & "   4     ""            57201   A/301     43099       13.73" & vbCrLf
    End If
End Sub



Private Sub FechaVtoVariosAsientoSinAgruparCta(AgruparCta As Boolean)
    
   If Not AgruparCta Then
        T = T & "   1    10/02/2007    43001   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43001       85.46" & vbCrLf
        
        T = T & "   2    11/03/2007    43001   S/251     57201               49.12" & vbCrLf
        T = T & "   2     ""            57201   S/251     43001       49.12" & vbCrLf
        
        T = T & "   3    11/03/2007    43001   S/252     57201               97.54" & vbCrLf
        T = T & "   3     ""            57201   S/252     43001       97.54" & vbCrLf
        
        T = T & "   4    11/03/2007    43002   A/301     57201               13.73" & vbCrLf
        T = T & "   4     ""            57201   A/301     43002       13.73" & vbCrLf
        
    Else
        'Todo el apunte del banco junto
        T = T & "   1    10/02/2007    43001   S/123     57201               85.46" & vbCrLf
        T = T & "   ""     ""            57201   S/123     43001       85.46" & vbCrLf
        
        T = T & "   2    11/03/2007    43001   Vtos2     57201              146.66" & vbCrLf
        T = T & "   ""     ""            57201                        146.66" & vbCrLf
        
        T = T & "   3    11/03/2007    43002   A/301     57201               13.73" & vbCrLf
        T = T & "   3     ""            57201   A/301     43002       13.73" & vbCrLf
        
    End If
End Sub

