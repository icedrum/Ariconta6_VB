VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESRemesasGrab 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "1"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "frmTESRemesasGrab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCambioRemesa 
      Height          =   5355
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7725
      Begin VB.CommandButton cmdRemeTipo1 
         Caption         =   "Crear Soporte"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   4320
         TabIndex        =   20
         Top             =   4800
         Width           =   1515
      End
      Begin VB.Frame FrameTipo1_2 
         Height          =   3375
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   7425
         Begin VB.CheckBox chkAgrupaVtosFichero 
            Caption         =   "Agrupar vencimientos cliente"
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
            Left            =   3690
            TabIndex        =   19
            Top             =   2850
            Width           =   3585
         End
         Begin VB.CheckBox chkAnticipoCredito 
            Caption         =   "Anticipo Crédito"
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
            Left            =   3690
            TabIndex        =   18
            Top             =   2460
            Width           =   2745
         End
         Begin VB.OptionButton optSepaXML 
            Caption         =   "Fecha vencimiento del recibo"
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
            Left            =   3660
            TabIndex        =   17
            Top             =   2040
            Width           =   3315
         End
         Begin VB.OptionButton optSepaXML 
            Caption         =   "Fecha de cobro digitada"
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
            Left            =   3660
            TabIndex        =   16
            Top             =   1680
            Width           =   3075
         End
         Begin VB.CheckBox chkSEPA_GraboNIF 
            Caption         =   "COR1"
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   30
            Top             =   3600
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox chkSEPA_GraboNIF 
            Caption         =   "SEPA 19 - Empresas CIF"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   3600
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2295
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
            Index           =   18
            Left            =   1470
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   1860
            Width           =   1335
         End
         Begin VB.TextBox Text7 
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
            Left            =   270
            MaxLength       =   3
            TabIndex        =   14
            Top             =   1860
            Width           =   885
         End
         Begin VB.TextBox Text7 
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
            Left            =   210
            MaxLength       =   40
            TabIndex        =   13
            Top             =   1110
            Width           =   7005
         End
         Begin VB.ComboBox cmbReferencia 
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
            ItemData        =   "frmTESRemesasGrab.frx":000C
            Left            =   4980
            List            =   "frmTESRemesasGrab.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   480
            Width           =   2235
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
            Index           =   9
            Left            =   210
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   450
            Width           =   1275
         End
         Begin VB.ComboBox cboTipoRemesa 
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
            ItemData        =   "frmTESRemesasGrab.frx":003F
            Left            =   1740
            List            =   "frmTESRemesasGrab.frx":0041
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   480
            Width           =   3105
         End
         Begin VB.Image ImageAyuda 
            Height          =   240
            Index           =   0
            Left            =   3330
            Top             =   1950
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   18
            Left            =   2550
            Top             =   1590
            Width           =   240
         End
         Begin VB.Label Label7 
            Caption         =   "F.Present."
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
            Index           =   3
            Left            =   1500
            TabIndex        =   27
            Top             =   1620
            Width           =   1260
         End
         Begin VB.Label Label8 
            Caption         =   "Suf.OEM"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   300
            TabIndex        =   26
            Top             =   1620
            Width           =   915
         End
         Begin VB.Label Label8 
            Caption         =   "Identificacion ordenante"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   270
            TabIndex        =   25
            Top             =   870
            Width           =   1905
         End
         Begin VB.Label Label7 
            Caption         =   "Referencia domiciliacion"
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
            Index           =   2
            Left            =   4980
            TabIndex        =   24
            Top             =   210
            Width           =   2385
         End
         Begin VB.Label Label7 
            Caption         =   "Norma"
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
            Left            =   1740
            TabIndex        =   23
            Top             =   210
            Width           =   1155
         End
         Begin VB.Label Label7 
            Caption         =   "F. COBRO"
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
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   210
            Width           =   990
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   9
            Left            =   1230
            Top             =   210
            Width           =   240
         End
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   825
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   7425
         Begin VB.TextBox Text3 
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
            Index           =   7
            Left            =   4980
            TabIndex        =   28
            Text            =   "Text3"
            Top             =   360
            Width           =   2205
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
            Index           =   2
            Left            =   4500
            TabIndex        =   5
            Text            =   "Text3"
            Top             =   345
            Width           =   465
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
            Index           =   1
            Left            =   2400
            TabIndex        =   4
            Text            =   "Text3"
            Top             =   345
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
            Index           =   0
            Left            =   840
            TabIndex        =   3
            Text            =   "Text3"
            Top             =   345
            Width           =   1035
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Situación"
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
            Left            =   3480
            TabIndex        =   8
            Top             =   360
            Width           =   945
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
            Index           =   1
            Left            =   1800
            TabIndex        =   7
            Top             =   360
            Width           =   495
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
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   675
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   6090
         TabIndex        =   22
         Top             =   4800
         Width           =   1425
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
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
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   150
         Width           =   5175
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6240
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTESRemesasGrab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As Byte
    '
    '0 .- Pedir datos para ordenar cobros
    
    '3.- Reclamaciones
    '4.- Remesas
    
    
    '5.- Pregunta numero TALON pagare
    
    'Cambio situacion remesa
    '----------------------------
    '6.-  De A a B.   Generar banco
            
    '8.- Contabilizar remesa
        
    '9.- Devolucion remesa
        
    '10.- Mostrar vencimientos impagdos

    '11.- ACERCA DE
        
    '12  - Eliminar vtos
    
    '13.- Deuda total consolidada
    '14.-   "         ""      pero desde hasta
        
        
    '15.- Realizar transferencias
        
    '16.- Devolucion remesa desde fichero del banco
    '--------------------------------
    
    
    '17.- Eliminar informacion HCO remesas
    
    '18.- Selección de gastos para el listado de tesoreria
    
    '19.- Contabilizar gastos
    
    '20.- Seleccion de empresas disponibles, para el usuario
    
    
    '21- Listado pagos (cobros donde se indican las cuentas que quiero que apar
    
    
    'Mas sobre remesas.
    '22.- Cancelacion cliente
    '23.- Confirmacion remesa
    
    
    
    '24.- Impresion de todos los tipos de recibos
    
    '25.- Cambiar banco y/o fecha vto para la remesa de talon pagare
    
    '28 .- Devolucion remesa desde un vto
    
    
    '29 .- Recaudacion ejecutiva
    
    
    '31 .- Reclamaciones por email.
            'Tendra los que tienen email
    
    
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


Dim Rs As ADODB.Recordset
Dim SQL As String
Dim i As Integer
Dim IT As ListItem  'Comun
Dim PrimeraVez As Boolean
Dim Cancelado As Boolean
Dim CuentasCC As String

Private Sub cboTipoRemesa_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cmbReferencia_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub


'
'
'
Private Function CrearDisco() As Boolean
Dim B As Boolean
Dim FCobro As String
Dim AgruparVtos As Boolean
Dim IdFicheroGenerado As String
Dim HaceAgr As Integer

        CrearDisco = False
        If cboTipoRemesa.ListIndex < 0 Then
            MsgBox "Seleccione la norma para la remesa", vbExclamation
            Exit Function
        End If
    
        'El identificador REFERENCIA solo es valido para la norma 19
        If Me.cmbReferencia.ListIndex = 2 Then
            B = cboTipoRemesa.ListIndex = 0 Or cboTipoRemesa.ListIndex = 3
            If Not B Then
                MsgBox "Campo 'Referencia del recibo.' solo es válido para la norma 19", vbExclamation
                Exit Function
            End If
        End If
                
                
        If Text1(9).Text = "" Then
            MsgBox "Fecha cobro en blanco", vbExclamation
            Exit Function
        End If
        
        If Text1(18).Text = "" Then
            MsgBox "Fecha presentacion en blanco", vbExclamation
            Exit Function
        End If
        
        
        If Trim(Text7(0).Text) = "" Then Text7(0).Text = UCase(vEmpresa.nomempre)
        AgruparVtos = False
        HaceAgr = 0
        If Me.chkAgrupaVtosFichero.Value = 1 Then
            HaceAgr = ComprobacionesAgrupaFichero
            If HaceAgr = -1 Then Exit Function 'Error o cancelacion
            AgruparVtos = HaceAgr > 0
        Else
        
            If optSepaXML(0).Value Then MsgBox "Ha seleccionado fecha de cobro digitada: " & Text1(9).Text, vbInformation
        End If
        
                
        
        B = False
        If cboTipoRemesa.ListIndex = 0 Or cboTipoRemesa.ListIndex = 3 Then
            FCobro = Text1(9).Text
            If optSepaXML(1).Value Then FCobro = ""  'Ha selccionado por vencimiento
        
            SQL = Mid(Text7(1).Text & "   ", 1, 3) & "|" & Mid(Text7(0).Text & Space(40), 1, 40) & "|"
            If GrabarDisketteNorma19(App.Path & "\tmpRem.ari", Text3(0).Text & "|" & Text3(1).Text & "|", Text1(18).Text, SQL, Me.cmbReferencia.ListIndex, FCobro, True, chkSEPA_GraboNIF(0).Value = 1, chkSEPA_GraboNIF(1).Value = 1, cboTipoRemesa.ListIndex = 0, chkAnticipoCredito.Value = 1, IdFicheroGenerado, AgruparVtos) Then
                SQL = App.Path & "\tmpRem.ari"
                'Copio el disquete
                B = CopiarArchivo
            End If
        End If
        
        
        
        If B Then
            
            
            
            Set miRsAux = New ADODB.Recordset
            SQL = "Select * from remesas WHERE codigo=" & Text3(0).Text
            SQL = SQL & " AND anyo =" & Text3(1).Text
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                'Si a estaba generado el fichero, guardo en LOG
                If Asc(DBLet(miRsAux!Situacion, "T")) >= Asc("B") Then
                    'YA se habia generado
                    SQL = "Remesa  : " & Format(miRsAux!Codigo, "0000") & "/" & miRsAux!Anyo & "    " & DBLet(miRsAux!Descripcion, "T") & vbCrLf
                    SQL = SQL & "Banco: " & miRsAux!codmacta & " " & DevuelveDesdeBD("nommacta", "cuentas", "codmacta", miRsAux!codmacta, "T") & vbCrLf
                    SQL = SQL & "Situacion: " & miRsAux!Situacion & " Importe : " & miRsAux!Importe & "€ "
                    CuentasCC = "codrem= " & Text3(0).Text & " AND anyorem =" & Text3(1).Text & " AND 1"
                    CuentasCC = DevuelveDesdeBD("count(*)", "cobros", CuentasCC, "1")
                    SQL = SQL & " Vtos: " & CuentasCC
                    If DBLet(miRsAux!LlevaAgrupados, "N") > 0 Then SQL = SQL & " Cli. agr: " & miRsAux!LlevaAgrupados
                    SQL = SQL & vbCrLf
                    CuentasCC = DBLet(miRsAux!IdFicheroSEPA, "T")
                    If CuentasCC <> "" Then
                        CuentasCC = "ID_Fich: " & CuentasCC & "       Usuario: " & DBLet(miRsAux!usurioFich, "T")
                        SQL = SQL & CuentasCC
                    End If
                    CuentasCC = ""
                    vLog.Insertar 31, vUsu, SQL
                End If
            End If
            miRsAux.Close
            Set miRsAux = Nothing
            'Cambio la situacion de la remesa IdFicheroSEPA     usurioFich
            
            SQL = "UPDATE Remesas SET fecremesa = '" & Format(Text1(9).Text, FormatoFecha)
            SQL = SQL & "' , tipo = " & cboTipoRemesa.ListIndex & ", Situacion = 'B'"
            SQL = SQL & " , IdFicheroSEPA = " & DBSet(IdFicheroGenerado, "T") & ", LlevaAgrupados = " & HaceAgr
            SQL = SQL & " , usurioFich = " & DBSet(vUsu.Login, "T")
            
            
            SQL = SQL & " WHERE codigo=" & Text3(0).Text
            SQL = SQL & " AND anyo =" & Text3(1).Text
            If Ejecuta(SQL) Then CadenaDesdeOtroForm = "OK"
            
            
            
            If CadenaDesdeOtroForm = "OK" Then
            
                Set miRsAux = New ADODB.Recordset
                If Not UpdatearCobrosRemesa Then MsgBox "Error updateando cobros remesa", vbExclamation
                Set miRsAux = Nothing
                
            End If
            
            MsgBox "Fichero creado con exito", vbInformation
            CrearDisco = True
        End If
        
        
        
        
        
End Function


Private Function UpdatearCobrosRemesa() As Boolean
Dim Im As Currency
    On Error GoTo EUpdatearCobrosRemesa
    UpdatearCobrosRemesa = False
    
    SQL = "Select * from cobros WHERE codrem=" & Text3(0).Text
    SQL = SQL & " AND anyorem =" & Text3(1).Text
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
            While Not miRsAux.EOF
                SQL = "UPDATE cobros SET fecultco = '" & Format(Text1(9).Text, FormatoFecha) & "', impcobro = "
                Im = miRsAux!ImpVenci
                If Not IsNull(miRsAux!Gastos) Then Im = Im + miRsAux!Gastos
                SQL = SQL & TransformaComasPuntos(CStr(Im))
                
                SQL = SQL & " ,siturem = 'B'"
                
                
                'WHERE
                SQL = SQL & " WHERE numserie='" & miRsAux!NUmSerie
                SQL = SQL & "' AND  numfactu =  " & miRsAux!NumFactu
                SQL = SQL & "  AND  fecfactu =  '" & Format(miRsAux!FecFactu, FormatoFecha)
                SQL = SQL & "' AND  numorden =  " & miRsAux!numorden
                'Muevo siguiente
                miRsAux.MoveNext
                
                'Ejecuto SQL
                If Not Ejecuta(SQL) Then MsgBox "Error: " & SQL, vbExclamation
            Wend
    End If
    miRsAux.Close
                    
                    
                    
    UpdatearCobrosRemesa = True
    Exit Function
EUpdatearCobrosRemesa:
    
End Function


Private Sub cmdRemeTipo1_Click(Index As Integer)
Dim B As Boolean
    Select Case Index
    Case 0
    Case 1
        'Generar diskete
        B = CrearDisco
        
        vControl.UltReferRem = CStr(cmbReferencia.ListIndex)
        vControl.Grabar
        
        If B Then Unload Me
    End Select
    
    
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
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
    CargaImagenesAyudas Me.Image1, 2
    CargaImagenesAyudas ImageAyuda, 3
    
    FrameCambioRemesa.visible = False
    
    Label5(1).Caption = "Generar soporte magnético"

    Caption = "Situacion remesas"
    
    FrameCambioRemesa.visible = True
    
    For H = 1 To 3
            Text3(H - 1).Text = RecuperaValor(CadenaDesdeOtroForm, H)
    Next H
    Text3(7).Text = DevuelveDesdeBD("descsituacion", "usuarios.wtiposituacionrem", "situacio", Text3(2).Text, "T")  ' RecuperaValor(CadenaDesdeOtroForm, 6)
    H = Val(RecuperaValor(CadenaDesdeOtroForm, 7))
    Text3(7).Tag = H
    'Por defecto
    chkSEPA_GraboNIF(0).visible = False
    chkSEPA_GraboNIF(1).visible = False
    chkSEPA_GraboNIF(0).Value = 0
    
    ImageAyuda(0).visible = vParamT.NuevasNormasSEPA
    
    If H = 2 Then
        SubTipo = vbPagare
    ElseIf H = 3 Then
        SubTipo = vbTalon
    Else
        SubTipo = vbTipoPagoRemesa
            
        cboTipoRemesa.Clear
        If vParamT.NuevasNormasSEPA Then
            chkSEPA_GraboNIF(0).visible = True
            chkSEPA_GraboNIF(1).visible = True
            SQL = CheckValueLeer("FCob")
            If SQL = "1" Then
                Me.optSepaXML(1).Value = True
            Else
                Me.optSepaXML(0).Value = True
            End If
            
            If vParamT.NormasFormatoXML Then
                'Añadimos el XML
                Me.cboTipoRemesa.AddItem "Adeudos directos SEPA XML"
                Me.cboTipoRemesa.ListIndex = 0
            End If
        End If
    End If

    'El cuarto valor sera la fecha remesa
    'CadenaDesdeOtroForm
    FrameTipo1_2.visible = (Opcion = 7)
        
    Text1(9).Text = Format(Now, "dd/mm/yyyy")
    Text1(18).Text = Text1(9).Text
    
    Me.cmbReferencia.ListIndex = Val(vControl.UltReferRem)
    
    Text7(0).Text = UCase(vEmpresa.nomempre)
    
    Set miRsAux = New ADODB.Recordset
    SQL = RecuperaValor(CadenaDesdeOtroForm, 5)
    SQL = "Select N1914GrabaNifDeudor,sufijoem from bancos where codmacta = '" & SQL & "'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Text7(1).Text = DBLet(miRsAux!sufijoem, "T")
    If vParamT.NuevasNormasSEPA Then chkSEPA_GraboNIF(0).Value = DBLet(miRsAux!N1914GrabaNifDeudor, "N")
    miRsAux.Close
    Set miRsAux = Nothing
        
    H = FrameCambioRemesa.Height + 240
    W = FrameCambioRemesa.Width
    CadenaDesdeOtroForm = ""
    
   
    Me.Height = H + 360
    Me.Width = W + 90
    
    H = Opcion
    If Opcion = 7 Then H = 6
    Me.cmdCancelar(H).Cancel = True
       
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Rs = Nothing
    Set miRsAux = Nothing
    If Opcion = 7 Then
        'La seleccion cobro o vencimiento
        CheckValueGuardar "FCob", Me.optSepaXML(1).Value   'seimpre tendremos true
    End If
    
    NumeroDocumento = "" 'Para reestrablecerlo siempre

End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(Image1(9).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text1(Index).Text <> "" Then frmC.Fecha = CDate(Text1(Index).Text)
    Image1(9).Tag = Index
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

Private Sub KEYpress(ByRef Tecla As Integer)
    If Tecla = 13 Then
        Tecla = 0
        SendKeys "{tab}"
    End If
End Sub


Private Sub ImageAyuda_Click(Index As Integer)
    
    Select Case Index
    Case 0
        If vParamT.NuevasNormasSEPA Then

            SQL = "Adeudos directos SEPA" & vbCrLf & vbCrLf & vbCrLf
            SQL = SQL & " Según la fecha seleccionada girará los vencimientos de la remesa:"
            
            SQL = SQL & vbCrLf & " Cobro.  Todos los vencimientos a esa fecha"
            SQL = SQL & vbCrLf & " Vencimiento.  Cada uno con su fecha"
            SQL = SQL & vbCrLf & vbCrLf & " Agrupar vencimentos. "
            SQL = SQL & "Generará una unica linea en el fichero generado agrupando los vencimientos de un mismo cliente."

        Else
            SQL = "Generacion antigua N19"
        End If
    End Select
    MsgBox ImageAyuda(Index).ToolTipText & vbCrLf & SQL, vbInformation
End Sub


Private Sub optSepaXML_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
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
        End If
    End With
End Sub

Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Function CopiarArchivo() As Boolean
On Error GoTo ECopiarArchivo

    CopiarArchivo = False
    'cd1.CancelError = True
    cd1.FileName = ""
    cd1.ShowSave
    If cd1.FileName <> "" Then
    
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo " & cd1.FileName & " ya existe" & vbCrLf & vbCrLf & "¿Sobreescribir?", vbQuestion + vbYesNo) = vbNo Then Exit Function
            Kill cd1.FileName
        End If
        'Hacemos la copia
        FileCopy SQL, cd1.FileName
        CopiarArchivo = True
    End If
    
    
    Exit Function
ECopiarArchivo:
    MuestraError Err.Number, "Copiar Archivo"
End Function



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
        For i = 0 To 1
            ' contatalonpte
            SQL = "pagarecta"
            If i = 1 Then SQL = "contatalonpte"
            CtaPte = (DevuelveDesdeBD(SQL, "paramtesor", "codigo", "1") = "1")
            
            'Repetiremos el proceso dos veces
            SQL = "Select * from scarecepdoc where fechavto<='" & Format(Text1(17).Text, FormatoFecha) & "'"
            SQL = SQL & " AND   talon = " & i
            Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                    'Si lleva cta puente habra que ver si esta contbilizada
                    J = 0
                    If CtaPte Then
                        If Val(Rs!Contabilizada) = 0 Then
                            'Veo si tiene lineas. S
                            SQL = DevuelveDesdeBD("count(*)", "slirecepdoc", "id", CStr(Rs!Codigo))
                            If SQL = "" Then SQL = "0"
                            If Val(SQL) > 0 Then
                                CuentasCC = CuentasCC & Rs!Codigo & " - No contabilizada" & vbCrLf
                                J = 1
                            End If
                        End If
                    End If
                    If J = 0 Then
                        'Si va benee
                        If Val(DBLet(Rs!llevadobanco, "N")) = 0 Then
                            SQL = DevuelveDesdeBD("count(*)", "slirecepdoc", "id", CStr(Rs!Codigo))
                            If SQL = "" Then SQL = "0"
                            If Val(SQL) > 0 Then
                                CuentasCC = CuentasCC & Rs!Codigo & " - Sin llevar a banco" & vbCrLf
                                J = 1
                            End If
                    
                        End If
                    End If
                    'Esta la borraremos
                    If J = 0 Then CualesEliminar = CualesEliminar & ", " & Rs!Codigo
                    
                    Rs.MoveNext
            Wend
            Rs.Close
            
            
            
        Next i
        
        

        
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
            i = InStr(J, CualesEliminar, ",")
            If i > 0 Then
                J = i + 1
                SQL = SQL & "X"
            End If
        Loop Until i = 0
        
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


Private Sub SituarComboReferencia(Leer As Boolean)
Dim NF As Integer
    
    On Error GoTo eSituarComboReferencia
    
    SQL = App.Path & "\cboremref.dat"
    If Leer Then
        i = 2
        If Dir(SQL, vbArchive) <> "" Then
            NF = FreeFile
            Open SQL For Input As #NF
            If Not EOF(NF) Then
                Line Input #NF, SQL
                If SQL <> "" Then
                    If IsNumeric(SQL) Then
                        If Val(SQL) > 0 And Val(SQL) < 3 Then i = Val(SQL)
                    End If
                End If
            End If
            Close #NF
            
        End If
        Me.cmbReferencia.ListIndex = i
    Else
        'GUARDAR
        If Me.cmbReferencia.ListIndex = 2 Then
            If Dir(SQL, vbArchive) <> "" Then Kill SQL
        Else
            Open SQL For Output As #NF
            Print #NF, Me.cmbReferencia.ListIndex
            Close #NF
        End If
    End If
    Exit Sub
eSituarComboReferencia:
    Err.Clear

End Sub



' -1: ERROR     0: Ninguno     1 Si que hay agrupaciones
Private Function ComprobacionesAgrupaFichero() As Integer
Dim C As String

    On Error GoTo eComrpobacionesAgrupacionFichero
    Set miRsAux = New ADODB.Recordset
    ComprobacionesAgrupaFichero = -1
    
    
    'Si la referencia es el recibo, no puede agrupar
    If cmbReferencia.ListIndex = 2 Then
        MsgBox "No puede agrupar indicando como referencia domiciliacion el recibo", vbExclamation
        Exit Function
    End If
    
    
    SQL = "select codmacta,count(*) from cobros where codrem = " & Me.Text3(0).Text & " AND anyorem=" & Text3(1).Text & " group by codmacta having count(*) >1"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Msg = ""
    CuentasCC = ""
    i = 0
    While Not miRsAux.EOF
        i = i + 1
        Msg = Msg & miRsAux!codmacta & "|"
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
    
    If i = 0 Then
        ComprobacionesAgrupaFichero = 0
        Exit Function
    End If
    
    'Comprobare la cuenta IBAN
    MsgErr = ""
    K = 0
    NumRegElim = 0
    For J = 1 To i
        SQL = " AND codmacta = '" & RecuperaValor(Msg, CInt(J)) & "' GROUP BY iban "
        SQL = "select  iban,count(*) ctos from cobros where codrem = " & Me.Text3(0).Text & " AND anyorem=" & Text3(1).Text & SQL
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If miRsAux.EOF Then
            Err.Raise 513, , "no se encuentran vtos. en select: " & Mid(miRsAux.Source, 30)
        Else
            C = DevuelveDesdeBD("concat(codmacta,' ',nommacta)", "cuentas", "codmacta", RecuperaValor(Msg, CInt(J)), "T") & "   Nº:" & Format(miRsAux!ctos, "0000")
            SQL = miRsAux!IBAN
            miRsAux.MoveNext
            If miRsAux.EOF Then
                'perfeto. SOLO un IBAN
                CuentasCC = CuentasCC & "X"
                NumRegElim = NumRegElim + Val(Right(C, 4))
                SQL = ""
            Else
                SQL = C & vbCrLf & SQL & " // " & miRsAux!IBAN & vbCrLf
            End If
        End If
        If SQL <> "" Then MsgErr = MsgErr & SQL & vbCrLf
        miRsAux.Close
    Next J
    
    
    Set miRsAux = Nothing
    
    If MsgErr <> "" Then
        MsgErr = "IBAN distinto para agrupacion por cliente:" & vbCrLf & MsgErr
        MsgBox MsgErr, vbExclamation
        
    Else
        SQL = ""
        If optSepaXML(0).Value Then SQL = "Ha seleccionado fecha de cobro digitada: " & Text1(9).Text & vbCrLf
        CuentasCC = "Clientes agrupados: " & Len(CuentasCC) & vbCrLf
        CuentasCC = CuentasCC & "Total vencimientos agrupados: " & NumRegElim & vbCrLf
        SQL = SQL & CuentasCC & " ¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then ComprobacionesAgrupaFichero = i
    End If
eComrpobacionesAgrupacionFichero:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
    CuentasCC = ""
End Function
