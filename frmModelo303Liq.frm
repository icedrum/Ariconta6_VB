VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModelo303Liq 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConcepto 
      Caption         =   "Selecci�n"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtCuota 
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
         Left            =   4680
         TabIndex        =   0
         Top             =   2400
         Width           =   2115
      End
      Begin VB.Frame FramePeriodo 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   90
         TabIndex        =   21
         Top             =   1290
         Width           =   3675
         Begin VB.TextBox txtperiodo 
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
            Left            =   960
            TabIndex        =   7
            Top             =   150
            Width           =   675
         End
         Begin VB.TextBox txtperiodo 
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
            Left            =   2670
            TabIndex        =   8
            Top             =   150
            Width           =   645
         End
         Begin VB.Label Label3 
            Caption         =   "Inicio"
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
            Index           =   26
            Left            =   270
            TabIndex        =   23
            Top             =   150
            Width           =   870
         End
         Begin VB.Label Label3 
            Caption         =   "Fin"
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
            Index           =   27
            Left            =   2220
            TabIndex        =   22
            Top             =   165
            Width           =   390
         End
      End
      Begin VB.ComboBox cmbPeriodo 
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
         ItemData        =   "frmModelo303Liq.frx":0000
         Left            =   330
         List            =   "frmModelo303Liq.frx":0002
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   930
         Width           =   3330
      End
      Begin VB.TextBox txtAno 
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
         Index           =   0
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   9
         Tag             =   "imgConcepto"
         Top             =   930
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Cuotas a compensar per�odos anteriores"
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
         Left            =   360
         TabIndex        =   24
         Top             =   2400
         Width           =   4125
      End
      Begin VB.Label lblFecha1 
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
         Left            =   2580
         TabIndex        =   14
         Top             =   3990
         Width           =   4095
      End
      Begin VB.Label lblFecha 
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
         Left            =   2580
         TabIndex        =   13
         Top             =   3630
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "Per�odo"
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
         Index           =   7
         Left            =   360
         TabIndex        =   12
         Top             =   570
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "A�o"
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
         Index           =   6
         Left            =   3990
         TabIndex        =   11
         Top             =   570
         Width           =   960
      End
   End
   Begin VB.Frame frameConceptoDer 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   7110
      TabIndex        =   15
      Top             =   0
      Width           =   4485
      Begin VB.CheckBox chk1 
         Caption         =   "Realizar apunte contable de cancelaci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   90
         TabIndex        =   3
         Top             =   4650
         Value           =   1  'Checked
         Width           =   4335
      End
      Begin VB.Frame FrameSeccion 
         BorderStyle     =   0  'None
         Height          =   3525
         Left            =   180
         TabIndex        =   19
         Top             =   1020
         Width           =   4185
         Begin MSComctlLib.ListView ListView1 
            Height          =   2880
            Index           =   1
            Left            =   60
            TabIndex        =   2
            Top             =   510
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   5080
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
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
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   3390
            Picture         =   "frmModelo303Liq.frx":0004
            ToolTipText     =   "Quitar al Debe"
            Top             =   120
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   3750
            Picture         =   "frmModelo303Liq.frx":014E
            ToolTipText     =   "Puntear al Debe"
            Top             =   120
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Empresas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   30
            TabIndex        =   20
            Top             =   180
            Width           =   1110
         End
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   1350
         TabIndex        =   1
         Top             =   570
         Width           =   1485
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3840
         TabIndex        =   16
         Top             =   270
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
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   2
         Left            =   1020
         Picture         =   "frmModelo303Liq.frx":0298
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   210
         TabIndex        =   17
         Top             =   570
         Width           =   690
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
      Left            =   10350
      TabIndex        =   5
      Top             =   5490
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccion 
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
      Index           =   1
      Left            =   8760
      TabIndex        =   4
      Top             =   5490
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
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
      Left            =   120
      TabIndex        =   18
      Top             =   5550
      Visible         =   0   'False
      Width           =   6855
   End
End
Attribute VB_Name = "frmModelo303Liq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 408

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************
'
'  3 espacios
'       -Los desde hasta,
'       -las opciones / ordenacion
'       -el tipo salida
'
' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************

    ' en tmpliquidaiva la columna cliente indica
    '                   0- Facturas clientes
    '                   1- Facturas clientes RECARGO EQUIVALENCIA
    '                   2- Facturas proveedores
    '                   3- Facturas Proveedores recargo equivalencia
    '                   4- Facturas Proveedores no deducible




Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1
Private frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1

Private Sql As String
Dim cad As String
Dim RC As String
Dim I As Integer
Dim IndCodigo As Integer
Dim tabla As String
Dim ImpTotal As Currency
Dim ImpCompensa As Currency
Dim Periodo As String
Dim Consolidado As String
Dim SqlLog As String

Dim vFecha1 As String
Dim vFecha2 As String
Dim M1 As Integer
Dim M2 As Integer
Dim vCta As String
Dim ImpLiqui As Currency


Private Sub cmbPeriodo_Validate(Index As Integer, Cancel As Boolean)
    
    If cmbPeriodo(0).ListIndex > 0 Then
        txtperiodo(0).Text = cmbPeriodo(0).ListIndex + 1
        txtperiodo(1).Text = cmbPeriodo(0).ListIndex + 1
    End If
    
    
    CargarFechas
    
End Sub

Private Sub cmdAccion_Click(Index As Integer)
Dim Pregunta As Boolean
Dim B As Boolean
Dim MultiEmpresa As Boolean


    If Not DatosOK Then Exit Sub
    
    
    ImpTotal = 0
    If txtCuota(0).Text <> "" Then
        ImpTotal = ImporteFormateado(txtCuota(0).Text)
    End If
    cadParam = cadParam & "CompensacionAnterior=" & ImpTotal & "|"
    numParam = numParam + 1
    ImpTotal = 0
    
    
    
'++
    'AHora generaremos la liquidacion para todos los periodos k abarque la seleecion
    Screen.MousePointer = vbHourglass
    'Guardamos el valor del chk del IVA
'--
'    ModeloIva False
    Label13.Caption = "Elimina datos anteriores"
    Label13.visible = True
    Label13.Refresh
    If GeneraLasLiquidaciones Then
        Label13.Caption = ""
        Label13.Refresh
        espera 0.5
        'Periodos
        Sql = ""
        For I = 0 To 1
            Sql = Sql & txtperiodo(I).Text & "|"
        Next I
        Sql = Sql & txtAno(0).Text & "|"
        I = 1
        
        Periodo = Sql & I & "|"
        
    
        SqlLog = "Periodo : " & txtAno(0) & " / " & Me.cmbPeriodo(0).Text & vbCrLf
        SqlLog = SqlLog & "Empresas : "
        
    
        'Empresas para consolidado
        Pregunta = True
        Sql = ""
        I = EmpresasSeleccionadas
        MultiEmpresa = I > 1
        If Not MultiEmpresa Then
            B = False
            For I = 1 To Me.ListView1(1).ListItems.Count
            
                SqlLog = SqlLog & "ariconta" & Me.ListView1(1).ListItems(I).Text & vbCrLf
                
                If ListView1(1).ListItems(I).Checked Then
                    
                    NumConta = Me.ListView1(1).ListItems(I).Text
                    
                    ImprimirAsientoContable
                    
                    CadenaDesdeOtroForm = ""
                    If HayRegParaInforme("tmpconext", "codusu=" & vUsu.Codigo) Then

                        Set frmMens = New frmMensajes
                        frmMens.Parametros = chk1.Value
                        frmMens.Opcion = 29
                        frmMens.Show vbModal
                        
                        Set frmMens = Nothing
    
                    End If

                    If CadenaDesdeOtroForm = "" Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    
                    
                    
                    If chk1.Value Then
                        If RealizarAsientoContable(MultiEmpresa) Then
                            B = True
                            Exit For
                        End If
                    Else
                        B = ActualizarLiquidacion(False, 0, 0, 0)
                        If B Then
                            B = True
                            Exit For
                        End If
                    End If
                End If
            Next I
        Else
            'Mas de una empresa
            Sql = "'Empresas seleccionadas:' + Chr(13) "
            B = False
            
            
            ImprimirAsientoContable
            
            CadenaDesdeOtroForm = ""
            If HayRegParaInforme("tmpconext", "codusu=" & vUsu.Codigo) Then

                Set frmMens = New frmMensajes
                frmMens.Parametros = chk1.Value
                frmMens.Opcion = 29
                frmMens.Show vbModal
                
                Set frmMens = Nothing

            End If

            If CadenaDesdeOtroForm = "" Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        
            
            
            SqlLog = ""
            For I = 1 To Me.ListView1(1).ListItems.Count
                If Me.ListView1(1).ListItems(I).Checked Then SqlLog = SqlLog & " .-" & ListView1(1).ListItems(I).SubItems(1) & vbCrLf
            Next I
         
            If chk1.Value Then
                
                B = RealizarAsientoContable(MultiEmpresa)
                 
            Else
                'No realiza el apunte, solo actualizamos
                For I = 1 To Me.ListView1(1).ListItems.Count
                    If Me.ListView1(1).ListItems(I).Checked Then
                        NumConta = ListView1(1).ListItems(I).Text
                        If NumConta = vEmpresa.codempre Then NumConta = 0
                        ActualizarLiquidacion False, 0, 0, NumConta
                    End If
                Next I
                B = True
            End If
         
         
            
        End If
        
        If B Then
            If SqlLog <> "" Then vLog.Insertar 15, vUsu, SqlLog
        
        
            MsgBox "Proceso realizado correctamente.", vbExclamation
            Unload Me
        End If


    
    End If
    Label13.visible = False
    Me.Refresh
    Screen.MousePointer = vbDefault


    
    
End Sub

Private Function ActualizarLiquidacion(DentroDeTrans As Boolean, NumAsiento As Long, NumDiari As Integer, codempre As Integer) As Boolean
Dim Sql As String
Dim I As Integer
    On Error GoTo eActualizarLiquidacion

    If Not DentroDeTrans Then Conn.BeginTrans

    ActualizarLiquidacion = False
    ' actualizamos los parametros
    I = IIf(codempre = 0, vEmpresa.codempre, codempre)
    Sql = "update ariconta" & I & ".parametros set anofactu = " & DBSet(txtAno(0).Text, "N")
    I = txtperiodo(0)
    Sql = Sql & ", perfactu = " & DBSet(I, "N")
    Conn.Execute Sql

    If codempre = 0 Then
        vParam.Anofactu = txtAno(0).Text
        vParam.perfactu = I

        If vParam.periodos = 0 Then
            I = I + 12
        End If
    End If
    
    Sql = IIf(codempre = 0, vEmpresa.codempre, codempre)
    Sql = "insert into ariconta" & Sql & ".liqiva (anoliqui,periodo,escomplem,importe,numdiari,numasien,fechaent) values ("
    Sql = Sql & DBSet(txtAno(0).Text, "N") & "," & DBSet(I, "N") & ",0," & DBSet(ImpLiqui, "N") & "," & DBSet(NumDiari, "N") & "," & DBSet(NumAsiento, "N") & "," & DBSet(txtFecha(2).Text, "F") & ")"
    Conn.Execute Sql
    
    If Not DentroDeTrans Then Conn.CommitTrans
    
    ActualizarLiquidacion = True
    Exit Function


eActualizarLiquidacion:
    If Not DentroDeTrans Then Conn.RollbackTrans
    MuestraError Err.Number, "Actualizar Liquidaci�n", Err.Description
End Function


Private Function RealizarAsientoContable(MultiEmpresa As Boolean) As Boolean
    If MultiEmpresa Then
        RealizarAsientoContable = ApunteContableMultiempresa
    Else
        RealizarAsientoContable = RealizarAsientoContableUnaEmpresa
    End If
End Function


'Como estaba , no toco nada. Lo dejo ASI, ejemplo ARIADNA
Private Function RealizarAsientoContableUnaEmpresa() As Boolean
Dim Mc As Contadores
Dim B As Boolean
Dim Numdocum As String
Dim Ampconce As String
Dim MaxPos As Long
Dim NomConce As String
Dim NumAsien As Long

    On Error GoTo eRealizarAsientoContable
    
    RealizarAsientoContableUnaEmpresa = False
    
    Set Mc = New Contadores
    
    Conn.BeginTrans
    
    I = FechaCorrecta2(CDate(txtFecha(2).Text), True)
    If I > 1 Then
        Err.Raise 513, , "Fecha incorrecta"
    Else
        If Mc.ConseguirContador("0", (I = 0), False) = 0 Then
            NumAsien = Mc.Contador
        Else
            Err.Raise 513, , "Conseguir contador"
        End If
    End If
    ' insertamos en cabecera
    Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion ) SELECT " & vParam.numdia303 & "," & DBSet(txtFecha(2).Text, "F") & "," & DBSet(NumAsien, "N")
    Sql = Sql & ",'Liquidaci�n de " & Me.cmbPeriodo(0).Text & " de " & txtAno(0).Text & "'," & DBSet(Now, "F") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Liquidaci�n'"
    Sql = Sql & " from parametros "
    Conn.Execute Sql
    
    
    NomConce = DevuelveValor("select nomconce from conceptos where codconce = " & vParam.conce303)
    Numdocum = "LIQ." & txtAno(0).Text & "-" & txtperiodo(1).Text
    
    If vParam.periodos = 0 Then
        Ampconce = NomConce & " Liq.303 " & txtperiodo(0).Text & "T"
    Else
        Ampconce = NomConce & " Liq.303 " & cmbPeriodo(0).Text
    End If
    
    MaxPos = DevuelveValor("select max(pos) from tmpconext where codusu = " & DBSet(vUsu.Codigo, "N"))
    
    ' insertamos en lineas
    Sql = "INSERT INTO hlinapu (numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,ctacontr) SELECT " & vParam.numdia303 & "," & DBSet(txtFecha(2).Text, "F") & "," & DBSet(NumAsien, "N")
    Sql = Sql & ",pos, cta," & DBSet(Numdocum, "T") & "," & DBSet(vParam.conce303, "N") & "," & DBSet(Ampconce, "T") & ",if(timported=0,null,timported), if(timporteh=0,null,timporteh), "
    If ImpLiqui > 0 Then
        Sql = Sql & "if(pos <> " & DBSet(MaxPos, "N") & "," & DBSet(vParam.CtaHPAcreedor, "T") & ",null) "
    Else
        Sql = Sql & "if(pos <> " & DBSet(MaxPos, "N") & "," & DBSet(vParam.CtaHPDeudor, "T") & ",null) "
    End If
    
    Sql = Sql & " from tmpconext where codusu =  " & vUsu.Codigo
    Sql = Sql & " order by pos "
    Conn.Execute Sql
    
    
    
    SqlLog = SqlLog & "Asiento contable: " & DBSet(NumAsien, "N") & " - " & txtFecha(2).Text & " - " & vParam.numdia303 & vbCrLf
    
    
    B = ActualizarLiquidacion(True, NumAsien, vParam.numdia303, 0)
    
    If B Then
        Conn.CommitTrans
        RealizarAsientoContableUnaEmpresa = True
        Exit Function
    End If
    
eRealizarAsientoContable:
    Conn.RollbackTrans
    MuestraError Err.Number, "Realizar Asiento contable", Err.Description
End Function


Private Function ApunteContableMultiempresa() As Boolean
Dim Empres As Collection
Dim Z As Integer
Dim vaBien As Boolean

    Set Empres = New Collection
    ApunteContableMultiempresa = False
    Sql = "SELECT distinct numasien from tmpconext where codusu =  " & vUsu.Codigo
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Empres.Add CInt(miRsAux!NumAsien)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    If Empres.Count = 0 Then
        MsgBox "Ningun dato en tabla temporal", vbExclamation
        Exit Function
    End If
    
    
    
    Conn.BeginTrans
    vaBien = True
    For Z = 1 To Empres.Count
        vaBien = RealizarAsientoContableDeLaEmpresaMulti(Empres.Item(Z))
        If Not vaBien Then Exit For
    Next Z
    If vaBien Then
        Conn.CommitTrans
        ApunteContableMultiempresa = True
    Else
        Conn.RollbackTrans
    End If
End Function

'
'Como estaba , no toco nada. Lo dejo ASI, ejemplo ARIADNA
Private Function RealizarAsientoContableDeLaEmpresaMulti(ByVal NumeroConta As Integer) As Boolean
Dim B As Boolean
Dim Numdocum As String
Dim Ampconce As String
Dim MaxPos As Long
Dim NomConce As String
Dim NumAsi As Long
Dim Diario As Integer
Dim CodConce As Integer
Dim SqlInsert As String
Dim Importe As Currency
Dim ImporteTemporal As Currency

    On Error GoTo eRealizarAsientoContable
    Set miRsAux = New ADODB.Recordset
    
    RealizarAsientoContableDeLaEmpresaMulti = False
    
    
    
    I = FechaCorrecta2(CDate(txtFecha(2).Text))
    
    Sql = "Select * from ariconta" & NumeroConta & ".contadores WHERE TipoRegi = '0' "
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If I = 0 Then
        NumAsi = miRsAux!contado1
        Sql = "contado1"
    Else
        'sigiente
        NumAsi = miRsAux!contado2
        Sql = "contado2"
    End If
    NumAsi = NumAsi + 1
    Sql = "UPDATE ariconta" & NumeroConta & ".contadores SET " & Sql & "= " & NumAsi & " WHERE TipoRegi = '0' "
    miRsAux.Close
    Conn.Execute Sql
    
    
    Sql = "diario303"
    Ampconce = DevuelveDesdeBD("conce303", "ariconta" & NumeroConta & ".parametros", "1", "1", "N", Sql)
    Diario = Val(Sql)
    CodConce = Val(Ampconce)
    
    
    ' insertamos en cabecera
    Sql = "INSERT INTO ariconta" & NumeroConta & ".hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion )"
    Sql = Sql & " VALUES (" & Diario & "," & DBSet(txtFecha(2).Text, "F") & "," & DBSet(NumAsi, "N")
    Sql = Sql & ",'Liquidaci�n de " & Me.cmbPeriodo(0).Text & " de " & txtAno(0).Text & "'," & DBSet(Now, "F") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Liquidaci�n')"
    Conn.Execute Sql
    
    
    NomConce = DevuelveValor("select nomconce from ariconta" & NumeroConta & ".conceptos where codconce = " & CodConce)
    Numdocum = "LIQ." & txtAno(0).Text & "-" & txtperiodo(1).Text
    

        If vParam.periodos = 0 Then
            Ampconce = NomConce & " Liq.303 " & txtperiodo(0).Text & "T"
        Else
            Ampconce = NomConce & " Liq.303 " & cmbPeriodo(0).Text
        End If

    
    ' insertamos en lineas
    Sql = "SELECT * "
    Sql = Sql & " from tmpconext where codusu =  " & vUsu.Codigo
    If NumeroConta <> vEmpresa.codempre Then Sql = Sql & " AND numasien= " & NumeroConta
    Sql = Sql & " order by pos "
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SqlInsert = ""
    K = 0
    Importe = 0
    While Not miRsAux.EOF
        K = K + 1
       
        
        Sql = miRsAux!Cta
        If NumeroConta = vEmpresa.codempre Then
            'Empresa desde la que lanzamsos
            'Las cuentas de las otras secciones van a partidas pendientes de aplicacion
            If miRsAux!NumAsien <> vEmpresa.codempre Then Sql = vParamT.Par_pen_apli
        End If
        Sql = DBSet(Sql, "T")
        Sql = ", (" & Diario & "," & DBSet(txtFecha(2).Text, "F") & "," & DBSet(NumAsi, "N") & "," & K & "," & Sql
        Sql = Sql & "," & DBSet(Numdocum, "T") & "," & CodConce & "," & DBSet(Ampconce, "T") & ","
        Sql = Sql & IIf(miRsAux!timported = 0, "null", DBSet(miRsAux!timported, "N")) & ","
        Sql = Sql & IIf(miRsAux!timporteH = 0, "null", DBSet(miRsAux!timporteH, "N")) & ",null)"
        
        Importe = Importe + miRsAux!timported - miRsAux!timporteH
        SqlInsert = SqlInsert & Sql
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    If NumeroConta <> vEmpresa.codempre Then
        'Si NO es la empresa origen hau que cuadrar
        If Importe < 0 Then
            Sql = "ctahpacreedor"
        Else
            Sql = "ctahpdeudor"
        End If
        Sql = DevuelveDesdeBD(Sql, "ariconta" & NumeroConta & ".parametros", "1", "1", "T")
        
        K = K + 1
        Sql = DBSet(Sql, "T") & "," & DBSet(Numdocum, "T")
        Sql = ", (" & Diario & "," & DBSet(txtFecha(2).Text, "F") & "," & DBSet(NumAsi, "N") & "," & K & "," & Sql
        Sql = Sql & "," & DBSet(CodConce, "N") & "," & DBSet(Ampconce, "T") & ","
        If Importe < 0 Then
            Sql = Sql & DBSet(Abs(Importe), "N") & ",NULL,NULL)"
        Else
            Sql = Sql & "NULL," & DBSet(Importe, "N") & ",NULL)"
        End If
            
        SqlInsert = SqlInsert & Sql
        
    End If
    
    Set miRsAux = Nothing
    SqlInsert = Mid(SqlInsert, 2)
    Sql = "INSERT INTO ariconta" & NumeroConta & ".hlinapu (numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,ctacontr) VALUES " & SqlInsert
    Conn.Execute Sql
    
    
    
    SqlLog = SqlLog & "Asiento contable: " & DBSet(NumAsi, "N") & " - " & txtFecha(2).Text & vbCrLf
    
    
    If NumeroConta <> vEmpresa.codempre Then
        'El importe que graba en el apunte es el de "SU" empresa, no el total de todas
        ImporteTemporal = ImpLiqui
        ImpLiqui = Importe
    Else
        'Pongo numeroconta a CERO, porque es la empresa desde la que estan lanzando la liquidacion
        NumeroConta = 0
    End If
    B = ActualizarLiquidacion(True, NumAsi, Diario, NumeroConta)
    
    'Reeestablezco el importe
    If NumeroConta = 0 Then NumeroConta = vEmpresa.codempre
    If NumeroConta <> vEmpresa.codempre Then ImpLiqui = ImporteTemporal
     
    
    If B Then
        
        RealizarAsientoContableDeLaEmpresaMulti = True
        
    End If
    Exit Function
eRealizarAsientoContable:
    
    MuestraError Err.Number, "Realizar Asiento contable: " & NumeroConta, Err.Description
End Function



Private Sub ImprimirAsientoContable()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim SqlInsert As String
Dim SqlInsert2 As String
Dim SqlValues As String
Dim SqlValues2 As String
Dim Importe As Currency
Dim vDebe As Currency
Dim vHaber As Currency
Dim I As Long
Dim codempre As Integer
Dim Aux As String

    On Error GoTo eImprimirAsientoContable
    NumConta = vEmpresa.codempre
    Sql = "delete from ariconta" & NumConta & ".tmpconext where codusu= " & vUsu.Codigo
    Conn.Execute Sql
    
    ' para visualizar los saldos
    Sql = "delete from ariconta" & NumConta & ".tmpconextcab where codusu= " & vUsu.Codigo
    Conn.Execute Sql
    
    ' codigo = 0 debe
    '          1 haber
    
    SqlInsert = "insert into ariconta" & NumConta & ".tmpconext(codusu,pos,cta,timported,timporteh,ampconce,numasien) values "
    SqlInsert2 = "insert into ariconta" & NumConta & ".tmpconextcab(codusu,cta,acumtotT) values "
    
    Sql = "select codempre,cliente, codmacta, sum(coalesce(ivas,0)) importe from ariconta" & NumConta & ".tmpliquidaiva where codusu = " & vUsu.Codigo
    Sql = Sql & " group by 1,2,3 "
    Sql = Sql & " having sum(coalesce(ivas,0)) <> 0"
    Sql = Sql & " order by 1,2,3 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SqlValues = ""
    I = 0
    While Not Rs.EOF
        I = I + 1
    
        Importe = DBLet(Rs!Importe, "N")
    
        SqlValues = SqlValues & "(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(I, "N") & "," & DBSet(Rs!codmacta, "T") & ","
    
        If DBLet(Rs!Cliente, "N") = 0 Then ' clientes
            If Importe >= 0 Then
                SqlValues = SqlValues & DBSet(Importe, "N") & "," & "0," ' clientes positivo al debe
            Else
                SqlValues = SqlValues & "0," & DBSet(Importe * (-1), "N") & "," ' clientes negativo al haber
            End If
        Else 'proveedores
            If Importe >= 0 Then
                SqlValues = SqlValues & "0," & DBSet(Importe, "N") & "," ' clientes positivo al haber
            Else
                SqlValues = SqlValues & DBSet(Importe * (-1), "N") & "," & "0," ' clientes negativo al debe
            End If
        End If
        'ampconce: Llevaremos el nommacta porque puede ser que sea de otras empresas
        'numasien: codempre
        codempre = Rs!codempre
        Aux = "ariconta" & codempre & ".cuentas"
        Aux = DevuelveDesdeBD("nommacta", Aux, "codmacta", Rs!codmacta, "T")
        If Aux = "" Then Aux = Rs!codmacta
        SqlValues = SqlValues & DBSet(Aux, "T") & "," & codempre & "),"
        
    
        
        ' cargamos cual es el saldo entre la fecha de inicio de ejercicio y la fecha de liquidacion
        Sql = "select abs(sum(coalesce(timported,0)) - sum(coalesce(timporteh,0))) from ariconta" & codempre & ".hlinapu where codmacta =  " & DBSet(Rs!codmacta, "T")
        Sql = Sql & " and fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vFecha2, "F")
    
        
        
        SqlValues2 = SqlValues2 & "(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(Rs!codmacta, "T") & "," & DBSet(DevuelveValor(Sql), "N") & "),"
        
        Rs.MoveNext
    Wend
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        
        Conn.Execute SqlInsert & SqlValues
        
        ' los saldos
        SqlValues2 = Mid(SqlValues2, 1, Len(SqlValues2) - 1)
        
        Conn.Execute SqlInsert2 & SqlValues2
        
        
        'Abril 2020
        If vParam.InscritoDeclarDUA Then
            Rs.Close
            'Para cada factura de DUA vempos el importe de IVA y a contra el proveedor lo descontamos del resltado
            'De momento, si tiene lo de inscrito en DUA, no puede consolidar
            Sql = " WHERE fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
            Sql = Sql & " and codopera = 6 " 'duas
            Sql = "Select codmacta, sum(totivas) importe ,nommacta   FROM ariconta" & NumConta & ".factpro " & Sql
            Sql = Sql & " GROUP by codmacta order by codmacta "
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SqlValues = ""
            While Not Rs.EOF
                I = I + 1
                Importe = DBLet(Rs!Importe, "N")
            
                SqlValues = SqlValues & ", (" & DBSet(vUsu.Codigo, "N") & "," & DBSet(I, "N") & "," & DBSet(Rs!codmacta, "T") & ","
                
                If Importe < 0 Then
                    SqlValues = SqlValues & "0," & DBSet(Importe * (-1), "N") & "," '  prov negativo al haber
                Else
                    SqlValues = SqlValues & DBSet(Importe, "N") & "," & "0," ' prov  al debe
                End If
                Aux = DBLet(Rs!Nommacta, "T")
                codempre = NumConta
                If Aux = "" Then
                    Aux = "ariconta" & codempre & ".cuentas"
                    Aux = DevuelveDesdeBD("nommacta", Aux, "codmacta", Rs!codmacta, "T")
                End If
                If Aux = "" Then Aux = "ERROR obteniedno cta"
                
                SqlValues = SqlValues & DBSet(Aux, "T") & "," & codempre & ")"
                
            
                'SqlValues2 = SqlValues2 & "(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(Rs!codmacta, "T") & "," & DBSet(DevuelveValor(SQL), "N") & "),"
            
                Rs.MoveNext
            Wend
            
            
            
            If SqlValues <> "" Then
                SqlValues = Mid(SqlValues, 2)
                Conn.Execute SqlInsert & SqlValues
            End If
        End If
        
    
        Sql = "select sum(timported) from ariconta" & NumConta & ".tmpconext where codusu = " & vUsu.Codigo
        vDebe = DevuelveValor(Sql)
        
        Sql = "select sum(timporteh) from ariconta" & NumConta & ".tmpconext where codusu = " & vUsu.Codigo
        vHaber = DevuelveValor(Sql)
    
        SqlValues = ""
        I = I + 1
        If vDebe - vHaber > 0 Then
            SqlValues = vParam.CtaHPAcreedor
            If SqlValues = "" Then
                SqlValues = "HP Acreedor"
                Aux = SqlValues
            Else
                Aux = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", SqlValues, "T")
                If Aux = "" Then Aux = SqlValues
            End If
            SqlValues = "(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(I, "N") & "," & DBSet(SqlValues, "T") & ",0," & DBSet(vDebe - vHaber, "N")
        Else
            If vDebe - vHaber < 0 Then
                SqlValues = vParam.CtaHPDeudor
                If SqlValues = "" Then
                    SqlValues = "HP Deudor"
                    Aux = SqlValues
                Else
                    Aux = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", SqlValues, "T")
                    If Aux = "" Then Aux = SqlValues
                End If
                SqlValues = "(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(I, "N") & "," & DBSet(SqlValues, "T") & "," & DBSet(vHaber - vDebe, "N") & ",0"
            End If
        End If
        If SqlValues <> "" Then
            SqlValues = SqlValues & "," & DBSet(Aux, "T") & "," & vEmpresa.codempre & ")"
            'Apunte de la diferencia debe - haber
            Conn.Execute SqlInsert & SqlValues
    
    
        End If
        ImpLiqui = vDebe - vHaber
    
    
    End If

    Set Rs = Nothing
    
    Exit Sub

eImprimirAsientoContable:
    MuestraError Err.Number, "Imprimir Asiento Contable", Err.Description
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub Form_Load()
    Me.Icon = frmppal.Icon
        
    'Otras opciones
    Me.Caption = "Liquidaci�n de Iva"

     
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
     
   
    CargarListView 1
    
    PonerPeriodoPresentacion303
     
     
    FrameSeccion.Enabled = vParam.EsMultiseccion
    
    'FramePeriodo.Enabled = (Me.cmbPeriodo(0).ListIndex = 0)
    'FramePeriodo.Visible = (Me.cmbPeriodo(0).ListIndex = 0)
    
    FramePeriodo.Enabled = False
    FramePeriodo.visible = False
    
    
    CargarFechas
    
    
    
    txtFecha(2).Text = Format(vFecha2, "dd/mm/yyyy")
     
    
End Sub

Private Sub CargarFechas()
    
    If vParam.periodos = 1 Then
        'Esamos en mensual
        If Me.cmbPeriodo(0).ListIndex > 11 Then
            MsgBox "Error en el periodo a tratar.", vbExclamation
            Exit Sub
        End If
        
        If txtAno(0).Text = "" Then Exit Sub
        
        vFecha1 = CDate("01/" & Me.cmbPeriodo(0).ListIndex + 1 & "/" & Me.txtAno(0))
        M1 = DiasMes(Me.cmbPeriodo(0).ListIndex + 1, Me.txtAno(0))
        vFecha2 = CDate(M1 & "/" & Me.cmbPeriodo(0).ListIndex + 1 & "/" & Me.txtAno(0))
        
    Else
        'IVA TRIMESTRAL
        If Me.cmbPeriodo(0).ListIndex > 4 Then
            MsgBox "Error en el periodo a tratar.", vbExclamation
            Exit Sub
        End If
        M2 = ((Me.cmbPeriodo(0).ListIndex) * 3) + 1
        vFecha1 = CDate("01/" & M2 & "/" & Me.txtAno(0))
        M2 = ((Me.cmbPeriodo(0).ListIndex) * 3) + 3
        M1 = DiasMes(CByte(M2), Me.txtAno(0))
        vFecha2 = CDate(M1 & "/" & M2 & "/" & Me.txtAno(0))
    End If
    
End Sub




Private Sub PonerPeriodoPresentacion303()

    cmbPeriodo(0).Clear
    If vParam.periodos = 0 Then
        'Liquidacion TRIMESTRAL
        
        For I = 1 To 4
            If I = 1 Or I = 3 Then
                CadenaDesdeOtroForm = "er"
            Else
                CadenaDesdeOtroForm = "�"
            End If
            CadenaDesdeOtroForm = I & CadenaDesdeOtroForm & " "
            Me.cmbPeriodo(0).AddItem CadenaDesdeOtroForm & " trimestre"
            Me.cmbPeriodo(0).ItemData(cmbPeriodo(0).NewIndex) = I
            
        Next I
    Else
        'Liquidacion MENSUAL
        For I = 1 To 12
            CadenaDesdeOtroForm = MonthName(I)
            CadenaDesdeOtroForm = UCase(Mid(CadenaDesdeOtroForm, 1, 1)) & LCase(Mid(CadenaDesdeOtroForm, 2))
            Me.cmbPeriodo(0).AddItem CadenaDesdeOtroForm
            Me.cmbPeriodo(0).ItemData(cmbPeriodo(0).NewIndex) = I
        Next
    End If
    
    
    'Leeremos ultimo valor liquidado
    
    txtAno(0).Text = vParam.Anofactu
    I = vParam.perfactu + 1
    If vParam.periodos = 0 Then
        NumRegElim = 4
    Else
        NumRegElim = 12
    End If
        
    If I > NumRegElim Then
            I = 1
            txtAno(0).Text = vParam.Anofactu + 1
    End If
    Me.cmbPeriodo(0).ListIndex = I - 1
     
     
    txtperiodo(0).Text = I 'Me.cmbPeriodo(0).ListIndex
    txtperiodo(1).Text = I 'Me.cmbPeriodo(0).ListIndex
    
     
    
    CadenaDesdeOtroForm = ""
End Sub


Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim I As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency

    Screen.MousePointer = vbHourglass
    
    Select Case Index
        ' tabla de codigos de iva
        Case 0
            For I = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(I).Checked = False
            Next I
        Case 1
            For I = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(I).Checked = True
            Next I
    End Select
    
    Screen.MousePointer = vbDefault


End Sub


Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1, 2
        IndCodigo = Index
    
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If txtFecha(Index).Text <> "" Then frmF.Fecha = CDate(txtFecha(Index).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco txtFecha(Index)
        
    End Select
    
    Screen.MousePointer = vbDefault

End Sub



Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub





Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        imgFec_Click Indice
    End Select
End Sub



'Cojera los importes y los formateara para los programitas de hacineda
Private Sub GeneraCadenaImportes()
Dim TotalClien As Currency
Dim TotalProve As Currency
Dim HayReg As Boolean
Dim Rs As ADODB.Recordset

    TotalClien = 0

    'En devuelveimporte
    ' Tipo 0:   11 enteros y 2 decimales
    ' Tipo 1:   2 ente y 2 decimales
    ' Tipo 2:   1 entero y 2 decimales
    ' tipo 3:   3 enetero y dos decimales

    
    Sql = "select iva,  bases, ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 0 "
    Sql = Sql & " order by 1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not Rs.EOF
        I = I + 1
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!IVA, "N"), 3
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    
    'por si hay menos de 3 porcentajes de iva hay que rellenarlos a ceros
    For J = I + 1 To 3
        DevuelveImporte 0, 0
        DevuelveImporte 0, 3
        DevuelveImporte 0, 0
    Next J
    
    Set Rs = Nothing
    
    'Adquisiciones intra
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 10 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    HayReg = False
    
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    ' Inversion de sujeto pasivo
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 12 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'modificacion bases y cuotas (no tenemos)
    DevuelveImporte 0, 0
    DevuelveImporte 0, 0
    
    
    'Los recargos
    Sql = "select iva,  bases, ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 1 "
    Sql = Sql & " order by 1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not Rs.EOF
        I = I + 1
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!IVA, "N"), 3
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    
    'por si hay menos de 3 porcentajes de iva hay que rellenarlos a ceros
    For J = I + 1 To 3
        DevuelveImporte 0, 0
        DevuelveImporte 0, 3
        DevuelveImporte 0, 0
    Next J
    
    Set Rs = Nothing
    
    'modificacion bases y cuotas del recargo de equivalencia (no tenemos)
    DevuelveImporte 0, 0
    DevuelveImporte 0, 0
    

    'total
    DevuelveImporte TotalClien, 0
    
    '------------------------------------------------------------------------
    '------------------------------------------------------------------------
    'DEDUCIBLE
    TotalProve = 0
    
'    'operaciones interiores
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'operaciones interiores BIENES INVERSION
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 30 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'importaciones
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 32 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'importaciones BIEN INVERSION
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 34 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    
    
    'adqisiciones intracom
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 36 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'adqisiciones intracom BIEN INVERSION
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 38 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing

    ' rectificacion de deducciones tampoco tenemos
    DevuelveImporte 0, 0
    DevuelveImporte 0, 0

    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 42 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
    End If
    
    Set Rs = Nothing
    

    DevuelveImporte 0, 0  'Regularizacion inversiones
    DevuelveImporte 0, 0  'Regularizacion por aplicacion del porcentaje def de prorrata

    
    'total a deducir
    DevuelveImporte TotalProve, 0
    
    
    'Diferencia
    DevuelveImporte TotalClien - TotalProve, 0  'Regularizacion inversiones
    
    ImpTotal = TotalClien - TotalProve
    
    
End Sub


'Ahora desde un importe, antes Desde un text box
Private Sub DevuelveImporte(Importe As Currency, Tipo As Byte)
Dim J As Integer
Dim Aux As String
Dim Resul As String

    Dim modelo As Integer
    modelo = 4

    Resul = ""
    If Importe < 0 Then
        Aux = ""
        Resul = "N"
    Else
        Aux = "0"
    End If
    Importe = Importe * 100
'++ hasta aqui

    
    'Tipo sera la mascara.
    ' si Modelo<>303
        ' Tipo 0:   11 enteros y 2 decimales
        'Else
        ' Tipo 0:   15 enteros y 2 decimales
    ' Tipo 1:   2 ente y 2 decimales
    ' Tipo 2:   1 entero y 2 decimales
    ' tipo 3:   3 enetero y dos decimales
    Select Case Tipo
    Case 1
        Aux = Aux & "000"
    Case 2
        Aux = Aux & "00"
    Case 3
        Aux = Aux & "0000"
    Case Else
        If modelo = 4 Then
            Aux = Aux & String(16, "0")  '15 enteros 2 decima  17-1
        Else
            Aux = Aux & String(10, "0")   '11 enteros 2 decimales  13-1
        End If
    End Select
    
    cad = cad & Resul & Format(Importe, Aux)
        
End Sub



Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = True
        
    
    indRPT = "0408-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "FacturasCliFecha.rpt"

    cadParam = cadParam & "pFecha=""" & txtFecha(2).Text & """|"
    numParam = numParam + 1
    
    Sql = ""
    If EmpresasSeleccionadas = 1 Then
        For I = 1 To Me.ListView1(1).ListItems.Count
            If ListView1(1).ListItems(I).Checked Then
                If Me.ListView1(1).ListItems(I).Text <> vEmpresa.nomempre Then Sql = Me.ListView1(1).ListItems(I).SubItems(1)
            End If
        Next I
    Else
        'Mas de una empresa
        Sql = "'Empresas seleccionadas:' + Chr(13) "
        For I = 1 To Me.ListView1(1).ListItems.Count
            Sql = Sql & " + '        " & Me.ListView1(1).ListItems(I).Text & "' + Chr(13)"
        Next I
    End If
    
    cadParam = cadParam & "empresas = """ & Sql & """|"
    numParam = numParam + 1
    

    cadParam = cadParam & "pPeriodo1=" & txtperiodo(0).Text & "|"
    cadParam = cadParam & "pPeriodo2=" & txtperiodo(1).Text & "|"
    cadParam = cadParam & "pAno=" & txtAno(0).Text & "|"
    numParam = numParam + 3
    
    
    cadFormula = "{tmpliquidaiva.codusu} = " & vUsu.Codigo
    
    ImprimeGeneral
    
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub

Private Function CargarTemporal() As Boolean
Dim Sql As String

    On Error GoTo eCargarTemporal

    CargarTemporal = False
    
    Sql = "delete from tmpfaclin where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    Sql = "insert into tmpfaclin (codusu, codigo, numserie, nomserie, numfac, fecha, cta, cliente, nif, imponible, impiva, total, retencion,"
    Sql = Sql & " recargo, tipoopera, tipoformapago) "
    Sql = Sql & " select distinct " & vUsu.Codigo & ",0, factcli.numserie, contadores.nomregis, factcli.numfactu, factcli.fecfactu, factcli.codmacta, "
    Sql = Sql & " factcli.nommacta, factcli.nifdatos, factcli.totbases, factcli.totivas, factcli.totfaccl, factcli.trefaccl, "
    Sql = Sql & " factcli.totrecargo, tipofpago.descformapago , aa.denominacion"
    Sql = Sql & " from " & tabla
    Sql = Sql & " where " & cadselect
    
    Conn.Execute Sql
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal Resumen", Err.Description
End Function

Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String
Dim I As Integer


    MontaSQL = False
    
            
    Sql = ""
    For I = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then
            Sql = Sql & Me.ListView1(1).ListItems(I).Text & ","
        End If
    Next I
    
    If Sql <> "" Then
        ' quitamos la ultima coma
        Sql = Mid(Sql, 1, Len(Sql) - 1)
        
        If Not AnyadirAFormula(cadselect, "factcli_totales.codigiva in (" & Sql & ")") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "{factcli_totales.codigiva} in [" & Sql & "]") Then Exit Function
    Else
        If Not AnyadirAFormula(cadselect, "factcli_totales.codigiva is null") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "isnull({factcli_totales.codigiva})") Then Exit Function
    End If
    
    
    If cadFormula <> "" Then cadFormula = "(" & cadFormula & ")"
    If cadselect <> "" Then cadselect = "(" & cadselect & ")"
    
    If Not CargarTemporal Then Exit Function
    
    cadFormula = "{tmpfaclin.codusu} = " & vUsu.Codigo
    
            
    MontaSQL = True
End Function

Private Sub txtAno_GotFocus(Index As Integer)
    ConseguirFoco txtAno(Index), 3
End Sub

Private Sub txtAno_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAno_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    txtAno(Index).Text = Trim(txtAno(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'A�o
            txtAno(Index).Text = Format(txtAno(Index).Text, "0000")
            
    End Select

End Sub


Private Sub txtCuota_GotFocus(Index As Integer)
    ConseguirFoco txtCuota(Index), 3
End Sub

Private Sub txtCuota_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCuota_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    txtCuota(Index).Text = Trim(txtCuota(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'Cuota
            If Not PonerFormatoDecimal(txtCuota(0), 1) Then
                txtCuota(0).Text = ""
            Else
                If ImporteFormateado(txtCuota(0).Text) < 0 Then
                    MsgBox "Importe positivo", vbExclamation
                    txtCuota(0).Text = ""
                    PonFoco txtCuota(0)
                End If
            End If
    End Select

End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    PonerFormatoFecha txtFecha(Index)
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtFecha(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub

Private Function DatosOK() As Boolean
    
    DatosOK = False
    
         
    If cmbPeriodo(0).ListIndex = -1 Or txtperiodo(0).Text = "" Then
        MsgBox "Campos per�odo no pueden estar vacios", vbExclamation
        Exit Function
    End If
    
    If cmbPeriodo(0).ListIndex = 0 Then
        For I = 0 To 1
            If Me.txtperiodo(I).Text = "" Then
                MsgBox "Campos per�odo no pueden estar vacios", vbExclamation
                Exit Function
            End If
        Next I
        
        If Val(txtperiodo(0).Text) > Val(txtperiodo(1).Text) Then
            MsgBox "Per�odo desde mayor que per�odo hasta.", vbExclamation
            Exit Function
        End If
        
        
        If vParam.periodos = 1 Then
            If Val(txtperiodo(0).Text) > 12 Or Val(txtperiodo(1).Text) > 12 Then
                MsgBox "Per�odo no puede ser superior a 12.", vbExclamation
                Exit Function
            End If
        Else
            'TRIMESTRAL
            If Val(txtperiodo(0).Text) > 4 Or Val(txtperiodo(1).Text) > 4 Then
                MsgBox "Per�odo no puede ser superior a 4.", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    If EmpresasSeleccionadas = 0 Then
        MsgBox "Seleccione una empresa", vbExclamation
        Exit Function
    End If

    'La empresa actual debe estar en la seleccion
    cad = ""
    For I = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then
            If Me.ListView1(1).ListItems(I).Text = vEmpresa.codempre Then cad = "OK"
        End If
    Next
    If cad = "" Then
        MsgBox "Debe generar la liquidacion desde una de las empresas seleccionadas", vbExclamation
        Exit Function
    End If


    ' comprobamos que las cuentas no esten a blancos
    If Me.chk1.Value = 1 Then
        If vParam.CtaHPAcreedor = "" Then
            MsgBox "Debe introducir una valor para Cuenta HP Acreedora. Revise.", vbExclamation
            Exit Function
        End If
        If vParam.CtaHPDeudor = "" Then
            MsgBox "Debe introducir una valor para Cuenta HP Deudora. Revise.", vbExclamation
            Exit Function
        End If
        
        
        If vParamT.Par_pen_apli = "" Then
            MsgBox "Falta configurar partidas pendientes de aplicacion", vbExclamation
            Exit Function
        End If
        
        'Para cada empresa seleccionada debe estar configurado la cuenta
        SqlLog = ""
        For I = 1 To ListView1(1).ListItems.Count
            If Me.ListView1(1).ListItems(I).Checked And ListView1(1).ListItems(I).Text <> vEmpresa.codempre Then
                For K = 1 To 2
                    J = CInt(ListView1(1).ListItems(I).Text)
                    RC = RecuperaValor("ctahpacreedor|ctahpdeudor|", CInt(K))
                    RC = DevuelveDesdeBD(RC, "ariconta" & J & ".parametros", "1", "1")
                    If RC <> "" Then RC = DevuelveDesdeBD("codmacta", "ariconta" & J & ".cuentas", "codmacta", RC, "T")
                    If RC = "" Then SqlLog = SqlLog & "Empresa: " & ListView1(1).ListItems(I).SubItems(1) & RecuperaValor(" ctahpacreedor| ctahpdeudor|", CInt(K)) & " no encontrado" & vbCrLf
                Next
            End If
        Next I
        
        If SqlLog <> "" Then
            MsgBox "Error configuracion" & vbCrLf & SqlLog, vbExclamation
            Exit Function
        End If
        
    End If



    DatosOK = True


End Function

Private Function EmpresasSeleccionadas() As Integer
Dim Sql As String
Dim I As Integer
Dim NSel As Integer

    NSel = 0
    For I = 1 To ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then NSel = NSel + 1
    Next I
    EmpresasSeleccionadas = NSel

End Function

Private Sub CargarListView(Index As Integer)
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

    ListView1(Index).ColumnHeaders.Add , , "C�digo", 600
    ListView1(Index).ColumnHeaders.Add , , "Descripci�n", 3200
    
    Sql = "SELECT codempre, nomempre, conta "
    Sql = Sql & " FROM usuarios.empresasariconta "
    
    If Not vParam.EsMultiseccion Then
        Sql = Sql & " where conta = " & DBSet(Conn.DefaultDatabase, "T")
    Else
        Sql = Sql & " where mid(conta,1,8) = 'ariconta'"
    End If
    Sql = Sql & " ORDER BY codempre "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        
        If vParam.EsMultiseccion Then
            If EsMultiseccion(DBLet(Rs!CONTA)) Then
                Set ItmX = ListView1(Index).ListItems.Add
                
                If DBLet(Rs!CONTA) = Conn.DefaultDatabase Then ItmX.Checked = True
                ItmX.Text = Rs.Fields(0).Value
                ItmX.SubItems(1) = Rs.Fields(1).Value
                ItmX.ToolTipText = Rs.Fields(1).Value
            End If
        Else
            Set ItmX = ListView1(Index).ListItems.Add
            
            ItmX.Checked = True
            ItmX.Text = Rs.Fields(0).Value
            ItmX.SubItems(1) = Rs.Fields(1).Value
            ItmX.ToolTipText = Rs.Fields(1).Value
        End If
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Empresas.", Err.Description
    End If
End Sub


Private Function GeneraLasLiquidaciones() As Boolean
    
    ' en tmpliquidaiva la columna cliente indica
    '                   0- Facturas clientes
    '                   1- Facturas clientes RECARGO EQUIVALENCIA
    '                   2- Facturas proveedores
    '                   3- Facturas Proveedores recargo equivalencia
    '                   4- Facturas Proveedores no deducible
    
    'Borramos los datos temporales
    Sql = "DELETE FROM tmpliquidaiva WHERE codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
    
    'Si alguna de las empresa esta inscriat devolucion IVA DUA, NO dejamos consolidar, ya cada factura hace un apunte en la liquidacion para
    ' esa dovlucion
    M2 = 0
    M1 = 0
    For I = 1 To Me.ListView1(1).ListItems.Count  'List2.ListCount - 1
        If Me.ListView1(1).ListItems(I).Checked Then
            M1 = M1 + 1 'Cuantas empresas
            Sql = "ariconta" & Me.ListView1(1).ListItems(I).Text & ".parametros"
            Sql = DevuelveDesdeBD("inscritoDeclarDUA", Sql, "1", "1")
            If Val(Sql) = 1 Then M2 = M2 + 1  'Cuantas llevan inscritoDeclarDUA
        End If
    Next I
    
    'Si hay mas de una empresa seleccionada
    If M1 > 1 Then
        'Si alguna lleva declaraDUA , no dejo continuar
        If M2 > 0 Then
            Sql = "Alguna empresa seleccionada esta inscrita a la devolucion IVA DUA."
            MsgBox Sql, vbExclamation
            Exit Function
        End If
    End If
    
    
    NumRegElim = 0
    M2 = 0
    M1 = 0
    'Para cada empresa
    'Para cada periodo
    For I = 1 To Me.ListView1(1).ListItems.Count  'List2.ListCount - 1
        If Me.ListView1(1).ListItems(I).Checked Then
            For CONT = CInt(txtperiodo(0).Text) To CInt(txtperiodo(1).Text)
                Label13.Caption = Mid(ListView1(1).ListItems(I).SubItems(1), 1, 20) & ".  " & CONT
                Label13.Refresh
                LiquidaIVA CByte(CONT), CInt(txtAno(0).Text), Me.ListView1(1).ListItems(I).Text, True   '(chkIVAdetallado.Value = 1)
            Next CONT
        End If
    Next I
    'Borraremos todos aquellos IVAS de Base imponible=0
    Sql = "DELETE From tmpliquidaiva WHERE codusu = " & vUsu.Codigo
    Sql = Sql & " AND bases = 0"
    Conn.Execute Sql
    
    
    GeneraLasLiquidaciones = True
End Function

Private Function LiquidaIVA(Periodo As Byte, Anyo As Integer, Empresa As Integer, Detallado As Boolean) As Boolean
Dim RIVA As Recordset
Dim TieneDeducibles As Boolean    'Para ahorrar tiempo
Dim HayRecargoEquivalencia As Boolean  'Para ahorrar tiempo tb
Dim IvasBienInversion As String 'Para saber si hemos comprado bien de inversion

    ' en tmpliquidaiva la columna cliente indica
    '                   0- Facturas clientes
    '                   1- Facturas clientes RECARGO EQUIVALENCIA
    '                   2- Facturas proveedores
    '                   3- Facturas Proveedores recargo equivalencia
    '                   4- Facturas Proveedores no deducible

    
    vCta = "ariconta" & Empresa
    
    'Para la cadena de busqueda
    LiquidaIVA = False
    

    '-----------------------------------------------
    '-----------------------------------------------
    '-----------------------------------------------
    'CLIENTES
    '-----------------------------------------------
    ' iva
    
    Sql = "insert into tmpliquidaiva(codusu,codmacta,bases,ivas,codempre,periodo,ano,cliente)"
    
    Sql = Sql & " select " & vUsu.Codigo & ",cuenta,sum(base),sum(iva), a, b," & Anyo & ",0    "
    Sql = Sql & " from ("
    
    Sql = Sql & " select " & vUsu.Codigo & ",tiposiva.cuentare cuenta,sum(baseimpo) base,sum(impoiva) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",0 "
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factcli_totales," & vCta & ".factcli"
    Sql = Sql & " where fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and tipodiva <> 3 " 'todos menos no deducible
    Sql = Sql & " and factcli_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
    Sql = Sql & " group by 1,2"
    Sql = Sql & " union "
    'isp
    Sql = Sql & " select " & vUsu.Codigo & ",tiposiva.cuentare cuenta,sum(baseimpo) base,sum(impoiva) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",0 "
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    'Sql = Sql & " and factpro.codopera = 4 " ' tipo de operacion inversion sujeto pasivo
    Sql = Sql & " and factpro.codopera in (1,4) " ' tipo de operacion inversion sujeto pasivo E intracom
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    Sql = Sql & " and factpro_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " group by 1,2"
    Sql = Sql & " union "
    'f isp
    
    
    ' recargo de equivalencia
    Sql = Sql & " select " & vUsu.Codigo & ",tiposiva.cuentarr cuenta,sum(baseimpo) base,sum(coalesce(imporec,0)) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",0 "
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factcli_totales," & vCta & ".factcli"
    Sql = Sql & " where fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and tipodiva <> 3 " 'todos menos no deducible
    Sql = Sql & " and factcli_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
    Sql = Sql & " group by 1,2"
    'isp
    Sql = Sql & " union "
    Sql = Sql & " select " & vUsu.Codigo & ",tiposiva.cuentarr cuenta,sum(baseimpo) base,sum(coalesce(imporec,0)) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",0 "
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and factpro.codopera = 4 " ' tipo de operacion inversion sujeto pasivo
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    Sql = Sql & " and factpro_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " group by 1,2"
    'f isp
        
    
    Sql = Sql & " ) aaaaa "
    
    Sql = Sql & " group by 1,2"
                    
    Conn.Execute Sql
    
    
    
    '-----------------------------------------------
    '-----------------------------------------------
    '-----------------------------------------------
    '           PROVEEDORES
    '-----------------------------------------------
    Sql = "insert into tmpliquidaiva(codusu,codmacta,bases,ivas,codempre,periodo,ano,cliente) "
    
    Sql = Sql & " select " & vUsu.Codigo & ",cuenta,sum(base),sum(iva), a, b," & Anyo & ",cliente    "
    Sql = Sql & " from ("
    Sql = Sql & " select " & vUsu.Codigo & ",tiposiva.cuentaso cuenta,sum(baseimpo) base,sum(impoiva) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",1 cliente"
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and tipodiva <> 3 " ' todos menos no deducible
    Sql = Sql & " and factpro_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    Sql = Sql & " group by 1,2"
    Sql = Sql & " union "
    
    ' recargo de equivalencia
    Sql = Sql & " select " & vUsu.Codigo & ",tiposiva.cuentasr cuenta,sum(baseimpo) base,sum(imporec) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",1 cliente"
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and tipodiva <> 3 " ' todos menos no deducible
    Sql = Sql & " and factpro_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    Sql = Sql & " group by 1,2"
   
    
    
    ' soportado no deducible
    'NO ENTRAN EN LA LIQUIDACION   16 septiembre 2019
    If False Then
        Sql = Sql & " union "
        
        Sql = Sql & " select " & vUsu.Codigo & ",tiposiva.cuentasn cuenta,sum(baseimpo) base,sum(impoiva) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",1 cliente"
        Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
        Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
        Sql = Sql & " and tipodiva = 3 " ' los no deducibles
        Sql = Sql & " and factpro_totales.codigiva = tiposiva.codigiva "
        Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
        Sql = Sql & " group by 1,2"
        
    End If
    
    
    Sql = Sql & " ) aaaaa "
    
    Sql = Sql & " group by 1,2"
                    
    Conn.Execute Sql
    
    
    
    'Si tiene cuotas a compensar
    
    If Me.txtCuota(0).Text <> "" Then
        If Empresa = vEmpresa.codempre Then
            'Es aqui donde hay que insertar la compensacion
            Sql = vParam.CtaHPDeudor
            If Sql = "" Then Sql = "COMPENSA"
            Sql = " VALUES  (" & vUsu.Codigo & "," & DBSet(Sql, "T") & ",1," & DBSet(txtCuota(0).Text, "N")
            Sql = Sql & "," & Empresa & "," & Periodo & "," & Anyo & ",1 )"
            Sql = "insert into tmpliquidaiva(codusu,codmacta,bases,ivas,codempre,periodo,ano,cliente)" & Sql
            Conn.Execute Sql
            
        End If
    End If
    
    
End Function






