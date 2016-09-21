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
      Height          =   5415
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6915
      Begin VB.Frame FramePeriodo 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   90
         TabIndex        =   19
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
            TabIndex        =   1
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
            TabIndex        =   2
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
            TabIndex        =   21
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
            TabIndex        =   20
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
         TabIndex        =   0
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
         TabIndex        =   3
         Tag             =   "imgConcepto"
         Top             =   930
         Width           =   765
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   3630
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "Período"
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
         TabIndex        =   9
         Top             =   570
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Año"
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
         TabIndex        =   8
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
      TabIndex        =   12
      Top             =   0
      Width           =   4455
      Begin VB.CheckBox chk1 
         Caption         =   "Realizar apunte contable de cancelación"
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
         TabIndex        =   22
         Top             =   4650
         Width           =   4335
      End
      Begin VB.Frame FrameSeccion 
         BorderStyle     =   0  'None
         Height          =   3525
         Left            =   180
         TabIndex        =   16
         Top             =   1020
         Width           =   4185
         Begin MSComctlLib.ListView ListView1 
            Height          =   2880
            Index           =   1
            Left            =   60
            TabIndex        =   17
            Top             =   510
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   5080
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
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
            TabIndex        =   18
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
         TabIndex        =   4
         Top             =   570
         Width           =   1485
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3840
         TabIndex        =   13
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
         TabIndex        =   14
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
      TabIndex        =   6
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
      Left            =   8790
      TabIndex        =   5
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
      TabIndex        =   15
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

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
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

Private SQL As String
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
    FramePeriodo.Enabled = False
    FramePeriodo.Visible = False
    
    CargarFechas
    
End Sub

Private Sub cmdAccion_Click(Index As Integer)
Dim Pregunta As Boolean
Dim B As Boolean

    If Not DatosOK Then Exit Sub
    
    
'++
    'AHora generaremos la liquidacion para todos los periodos k abarque la seleecion
    Screen.MousePointer = vbHourglass
    'Guardamos el valor del chk del IVA
'--
'    ModeloIva False
    Label13.Caption = "Elimina datos anteriores"
    Label13.Visible = True
    Label13.Refresh
    If GeneraLasLiquidaciones Then
        Label13.Caption = ""
        Label13.Refresh
        espera 0.5
        'Periodos
        SQL = ""
        For I = 0 To 1
            SQL = SQL & txtperiodo(I).Text & "|"
        Next I
        SQL = SQL & txtAno(0).Text & "|"
        I = 1
        
        Periodo = SQL & I & "|"
        
    
        SqlLog = "Periodo : " & txtAno(0) & " / " & Me.cmbPeriodo(0).Text & vbCrLf
        SqlLog = SqlLog & "Empresas : "
        
    
        'Empresas para consolidado
        Pregunta = True
        SQL = ""
        If EmpresasSeleccionadas = 1 Then
            B = False
            For I = 1 To Me.ListView1(1).ListItems.Count
            
                SqlLog = SqlLog & "ariconta" & Me.ListView1(1).ListItems(I).Text & vbCrLf
                
                If ListView1(1).ListItems(I).Checked Then
                    
                    NumConta = Me.ListView1(1).ListItems(I).Text
                    
                    ImprimirAsientoContable
                    
                    If chk1.Value Then
                        
                        If HayRegParaInforme("tmpconext", "codusu=" & vUsu.Codigo) Then
    
                            Set frmMens = New frmMensajes
                            
                            frmMens.Opcion = 29
                            frmMens.Show vbModal
                            
                            Set frmMens = Nothing
        
                        End If

                        If CadenaDesdeOtroForm = "OK" Then
                            If RealizarAsientoContable Then
                                B = True
                                Exit For
                            End If
                        End If
                    Else
                    
                        If MsgBox("¿ Desea actualizar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                            B = ActualizarLiquidacion(False)
                            If B Then
                                B = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next I
        Else
            'Mas de una empresa
            SQL = "'Empresas seleccionadas:' + Chr(13) "
            B = False
            For I = 1 To Me.ListView1(1).ListItems.Count
                            
                NumConta = Me.ListView1(1).ListItems(I).Text
         
                SqlLog = SqlLog & " ariconta" & NumConta & vbCrLf
                
         
                ImprimirAsientoContable
         
                If Pregunta Then
                    If MsgBox("¿ Desea actualizar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                        Pregunta = False
                        B = ActualizarLiquidacion(False)
                        If B Then
                            Pregunta = False
                        Else
                            Exit For
                        End If
                    
                    Else
                        Exit For
                    End If
                Else
                    B = ActualizarLiquidacion(False)
                End If
            Next I
        End If
        
        If B Then
            If SqlLog <> "" Then vLog.Insertar 15, vUsu, SqlLog
        
        
            MsgBox "Proceso realizado correctamente.", vbExclamation
            Unload Me
        End If


    
    End If
    Label13.Visible = False
    Me.Refresh
    Screen.MousePointer = vbDefault


    
    
End Sub

Private Function ActualizarLiquidacion(DentroDeTrans As Boolean, Optional NumAsiento As Long) As Boolean
Dim SQL As String
Dim I As Integer
    On Error GoTo eActualizarLiquidacion

    If Not DentroDeTrans Then Conn.BeginTrans

    ActualizarLiquidacion = False
    
    ' actualizamos los parametros
    SQL = "update ariconta" & NumConta & ".parametros set anofactu = " & DBSet(txtAno(0).Text, "N")
    I = txtperiodo(0)
    SQL = SQL & ", perfactu = " & DBSet(I, "N")
    Conn.Execute SQL


    vParam.anofactu = txtAno(0).Text
    vParam.perfactu = I

    If vParam.periodos = 0 Then
        I = I + 12
    End If

    SQL = "insert into ariconta" & NumConta & ".liqiva (anoliqui,periodo,escomplem,importe,numdiari,numasien,fechaent) values ("
    SQL = SQL & DBSet(txtAno(0).Text, "N") & "," & DBSet(I, "N") & ",0," & DBSet(ImpLiqui, "N") & "," & DBSet(vParam.numdia303, "N") & "," & DBSet(NumAsiento, "N") & "," & DBSet(txtFecha(2).Text, "F") & ")"
    Conn.Execute SQL
    
    If Not DentroDeTrans Then Conn.CommitTrans
    
    ActualizarLiquidacion = True
    Exit Function


eActualizarLiquidacion:
    If Not DentroDeTrans Then Conn.RollbackTrans
    MuestraError Err.Description, "Actualizar Liquidación", Err.Description
End Function

Private Function RealizarAsientoContable() As Boolean
Dim Mc As Contadores
Dim B As Boolean
Dim Numdocum As String
Dim Ampconce As String
Dim MaxPos As Long
Dim NomConce As String
Dim NumAsien As Long

    On Error GoTo eRealizarAsientoContable
    
    RealizarAsientoContable = False
    
    Set Mc = New Contadores
    
    Conn.BeginTrans
    
    I = FechaCorrecta2(CDate(txtFecha(2).Text))
    If Mc.ConseguirContador("0", (I = 0), False) = 0 Then
        NumAsien = Mc.Contador
    End If
    
    ' insertamos en cabecera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion ) SELECT " & vParam.numdia303 & "," & DBSet(txtFecha(2).Text, "F") & "," & DBSet(NumAsien, "N")
    SQL = SQL & ",'Liquidación de " & Me.cmbPeriodo(0).Text & " de " & txtAno(0).Text & "'," & DBSet(Now, "F") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Liquidación'"
    SQL = SQL & " from parametros "
    Conn.Execute SQL
    
    
    NomConce = DevuelveValor("select nomconce from conceptos where codconce = " & vParam.conce303)
    Numdocum = "LIQ." & txtAno(0).Text & "-" & txtperiodo(1).Text
    
    If vParam.periodos = 0 Then
        Ampconce = NomConce & " Liq.303 " & txtperiodo(0).Text & "T"
    Else
        Ampconce = NomConce & " Liq.303 " & cmbPeriodo(0).Text
    End If
    
    MaxPos = DevuelveValor("select max(pos) from tmpconext where codusu = " & DBSet(vUsu.Codigo, "N"))
    
    ' insertamos en lineas
    SQL = "INSERT INTO hlinapu (numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,ctacontr) SELECT " & vParam.numdia303 & "," & DBSet(txtFecha(2).Text, "F") & "," & DBSet(NumAsien, "N")
    SQL = SQL & ",pos, cta," & DBSet(Numdocum, "T") & "," & DBSet(vParam.conce303, "N") & "," & DBSet(Ampconce, "T") & ",if(timported=0,null,timported), if(timporteh=0,null,timporteh), "
    If ImpLiqui > 0 Then
        SQL = SQL & "if(pos <> " & DBSet(MaxPos, "N") & "," & DBSet(vParam.CtaHPAcreedor, "T") & ",null) "
    Else
        SQL = SQL & "if(pos <> " & DBSet(MaxPos, "N") & "," & DBSet(vParam.CtaHPDeudor, "T") & ",null) "
    End If
    
    SQL = SQL & " from tmpconext where codusu =  " & vUsu.Codigo
    SQL = SQL & " order by pos "
    Conn.Execute SQL
    
    
    
    SqlLog = SqlLog & "Asiento contable: " & DBSet(NumAsien, "N") & " - " & txtFecha(2).Text & " - " & vParam.numdia303 & vbCrLf
    
    
    B = ActualizarLiquidacion(True, NumAsien)
    
    If B Then
        Conn.CommitTrans
        RealizarAsientoContable = True
        Exit Function
    End If
    
eRealizarAsientoContable:
    Conn.RollbackTrans
    MuestraError Err.Description, "Realizar Asiento contable", Err.Description
End Function



Private Sub ImprimirAsientoContable()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim SqlInsert As String
Dim SqlInsert2 As String
Dim SqlValues As String
Dim SqlValues2 As String
Dim Importe As Currency
Dim vDebe As Currency
Dim vHaber As Currency
Dim I As Long

    On Error GoTo eImprimirAsientoContable

    SQL = "delete from ariconta" & NumConta & ".tmpconext where codusu= " & vUsu.Codigo
    Conn.Execute SQL
    
    ' para visualizar los saldos
    SQL = "delete from ariconta" & NumConta & ".tmpconextcab where codusu= " & vUsu.Codigo
    Conn.Execute SQL
    
    ' codigo = 0 debe
    '          1 haber
    
    SqlInsert = "insert into ariconta" & NumConta & ".tmpconext(codusu,pos,cta,timported,timporteh) values "
    SqlInsert2 = "insert into ariconta" & NumConta & ".tmpconextcab(codusu,cta,acumtotT) values "
    
    SQL = "select cliente, codmacta, sum(coalesce(ivas,0)) importe from ariconta" & NumConta & ".tmpliquidaiva where codusu = " & vUsu.Codigo
    SQL = SQL & " group by 1,2 "
    SQL = SQL & " having sum(coalesce(ivas,0)) <> 0"
    SQL = SQL & " order by 1,2 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SqlValues = ""
    I = 0
    While Not RS.EOF
        I = I + 1
    
        Importe = DBLet(RS!Importe, "N")
    
        SqlValues = SqlValues & "(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(I, "N") & "," & DBSet(RS!codmacta, "T") & ","
    
        If DBLet(RS!Cliente, "N") = 0 Then ' clientes
            If Importe >= 0 Then
                SqlValues = SqlValues & DBSet(Importe, "N") & "," & "0)," ' clientes positivo al debe
            Else
                SqlValues = SqlValues & "0," & DBSet(Importe * (-1), "N") & ")," ' clientes negativo al haber
            End If
        Else 'proveedores
            If Importe >= 0 Then
                SqlValues = SqlValues & "0," & DBSet(Importe, "N") & ")," ' clientes positivo al haber
            Else
                SqlValues = SqlValues & DBSet(Importe * (-1), "N") & "," & "0)," ' clientes negativo al debe
            End If
        End If
    
        ' cargamos cual es el saldo entre la fecha de inicio de ejercicio y la fecha de liquidacion
        SQL = "select abs(sum(coalesce(timported,0)) - sum(coalesce(timporteh,0))) from ariconta" & NumConta & ".hlinapu where codmacta =  " & DBSet(RS!codmacta, "T")
        SQL = SQL & " and fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vFecha2, "F")
    
        SqlValues2 = SqlValues2 & "(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(RS!codmacta, "T") & "," & DBSet(DevuelveValor(SQL), "N") & "),"
        
        RS.MoveNext
    Wend
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        
        Conn.Execute SqlInsert & SqlValues
        
        ' los saldos
        SqlValues2 = Mid(SqlValues2, 1, Len(SqlValues2) - 1)
        
        Conn.Execute SqlInsert2 & SqlValues2
        
    
        SQL = "select sum(timported) from ariconta" & NumConta & ".tmpconext where codusu = " & vUsu.Codigo
        vDebe = DevuelveValor(SQL)
        
        SQL = "select sum(timporteh) from ariconta" & NumConta & ".tmpconext where codusu = " & vUsu.Codigo
        vHaber = DevuelveValor(SQL)
    
        SqlValues = ""
        I = I + 1
        If vDebe - vHaber > 0 Then
            SqlValues = "(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(I, "N") & "," & DBSet(vParam.CtaHPAcreedor, "T") & ",0," & DBSet(vDebe - vHaber, "N") & ")"
        Else
            If vDebe - vHaber < 0 Then
                SqlValues = "(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(I, "N") & "," & DBSet(vParam.CtaHPDeudor, "T") & "," & DBSet(vHaber - vDebe, "N") & ",0)"
            End If
        End If
        'Apunte de la diferencia debe - haber
        Conn.Execute SqlInsert & SqlValues
    
        ImpLiqui = vDebe - vHaber
    
    
    End If

    Set RS = Nothing
    
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
    Me.Icon = frmPpal.Icon
        
    'Otras opciones
    Me.Caption = "Liquidación de Iva"

     
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
     
    
    CargarListView 1
    
    PonerPeriodoPresentacion303
     
     
    FrameSeccion.Enabled = vParam.EsMultiseccion
    
    FramePeriodo.Enabled = (Me.cmbPeriodo(0).ListIndex = 0)
    FramePeriodo.Visible = (Me.cmbPeriodo(0).ListIndex = 0)
    
    CargarFechas
    
    
    
    txtFecha(2).Text = Format(vFecha2, "dd/mm/yyyy")
     
    
End Sub

Private Sub CargarFechas()
    
    If vParam.periodos = 1 Then
        'Esamos en mensual
        If Me.cmbPeriodo(0).ListIndex > 12 Then
            MsgBox "Error en el periodo a tratar.", vbExclamation
            Exit Sub
        End If
        vFecha1 = CDate("01/" & Me.cmbPeriodo(0).ListIndex & "/" & Me.txtAno(0))
        M1 = DiasMes(Me.cmbPeriodo(0).ListIndex, Me.txtAno(0))
        vFecha2 = CDate(M1 & "/" & Me.cmbPeriodo(0).ListIndex & "/" & Me.txtAno(0))
        
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
                CadenaDesdeOtroForm = "º"
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
    
    txtAno(0).Text = vParam.anofactu
    I = vParam.perfactu + 1
    If vParam.periodos = 0 Then
        NumRegElim = 4
    Else
        NumRegElim = 12
    End If
        
    If I > NumRegElim Then
            I = 1
            txtAno(0).Text = vParam.anofactu + 1
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
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
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
Dim RS As ADODB.Recordset

    TotalClien = 0

    'En devuelveimporte
    ' Tipo 0:   11 enteros y 2 decimales
    ' Tipo 1:   2 ente y 2 decimales
    ' Tipo 2:   1 entero y 2 decimales
    ' tipo 3:   3 enetero y dos decimales

    
    SQL = "select iva,  bases, ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 0 "
    SQL = SQL & " order by 1 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
        I = I + 1
        DevuelveImporte DBLet(RS!Bases, "N"), 0
        DevuelveImporte DBLet(RS!iva, "N"), 3
        DevuelveImporte DBLet(RS!ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(RS!ivas, "N")
        
        RS.MoveNext
    Wend
    
    'por si hay menos de 3 porcentajes de iva hay que rellenarlos a ceros
    For J = I + 1 To 3
        DevuelveImporte 0, 0
        DevuelveImporte 0, 3
        DevuelveImporte 0, 0
    Next J
    
    Set RS = Nothing
    
    'Adquisiciones intra
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 10 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    HayReg = False
    
    While Not RS.EOF
        HayReg = True
        DevuelveImporte DBLet(RS!Bases, "N"), 0
        DevuelveImporte DBLet(RS!ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(RS!ivas, "N")
        
        RS.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set RS = Nothing
    
    ' Inversion de sujeto pasivo
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 12 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not RS.EOF
        HayReg = True
        DevuelveImporte DBLet(RS!Bases, "N"), 0
        DevuelveImporte DBLet(RS!ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(RS!ivas, "N")
        
        RS.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set RS = Nothing
    
    'modificacion bases y cuotas (no tenemos)
    DevuelveImporte 0, 0
    DevuelveImporte 0, 0
    
    
    'Los recargos
    SQL = "select iva,  bases, ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 1 "
    SQL = SQL & " order by 1 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
        I = I + 1
        DevuelveImporte DBLet(RS!Bases, "N"), 0
        DevuelveImporte DBLet(RS!iva, "N"), 3
        DevuelveImporte DBLet(RS!ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(RS!ivas, "N")
        
        RS.MoveNext
    Wend
    
    'por si hay menos de 3 porcentajes de iva hay que rellenarlos a ceros
    For J = I + 1 To 3
        DevuelveImporte 0, 0
        DevuelveImporte 0, 3
        DevuelveImporte 0, 0
    Next J
    
    Set RS = Nothing
    
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
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 2 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not RS.EOF
        HayReg = True
        DevuelveImporte DBLet(RS!Bases, "N"), 0
        DevuelveImporte DBLet(RS!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(RS!ivas, "N")
        
        RS.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set RS = Nothing
    
    'operaciones interiores BIENES INVERSION
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 30 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not RS.EOF
        HayReg = True
        DevuelveImporte DBLet(RS!Bases, "N"), 0
        DevuelveImporte DBLet(RS!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(RS!ivas, "N")
        
        RS.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set RS = Nothing
    
    'importaciones
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 32 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not RS.EOF
        HayReg = True
        DevuelveImporte DBLet(RS!Bases, "N"), 0
        DevuelveImporte DBLet(RS!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(RS!ivas, "N")
        
        RS.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set RS = Nothing
    
    'importaciones BIEN INVERSION
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 34 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not RS.EOF
        HayReg = True
        DevuelveImporte DBLet(RS!Bases, "N"), 0
        DevuelveImporte DBLet(RS!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(RS!ivas, "N")
        
        RS.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set RS = Nothing
    
    
    
    'adqisiciones intracom
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 36 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not RS.EOF
        HayReg = True
        DevuelveImporte DBLet(RS!Bases, "N"), 0
        DevuelveImporte DBLet(RS!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(RS!ivas, "N")
        
        RS.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set RS = Nothing
    
    'adqisiciones intracom BIEN INVERSION
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 38 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not RS.EOF
        HayReg = True
        DevuelveImporte DBLet(RS!Bases, "N"), 0
        DevuelveImporte DBLet(RS!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(RS!ivas, "N")
        
        RS.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set RS = Nothing

    ' rectificacion de deducciones tampoco tenemos
    DevuelveImporte 0, 0
    DevuelveImporte 0, 0

    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 42 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not RS.EOF
        HayReg = True
        DevuelveImporte DBLet(RS!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(RS!ivas, "N")
        
        RS.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
    End If
    
    Set RS = Nothing
    

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
    
    SQL = ""
    If EmpresasSeleccionadas = 1 Then
        For I = 1 To Me.ListView1(1).ListItems.Count
            If ListView1(1).ListItems(I).Checked Then
                If Me.ListView1(1).ListItems(I).Text <> vEmpresa.nomempre Then SQL = Me.ListView1(1).ListItems(I).SubItems(1)
            End If
        Next I
    Else
        'Mas de una empresa
        SQL = "'Empresas seleccionadas:' + Chr(13) "
        For I = 1 To Me.ListView1(1).ListItems.Count
            SQL = SQL & " + '        " & Me.ListView1(1).ListItems(I).Text & "' + Chr(13)"
        Next I
    End If
    
    cadParam = cadParam & "empresas = """ & SQL & """|"
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
Dim SQL As String

    On Error GoTo eCargarTemporal

    CargarTemporal = False
    
    SQL = "delete from tmpfaclin where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "insert into tmpfaclin (codusu, codigo, numserie, nomserie, numfac, fecha, cta, cliente, nif, imponible, impiva, total, retencion,"
    SQL = SQL & " recargo, tipoopera, tipoformapago) "
    SQL = SQL & " select distinct " & vUsu.Codigo & ",0, factcli.numserie, contadores.nomregis, factcli.numfactu, factcli.fecfactu, factcli.codmacta, "
    SQL = SQL & " factcli.nommacta, factcli.nifdatos, factcli.totbases, factcli.totivas, factcli.totfaccl, factcli.trefaccl, "
    SQL = SQL & " factcli.totrecargo, tipofpago.descformapago , aa.denominacion"
    SQL = SQL & " from " & tabla
    SQL = SQL & " where " & cadselect
    
    Conn.Execute SQL
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal Resumen", Err.Description
End Function

Private Function MontaSQL() As Boolean
Dim SQL As String
Dim SQL2 As String
Dim RC As String
Dim RC2 As String
Dim I As Integer


    MontaSQL = False
    
            
    SQL = ""
    For I = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then
            SQL = SQL & Me.ListView1(1).ListItems(I).Text & ","
        End If
    Next I
    
    If SQL <> "" Then
        ' quitamos la ultima coma
        SQL = Mid(SQL, 1, Len(SQL) - 1)
        
        If Not AnyadirAFormula(cadselect, "factcli_totales.codigiva in (" & SQL & ")") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "{factcli_totales.codigiva} in [" & SQL & "]") Then Exit Function
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
        Case 0 'Año
            txtAno(Index).Text = Format(txtAno(Index).Text, "0000")
            
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
        MsgBox "Campos período no pueden estar vacios", vbExclamation
        Exit Function
    End If
    
    If cmbPeriodo(0).ListIndex = 0 Then
        For I = 0 To 1
            If Me.txtperiodo(I).Text = "" Then
                MsgBox "Campos período no pueden estar vacios", vbExclamation
                Exit Function
            End If
        Next I
        
        If Val(txtperiodo(0).Text) > Val(txtperiodo(1).Text) Then
            MsgBox "Período desde mayor que período hasta.", vbExclamation
            Exit Function
        End If
        
        
        If vParam.periodos = 1 Then
            If Val(txtperiodo(0).Text) > 12 Or Val(txtperiodo(1).Text) > 12 Then
                MsgBox "Período no puede ser superior a 12.", vbExclamation
                Exit Function
            End If
        Else
            'TRIMESTRAL
            If Val(txtperiodo(0).Text) > 4 Or Val(txtperiodo(1).Text) > 4 Then
                MsgBox "Período no puede ser superior a 4.", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    If EmpresasSeleccionadas = 0 Then
        MsgBox "Seleccione una empresa", vbExclamation
        Exit Function
    End If

    ' comprobamos que las cuentas no esten a blancos
    If vParam.CtaHPAcreedor = "" Then
        MsgBox "Debe introducir una valor para Cuenta HP Acreedora. Revise.", vbExclamation
        Exit Function
    End If
    If vParam.CtaHPDeudor = "" Then
        MsgBox "Debe introducir una valor para Cuenta HP Deudora. Revise.", vbExclamation
        Exit Function
    End If
    



    DatosOK = True


End Function

Private Function EmpresasSeleccionadas() As Integer
Dim SQL As String
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
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

    ListView1(Index).ColumnHeaders.Add , , "Código", 600
    ListView1(Index).ColumnHeaders.Add , , "Descripción", 3200
    
    SQL = "SELECT codempre, nomempre, conta "
    SQL = SQL & " FROM usuarios.empresasariconta "
    
    If Not vParam.EsMultiseccion Then
        SQL = SQL & " where conta = " & DBSet(Conn.DefaultDatabase, "T")
    Else
        SQL = SQL & " where mid(conta,1,8) = 'ariconta'"
    End If
    SQL = SQL & " ORDER BY codempre "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        
        If vParam.EsMultiseccion Then
            If EsMultiseccion(DBLet(RS!CONTA)) Then
                Set ItmX = ListView1(Index).ListItems.Add
                
                If DBLet(RS!CONTA) = Conn.DefaultDatabase Then ItmX.Checked = True
                ItmX.Text = RS.Fields(0).Value
                ItmX.SubItems(1) = RS.Fields(1).Value
            End If
        Else
            Set ItmX = ListView1(Index).ListItems.Add
            
            ItmX.Checked = True
            ItmX.Text = RS.Fields(0).Value
            ItmX.SubItems(1) = RS.Fields(1).Value
        End If
        
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing

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
    SQL = "DELETE FROM tmpliquidaiva WHERE codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    
    NumRegElim = 0
    'Para cada empresa
    'Para cada periodo
    For I = 1 To Me.ListView1(1).ListItems.Count  'List2.ListCount - 1
        If Me.ListView1(1).ListItems(I).Checked Then
            For CONT = CInt(txtperiodo(0).Text) To CInt(txtperiodo(1).Text)
                Label13.Caption = Mid(ListView1(1).ListItems(I).SubItems(1), 1, 20) & ".  " & CONT
                Label13.Refresh
                LiquidaIVA CByte(CONT), CInt(txtAno(0).Text), Me.ListView1(1).ListItems(I).Text, True  '(chkIVAdetallado.Value = 1)
            Next CONT
        End If
    Next I
    'Borraremos todos aquellos IVAS de Base imponible=0
    SQL = "DELETE From tmpliquidaiva WHERE codusu = " & vUsu.Codigo
    SQL = SQL & " AND bases = 0"
    Conn.Execute SQL
    
    
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
    
    SQL = "insert into tmpliquidaiva(codusu,codmacta,bases,ivas,codempre,periodo,ano,cliente)"
    
    SQL = SQL & " select " & vUsu.Codigo & ",cuenta,sum(base),sum(iva), a, b," & Anyo & ",0    "
    SQL = SQL & " from ("
    
    SQL = SQL & " select " & vUsu.Codigo & ",tiposiva.cuentare cuenta,sum(baseimpo) base,sum(impoiva) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",0 "
    SQL = SQL & " from " & vCta & ".tiposiva," & vCta & ".factcli_totales," & vCta & ".factcli"
    SQL = SQL & " where fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    SQL = SQL & " and tipodiva <> 3 " 'todos menos no deducible
    SQL = SQL & " and factcli_totales.codigiva = tiposiva.codigiva "
    SQL = SQL & " and factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
    SQL = SQL & " group by 1,2"
    SQL = SQL & " union "
    'isp
    SQL = SQL & " select " & vUsu.Codigo & ",tiposiva.cuentare cuenta,sum(baseimpo) base,sum(impoiva) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",0 "
    SQL = SQL & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    SQL = SQL & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    SQL = SQL & " and factpro.codopera = 4 " ' tipo de operacion inversion sujeto pasivo
    SQL = SQL & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    SQL = SQL & " and factpro_totales.codigiva = tiposiva.codigiva "
    SQL = SQL & " group by 1,2"
    SQL = SQL & " union "
    'f isp
    
    
    ' recargo de equivalencia
    SQL = SQL & " select " & vUsu.Codigo & ",tiposiva.cuentarr cuenta,sum(baseimpo) base,sum(coalesce(imporec,0)) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",0 "
    SQL = SQL & " from " & vCta & ".tiposiva," & vCta & ".factcli_totales," & vCta & ".factcli"
    SQL = SQL & " where fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    SQL = SQL & " and tipodiva <> 3 " 'todos menos no deducible
    SQL = SQL & " and factcli_totales.codigiva = tiposiva.codigiva "
    SQL = SQL & " and factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
    SQL = SQL & " group by 1,2"
    'isp
    SQL = SQL & " union "
    SQL = SQL & " select " & vUsu.Codigo & ",tiposiva.cuentarr cuenta,sum(baseimpo) base,sum(coalesce(imporec,0)) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",0 "
    SQL = SQL & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    SQL = SQL & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    SQL = SQL & " and factpro.codopera = 4 " ' tipo de operacion inversion sujeto pasivo
    SQL = SQL & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    SQL = SQL & " and factpro_totales.codigiva = tiposiva.codigiva "
    SQL = SQL & " group by 1,2"
    'f isp
        
    
    SQL = SQL & " ) aaaaa "
    
    SQL = SQL & " group by 1,2"
                    
    Conn.Execute SQL
    
    
    
    '-----------------------------------------------
    '-----------------------------------------------
    '-----------------------------------------------
    '           PROVEEDORES
    '-----------------------------------------------
    SQL = "insert into tmpliquidaiva(codusu,codmacta,bases,ivas,codempre,periodo,ano,cliente) "
    
    SQL = SQL & " select " & vUsu.Codigo & ",cuenta,sum(base),sum(iva), a, b," & Anyo & ",cliente    "
    SQL = SQL & " from ("
    SQL = SQL & " select " & vUsu.Codigo & ",tiposiva.cuentaso cuenta,sum(baseimpo) base,sum(impoiva) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",1 cliente"
    SQL = SQL & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    SQL = SQL & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    SQL = SQL & " and tipodiva <> 3 " ' todos menos no deducible
    SQL = SQL & " and factpro_totales.codigiva = tiposiva.codigiva "
    SQL = SQL & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    SQL = SQL & " group by 1,2"
    SQL = SQL & " union "
    ' recargo de equivalencia
    SQL = SQL & " select " & vUsu.Codigo & ",tiposiva.cuentasr cuenta,sum(baseimpo) base,sum(imporec) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",1 cliente"
    SQL = SQL & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    SQL = SQL & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    SQL = SQL & " and tipodiva <> 3 " ' todos menos no deducible
    SQL = SQL & " and factpro_totales.codigiva = tiposiva.codigiva "
    SQL = SQL & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    SQL = SQL & " group by 1,2"
    SQL = SQL & " union "
    ' soportado no deducible
    SQL = SQL & " select " & vUsu.Codigo & ",tiposiva.cuentasn cuenta,sum(baseimpo) base,sum(impoiva) iva," & Empresa & " a," & Periodo & " b," & Anyo & ",1 cliente"
    SQL = SQL & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    SQL = SQL & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    SQL = SQL & " and tipodiva = 3 " ' los no deducibles
    SQL = SQL & " and factpro_totales.codigiva = tiposiva.codigiva "
    SQL = SQL & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    SQL = SQL & " group by 1,2"
    SQL = SQL & " ) aaaaa "
    
    SQL = SQL & " group by 1,2"
                    
    Conn.Execute SQL
    
    
    
End Function






