VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form frmEditorMenus2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor de Men�s"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   11790
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditorMenus2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   8550
      Left            =   90
      TabIndex        =   0
      Top             =   615
      Width           =   11565
      _Version        =   1114114
      _ExtentX        =   20399
      _ExtentY        =   15081
      _StockProps     =   64
      MultipleSelection=   0   'False
      FreezeColumnsAbs=   0   'False
      MultiSelectionMode=   -1  'True
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdCopiarOtraEmpresa 
         Height          =   375
         Left            =   240
         Picture         =   "frmEditorMenus2.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Copiar a otra empresa"
         Top             =   0
         Width           =   375
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   3600
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnCalculomenus_usuariosProd 
         Caption         =   "&C�lculo menus_usuarios Prod."
         Shortcut        =   ^C
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnExportacion 
         Caption         =   "Exportar Excel"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnFiltro 
      Caption         =   "Filtro"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnFiltro1 
         Caption         =   "A�o actual"
      End
      Begin VB.Menu mnFiltro2 
         Caption         =   "A�o actual y anterior"
      End
      Begin VB.Menu mnBarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnFiltro3 
         Caption         =   "Sin Filtro"
      End
   End
End
Attribute VB_Name = "frmEditorMenus2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MONICA  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funci� BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funci� BotonBuscar() canviar el nom de la clau primaria
' 5. En la funci� BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funci� PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar alg�n) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada bot� per a que corresponguen
' 9. En la funci� CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar adem�s els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funci� DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funci� SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es fa�a refer�ncia a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la b�squeda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edici� del camp
'   3.-  Inserci� de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Integer


Dim kaplicacion As String
' utilizado para buscar por checks
Private BuscaChekc As String


'Constants used to identify columns, this will be the column ItemIndex
Const COLUMN_IMPORTANCE = 0
Const COLUMN_ICON = 1
Const COLUMN_ATTACHMENT = 2
Const COLUMN_FROM = 3

Const COLUMN_NOMBRE = 4
Const COLUMN_VER = 5
Const COLUMN_CREARELIMINAR = 6
Const COLUMN_MODIFICAR = 7
Const COLUMN_IMPRIMIR = 8
Const COLUMN_ESPECIAL = 9

Const COLUMN_RECEIVED = 10
Const COLUMN_CONVERSATION = 11
Const COLUMN_CONTACTS = 12
Const COLUMN_MESSAGE = 13
Const COLUMN_CC = 14
Const COLUMN_CATEGORIES = 15
Const COLUMN_AUTOFORWARD = 16
Const COLUMN_DO_NOT_AUTOARCH = 17
Const COLUMN_DUE_BY = 18

'Array used to exctract icons from a bitmap (bitmap in Icons folder)
Dim iconArray(0 To 11) As Long

'Constants used to identify icons used in the ReportControl
Const COLUMN_MAIL_ICON = 1
Const COLUMN_IMPORTANCE_ICON = 2
Const COLUMN_CHECK_ICON = 3
Const RECORD_UNREAD_MAIL_ICON = 4
Const RECORD_READ_MAIL_ICON = 5
Const RECORD_REPLIED_ICON = 6
Const RECORD_IMPORTANCE_HIGH_ICON = 7
Const COLUMN_ATTACHMENT_ICON = 8
Const COLUMN_ATTACHMENT_NORMAL_ICON = 9
Const RECORD_IMPORTANCE_LOW_ICON = 10

Const IMPORTANCE_HIGH = 0
Const IMPORTANCE_NORMAL = 1
Const IMPORTANCE_LOW = 2

Const CHECKED_TRUE = 1
Const CHECKED_FALSE = 0

Const READ_TRUE = 1
Const READ_FALSE = 0

Const ATTACHMENTS_TRUE = 1
Const ATTACHMENTS_FALSE = 0

Dim MaxRowIndex As Long
Dim fntBold As StdFont
Dim fntStrike As StdFont


Private Sub cmdCopiarOtraEmpresa_Click()
      
      If Val(CodigoActual) = 0 Then Exit Sub
      
            CadenaDesdeOtroForm = "NO"  'Para que no seleccione ninguna empresa por defecto
            frmMensajes.Opcion = 4
            frmMensajes.Show vbModal
            If CadenaDesdeOtroForm = "" Then Exit Sub
            
            BuscaChekc = ""
            NumRegElim = RecuperaValor(CadenaDesdeOtroForm, 1)
            Indice = 0
            For i = 1 To NumRegElim
                Indice = Val(RecuperaValor(CadenaDesdeOtroForm, i + 1 + NumRegElim))
                BuscaChekc = BuscaChekc & RecuperaValor(CadenaDesdeOtroForm, i + 1) & vbCrLf
                
                If Indice = vEmpresa.codempre Then
                    MsgBox "No puede seleccionar la empresa actual: " & vEmpresa.nomempre, vbExclamation
                    Exit Sub
                End If
            Next
            
            BuscaChekc = "Va a copiar la configuracion del usuario para las siguientes empresas: " & vbCrLf & vbCrLf & BuscaChekc
            If MsgBox(BuscaChekc, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            
            
            For i = 1 To NumRegElim
                Indice = Val(RecuperaValor(CadenaDesdeOtroForm, i + 1 + NumRegElim))
                BuscaChekc = "DELETE FROM ariconta" & Indice & ".menus_usuarios WHERE codusu = " & CodigoActual
                Ejecuta BuscaChekc, False
                BuscaChekc = "INSERT IGNORE INTO ariconta" & Indice & ".menus_usuarios SELECT * FROM menus_usuarios WHERE codusu = " & CodigoActual
                Ejecuta BuscaChekc, False
                
            Next
            
            
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False

        CargaDatos "", False

    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    
    kaplicacion = "ariconta"
    
    '****************** canviar la consulta *********************************
    CadenaConsulta = "SELECT cast(concat(right(concat('0000',menus.codigo),4),'0000') as signed),menus.codigo, menus_usuarios.aplicacion, if(cast(concat(right(concat('0000',menus.codigo),4),'0000') as signed) mod 10000<>0,concat('     ', menus.descripcion), menus.descripcion),menus_usuarios.ver, IF(menus_usuarios.ver=1,'*','') as pver, menus_usuarios.creareliminar,  IF(menus_usuarios.creareliminar=1,'*','') as pcreareliminar, menus_usuarios.modificar,  IF(menus_usuarios.modificar=1,'*','') as pmodificar, menus_usuarios.imprimir,  IF(menus_usuarios.imprimir=1,'*','') as pimprimir, menus_usuarios.especial, IF(menus_usuarios.especial=1,'*','') as pespecial "
    CadenaConsulta = CadenaConsulta & " from menus, menus_usuarios "
    CadenaConsulta = CadenaConsulta & " where menus.aplicacion = " & DBSet(kaplicacion, "T")
    CadenaConsulta = CadenaConsulta & " and menus.padre = 0 "
    CadenaConsulta = CadenaConsulta & " and menus.codigo > 1 "
    CadenaConsulta = CadenaConsulta & " and menus.aplicacion = menus_usuarios.aplicacion and menus.codigo = menus_usuarios.codigo and menus_usuarios.codusu = " & DBSet(CodigoActual, "N")
    
    CadenaConsulta = CadenaConsulta & " UNION "
    CadenaConsulta = CadenaConsulta & " select cast(concat(right(concat('0000',hh.padre),4), right(concat('0000',hh.orden),4)) as signed), hh.codigo, hh.aplicacion, if(cast(concat(right(concat('0000',hh.padre),4), right(concat('0000',hh.orden),4)) as signed) mod 10000<>0,concat('     ', hh.descripcion), hh.descripcion), uu.ver, IF(uu.ver=1,'*','') as pver, uu.creareliminar,  IF(uu.creareliminar=1,'*','') as pcreareliminar, uu.modificar,  IF(uu.modificar=1,'*','') as pmodificar, uu.imprimir,  IF(uu.imprimir=1,'*','') as pimprimir, uu.especial, IF(uu.especial=1,'*','') as pespecial  "
    CadenaConsulta = CadenaConsulta & " from menus pp, menus hh, menus_usuarios uu "
    CadenaConsulta = CadenaConsulta & " where pp.aplicacion = " & DBSet(kaplicacion, "T")
    CadenaConsulta = CadenaConsulta & " AND hh.padre > 1 and "
    CadenaConsulta = CadenaConsulta & " pp.aplicacion = hh.aplicacion And hh.Padre = pp.Codigo and "
    CadenaConsulta = CadenaConsulta & " hh.aplicacion = uu.aplicacion and hh.codigo = uu.codigo and uu.codusu = " & DBSet(CodigoActual, "N")
    
    
    '************************************************************************
    
    wndReportControl.Icons = ReportControlGlobalSettings.Icons
    wndReportControl.PaintManager.NoItemsText = "Ning�n registro "
    
    
    CreateReportControl
    
End Sub


Private Sub wndReportControl_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim SQL As String
Dim IT As ListItem
Dim Codigo As String

Dim Ver As Integer
Dim CrearEliminar As Integer
Dim Modificar As Integer
Dim Imprimir As Integer
Dim Especial As Integer

Dim Padre As Integer
Dim aplicacion As String

Dim Inicio As Long
Dim CodigoAux As Long

    Codigo = RecuperaValor(Row.Record(0).Tag, 1)
    aplicacion = RecuperaValor(Row.Record(0).Tag, 2)
    
    Ver = 0
    CrearEliminar = 0
    Modificar = 0
    Imprimir = 0
    Especial = 0
    If DBLet(Row.Record(2).Checked, "B") Then Ver = 1
    If DBLet(Row.Record(3).Checked, "B") Then CrearEliminar = 1
    If DBLet(Row.Record(4).Checked, "B") Then Modificar = 1
    If DBLet(Row.Record(5).Checked, "B") Then Imprimir = 1
    If DBLet(Row.Record(6).Checked, "B") Then Especial = 1
    
    '[Monica]06/11/2019: en negrita si est� seleccionado
    If Ver = 1 Then
        Row.Record(1).Bold = True
    Else
        Row.Record(1).Bold = False
    End If
    
    
    SQL = "update menus_usuarios set ver = " & DBLet(Ver, "N")
    SQL = SQL & ", creareliminar = " & DBLet(CrearEliminar, "N")
    SQL = SQL & ", modificar = " & DBLet(Modificar, "N")
    SQL = SQL & ", imprimir = " & DBLet(Imprimir, "N")
    SQL = SQL & ", especial = " & DBLet(Especial, "N")
    SQL = SQL & " where codusu = " & DBSet(CodigoActual, "N") & " and codigo = " & DBSet(Codigo, "N")
    SQL = SQL & " and aplicacion = " & DBSet(aplicacion, "T")
    
    Conn.Execute SQL
    
    If Codigo < 100 Then
        ' si es padre marco o desmarco todos sus hijos, si me han tocado ver
        SQL = "update menus_usuarios set "
        If Item.Index = 2 Then SQL = SQL & "ver = " & DBSet(Ver, "N")
        If Item.Index = 3 Then SQL = SQL & "creareliminar = " & DBSet(CrearEliminar, "N")
        If Item.Index = 4 Then SQL = SQL & "modificar = " & DBSet(Modificar, "N")
        If Item.Index = 5 Then SQL = SQL & "imprimir = " & DBSet(Imprimir, "N")
        If Item.Index = 6 Then SQL = SQL & "especial = " & DBSet(Especial, "N")
        SQL = SQL & " where codusu = " & DBSet(CodigoActual, "N") & " and codigo in (select codigo from menus where padre = " & DBSet(Codigo, "N") & " and aplicacion = " & DBSet(aplicacion, "T") & ")"
        Conn.Execute SQL
        
        Inicio = Codigo * 100
        For i = 0 To wndReportControl.Records.Count - 1
            CodigoAux = RecuperaValor(wndReportControl.Records(i).Tag, 1)
            If CodigoAux >= Inicio And CodigoAux < Inicio + 100 Then
                If Item.Index = 2 Then wndReportControl.Records(i).Item(2).Checked = Ver
                '[Monica]06/11/2019: en negrita si est� seleccionado
                If Ver = 1 Then
                    wndReportControl.Records(i).Item(1).Bold = True
                Else
                    wndReportControl.Records(i).Item(1).Bold = False
                End If
                
                
                If Item.Index = 3 Then wndReportControl.Records(i).Item(3).Checked = CrearEliminar
                If Item.Index = 4 Then wndReportControl.Records(i).Item(4).Checked = Modificar
                If Item.Index = 5 Then wndReportControl.Records(i).Item(5).Checked = Imprimir
                If Item.Index = 6 Then wndReportControl.Records(i).Item(6).Checked = Especial
            End If
        Next i
        wndReportControl.Populate
    Else
        Padre = DevuelveValor("select padre from menus where codigo = " & DBSet(Codigo, "N") & " and aplicacion = 'ariagro'")
        If (Item.Index = 2 And Ver = 0) Or (Item.Index = 3 And CrearEliminar = 0) Or (Item.Index = 4 And Modificar = 0) Or (Item.Index = 5 And Imprimir = 0) Or (Item.Index = 6 And Especial = 0) Then  'si todos sus hijos estan desmarcados, desmarco el padre
            SQL = "select * from menus_usuarios where codusu = " & DBSet(CodigoActual, "N") & " and aplicacion = 'ariagro' and codigo in (select codigo from menus where aplicacion = 'ariagro' and padre = " & DBSet(Padre, "N") & ")"
            If Item.Index = 2 Then SQL = SQL & " and ver = 1 "
            If Item.Index = 3 Then SQL = SQL & " and creareliminar = 1 "
            If Item.Index = 4 Then SQL = SQL & " and modificar = 1 "
            If Item.Index = 5 Then SQL = SQL & " and imprimir = 1 "
            If Item.Index = 6 Then SQL = SQL & " and especial = 1 "

            If TotalRegistrosConsulta(SQL) = 0 Then
                SQL = "update menus_usuarios set "
                If Item.Index = 2 Then SQL = SQL & " Ver = 0"
                If Item.Index = 3 Then SQL = SQL & " creareliminar = 0"
                If Item.Index = 4 Then SQL = SQL & " modificar = 0"
                If Item.Index = 5 Then SQL = SQL & " imprimir = 0"
                If Item.Index = 6 Then SQL = SQL & " especial = 0"
                
                SQL = SQL & " where codusu = " & DBSet(CodigoActual, "N") & " and aplicacion = 'ariagro' and codigo = " & DBSet(Padre, "N")
                Conn.Execute SQL
                
                For i = 0 To wndReportControl.Records.Count - 1
                    CodigoAux = RecuperaValor(wndReportControl.Records(i).Tag, 1)
                    If CodigoAux = Padre Then
                        If Item.Index = 2 Then wndReportControl.Records(i).Item(2).Checked = False
                        '[Monica]06/11/2019: en negrita si est� seleccionado
                        wndReportControl.Records(i).Item(1).Bold = False
                        
                        If Item.Index = 3 Then wndReportControl.Records(i).Item(3).Checked = False
                        If Item.Index = 4 Then wndReportControl.Records(i).Item(4).Checked = False
                        If Item.Index = 5 Then wndReportControl.Records(i).Item(5).Checked = False
                        If Item.Index = 6 Then wndReportControl.Records(i).Item(6).Checked = False
                    End If
                Next i
                
                wndReportControl.Populate
            End If
        
        Else ' como marco un hijo, marco al padre
            SQL = "update menus_usuarios set "
            If Item.Index = 2 Then SQL = SQL & " Ver = 1"
            If Item.Index = 3 Then SQL = SQL & " creareliminar = 1"
            If Item.Index = 4 Then SQL = SQL & " modificar = 1"
            If Item.Index = 5 Then SQL = SQL & " imprimir = 1"
            If Item.Index = 6 Then SQL = SQL & " especial = 1"
            
            SQL = SQL & " where codusu = " & DBSet(CodigoActual, "N") & " and codigo = " & DBSet(Padre, "N") & " and aplicacion = 'ariagro'"
            Conn.Execute SQL
            
            For i = 0 To wndReportControl.Records.Count - 1
                CodigoAux = RecuperaValor(wndReportControl.Records(i).Tag, 1)
                If CodigoAux = Padre Then
                    If Item.Index = 2 Then wndReportControl.Records(i).Item(2).Checked = True
                    '[Monica]06/11/2019: en negrita si est� seleccionado
                    wndReportControl.Records(i).Item(1).Bold = True
                    If Item.Index = 3 Then wndReportControl.Records(i).Item(3).Checked = True
                    If Item.Index = 4 Then wndReportControl.Records(i).Item(4).Checked = True
                    If Item.Index = 5 Then wndReportControl.Records(i).Item(5).Checked = True
                    If Item.Index = 6 Then wndReportControl.Records(i).Item(6).Checked = True
                End If
            Next i
            
            wndReportControl.Populate
        End If
    End If
    
End Sub



Public Sub CreateReportControl()
    'Start adding columns
    Dim Column As ReportColumn
    
    wndReportControl.Columns.DeleteAll
    
    
    Set Column = wndReportControl.Columns.Add(1, "Nombre", 180, True)
    Set Column = wndReportControl.Columns.Add(2, "Ver", 70, True)
    Column.Alignment = xtpAlignmentIconLeft
    Set Column = wndReportControl.Columns.Add(3, "Crear/Eliminar", 70, True)
    Column.Alignment = xtpAlignmentIconLeft
    Set Column = wndReportControl.Columns.Add(4, "Modificar", 70, True)
    Column.Alignment = xtpAlignmentIconLeft
    Set Column = wndReportControl.Columns.Add(5, "Imprimir", 70, True)
    Column.Alignment = xtpAlignmentIconLeft
    Set Column = wndReportControl.Columns.Add(6, "Especial", 60, True)
    Column.Alignment = xtpAlignmentIconLeft
    
'    wndReportControl.PaintManager.CaptionFont = verdana
    
    
    Dim TextFont As StdFont
    Set TextFont = Me.Font
'    TextFont.SIZE = 10
    Set wndReportControl.PaintManager.TextFont = TextFont
    Set wndReportControl.PaintManager.CaptionFont = TextFont
    Set wndReportControl.PaintManager.PreviewTextFont = TextFont
    
    
    
    'cabecera
    wndReportControl.PaintManager.HeaderHeight = 30
    wndReportControl.PaintManager.CaptionFont.SIZE = 10
    wndReportControl.PaintManager.CaptionFont.Bold = False
    wndReportControl.PaintManager.CaptionFont.Italic = True
    wndReportControl.AllowColumnSort = False
    wndReportControl.AllowColumnRemove = False
    wndReportControl.AllowColumnReorder = False
    wndReportControl.AllowColumnResize = False
    
    'cuerpo
    wndReportControl.PaintManager.TextFont.SIZE = 10
    wndReportControl.PaintManager.TextFont.Italic = False
    wndReportControl.PaintManager.TextFont.Bold = False
    
    
    wndReportControl.PaintManager.MaxPreviewLines = 1
    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
                  
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Strikethrough to the currently set text font
    Set fntStrike = wndReportControl.PaintManager.TextFont
    fntStrike.Strikethrough = True
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Bold to the currently set text font
    Set fntBold = wndReportControl.PaintManager.TextFont
    fntBold.Bold = True
    
    
    
    
    
    'Any time you add or delete rows(by removing the attached record), you must call the
    'Populate method so the ReportControl will display the changes.
    'If rows are added, the rows will remain hidden until Populate is called.
    'If rows are deleted, the rows will remain visible until Populate is called.
    wndReportControl.Populate
    
    wndReportControl.SetCustomDraw xtpCustomBeforeDrawRow
    
    
End Sub



'Cuando modifiquemos o insertemos, pondremos el SQL entero
Public Sub CargaDatos(ByVal SQL As String, EsTodoSQL As Boolean)
Dim Aux  As String
Dim Inicial As Integer
Dim N As Integer
Dim v As Boolean
Dim T1 As Single



    On Error GoTo ECargaDatos

    Screen.MousePointer = vbHourglass
'    statusBar.Panels(1).Text = "Leyendo BD"
    wndReportControl.ShowItemsInGroups = False
    wndReportControl.Records.DeleteAll
    wndReportControl.Populate
    
    Set miRsAux = New ADODB.Recordset
    
    SQL = "SELECT cast(concat(right(concat('0000',menus.codigo),4),'0000') as signed),menus.codigo, menus_usuarios.aplicacion, if(cast(concat(right(concat('0000',menus.codigo),4),'0000') as signed) mod 10000<>0,concat('     ', menus.descripcion), menus.descripcion),menus_usuarios.ver, IF(menus_usuarios.ver=1,'*','') as pver, menus_usuarios.creareliminar,  IF(menus_usuarios.creareliminar=1,'*','') as pcreareliminar, menus_usuarios.modificar,  IF(menus_usuarios.modificar=1,'*','') as pmodificar, menus_usuarios.imprimir,  IF(menus_usuarios.imprimir=1,'*','') as pimprimir, menus_usuarios.especial, IF(menus_usuarios.especial=1,'*','') as pespecial "
    SQL = SQL & " from menus, menus_usuarios "
    SQL = SQL & " where menus.aplicacion =" & DBSet(kaplicacion, "T")
    SQL = SQL & " and menus.padre = 0 "
    SQL = SQL & " and menus.codigo > 1 "
    SQL = SQL & " and menus.aplicacion = menus_usuarios.aplicacion and menus.codigo = menus_usuarios.codigo and menus_usuarios.codusu = " & DBSet(CodigoActual, "N")
    
    
    SQL = SQL & " UNION "
    SQL = SQL & " select cast(concat(right(concat('0000',hh.padre),4), right(concat('0000',hh.orden),4)) as signed), hh.codigo, hh.aplicacion, if(cast(concat(right(concat('0000',hh.padre),4), right(concat('0000',hh.orden),4)) as signed) mod 10000<>0,concat('     ', hh.descripcion), hh.descripcion), uu.ver, IF(uu.ver=1,'*','') as pver, uu.creareliminar,  IF(uu.creareliminar=1,'*','') as pcreareliminar, uu.modificar,  IF(uu.modificar=1,'*','') as pmodificar, uu.imprimir,  IF(uu.imprimir=1,'*','') as pimprimir, uu.especial, IF(uu.especial=1,'*','') as pespecial  "
    SQL = SQL & " from menus pp, menus hh, menus_usuarios uu "
    SQL = SQL & " where pp.aplicacion =" & DBSet(kaplicacion, "T")
    SQL = SQL & " AND hh.padre > 1 and "
    SQL = SQL & " pp.aplicacion = hh.aplicacion And hh.Padre = pp.Codigo and "
    SQL = SQL & " hh.aplicacion = uu.aplicacion and hh.codigo = uu.codigo and uu.codusu = " & DBSet(CodigoActual, "N")
    
    SQL = SQL & " ORDER BY 1 "
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Inicial = 0
    T1 = Timer
    While Not miRsAux.EOF
        AddRecord2
        wndReportControl.Populate
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    wndReportControl.Populate
    
    
    
ECargaDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, SQL
    
    
    
'    statusBar.Panels(1).Text = ""
    Screen.MousePointer = vbDefault
End Sub

'socio, pendiente , nombre, matricula,licencia
'Leera los datos de mirsaux
Private Sub AddRecord2()

Dim Record As ReportRecord
Dim Socio As Boolean
Dim OtroIcono As Boolean
Dim NoActivo As Boolean
Dim Nombre As String

    On Error GoTo eAddRecord2
    

    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    'Codigo
    Set Item = Record.AddItem(miRsAux.Fields(0))
    Item.Value = DBLet(miRsAux.Fields(0))
    Item.Tag = DBLet(miRsAux.Fields(1)) & "|" & DBLet(miRsAux.Fields(2)) & "|"
    Record.Tag = DBLet(miRsAux.Fields(1)) & "|" & DBLet(miRsAux.Fields(2)) & "|"
    
    'Nombre
    Set Item = Record.AddItem(miRsAux.Fields(3))
    Item.Value = DBLet(miRsAux.Fields(3))
    
    '[Monica]04/11/2019
    If DBLet(miRsAux!Ver) = 1 Then
        Item.Bold = True
    Else
        Item.Bold = False
    End If
    
    'ver
    Set Item = Record.AddItem("")
    Item.HasCheckbox = True
    Item.TristateCheckbox = False
    Item.Checked = IIf(DBLet(miRsAux!Ver), CHECKED_TRUE, CHECKED_FALSE)

    'creareliminar
    Set Item = Record.AddItem("")
    Item.HasCheckbox = True
    Item.TristateCheckbox = False
    Item.Checked = IIf(DBLet(miRsAux!CrearEliminar), CHECKED_TRUE, CHECKED_FALSE)

    'modificar
    Set Item = Record.AddItem("")
    Item.HasCheckbox = True
    Item.TristateCheckbox = False
    Item.Checked = IIf(DBLet(miRsAux!Modificar), CHECKED_TRUE, CHECKED_FALSE)

    'imprimir
    Set Item = Record.AddItem("")
    Item.HasCheckbox = True
    Item.TristateCheckbox = False
    Item.Checked = IIf(DBLet(miRsAux!Imprimir), CHECKED_TRUE, CHECKED_FALSE)

    'especial
    Set Item = Record.AddItem("")
    Item.HasCheckbox = True
    Item.TristateCheckbox = False
    Item.Checked = IIf(DBLet(miRsAux!Especial), CHECKED_TRUE, CHECKED_FALSE)

    
eAddRecord2:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub






Private Function DatosOK() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim SQL As String
Dim Mens As String


    B = CompForm(Me)
    If Not B Then Exit Function
    
    
    DatosOK = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
End Sub



