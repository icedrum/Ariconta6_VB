VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVarios 
   Caption         =   "Ariconta 2014"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   Icon            =   "frmVarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9120
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Documentos de Interés"
      TabPicture(0)   =   "frmVarios.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin MSComctlLib.ListView ListView1 
         Height          =   4035
         Left            =   300
         TabIndex        =   1
         Top             =   990
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7117
         View            =   2
         LabelEdit       =   1
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Documentos disponibles"
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
         Left            =   300
         TabIndex        =   2
         Top             =   630
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Integer


Private Sub Form_Activate()
    Me.SSTab1.Tab = Opcion - 6 '5
    If Me.ListView1.Tag = 0 Then
        Me.ListView1.Tag = 1
        
        'Primera vez
        Cargadocumentos
        'ImageListDocumentos
        
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    
    Me.ListView1.Tag = 0
    
End Sub

Private Sub Cargadocumentos()
Dim RN As ADODB.Recordset
Dim cad As String
Dim IT As ListItem

    Set Me.ListView1.SmallIcons = frmPpal.ImageList1 'frmPpal.ImageListDocumentos

    cad = "select iddocumento,nombrearchi from usuarios.wfichdocs WHERE aplicacion='ariconta' order by orden "
    Set RN = New ADODB.Recordset
    RN.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        ListView1.ListItems.Add , "D" & Format(RN!iddocumento, "00000"), RN!nombrearchi, , 7   '1:PDF
       
        RN.MoveNext
    Wend
    RN.Close
End Sub

Private Sub ListView1_DblClick()
Dim Abrir As Boolean

    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
    
    Abrir = False 'antes \ImgFicFT
    If Dir(App.Path & "\temp\" & ListView1.SelectedItem & ".pdf", vbArchive) = "" Then
        Adodc1.ConnectionString = Conn
        Adodc1.RecordSource = "Select * from usuarios.wfichdocs where idDocumento=" & Mid(ListView1.SelectedItem.Key, 2)
        Adodc1.Refresh

        If LeerBinary(Adodc1.Recordset!Campo, App.Path & "\temp\" & ListView1.SelectedItem.Text & ".pdf") Then Abrir = True
    Else
        Abrir = True
        
    End If
    
    If Abrir Then LanzaVisorMimeDocumento Me.hWnd, App.Path & "\temp\" & ListView1.SelectedItem & ".pdf"
        
End Sub
