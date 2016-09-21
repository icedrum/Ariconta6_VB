VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListIntegracione 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integraciones"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   Icon            =   "frmListIntegracione.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEli 
      Caption         =   "Eliminar"
      Height          =   435
      Left            =   8760
      TabIndex        =   10
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   435
      Left            =   10080
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Tag             =   "ASIDIA"
      Top             =   960
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   8493
      View            =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Index           =   1
      Left            =   2940
      TabIndex        =   2
      Tag             =   "ASIHCO"
      Top             =   960
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   8493
      View            =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Index           =   2
      Left            =   5820
      TabIndex        =   4
      Tag             =   "FRACLI"
      Top             =   960
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   8493
      View            =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Index           =   3
      Left            =   8700
      TabIndex        =   6
      Tag             =   "FRAPRO"
      Top             =   960
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   8493
      View            =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "Ver"
      Height          =   435
      Left            =   8760
      TabIndex        =   11
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Diario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   4
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   8655
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   8760
      TabIndex        =   7
      Top             =   600
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   5820
      TabIndex        =   5
      Top             =   600
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Histórico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   2940
      TabIndex        =   3
      Top             =   600
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Diario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   2115
   End
End
Attribute VB_Name = "frmListIntegracione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Errores As Boolean

Dim L As Integer



Private Sub cmdEli_Click()
Dim Cad As String
    If MsgBox("Va a eliminar los archivos de errores. Desea continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    'ASientos
    Screen.MousePointer = vbHourglass
    For L = 0 To 3
        Cad = vConfig.Integraciones & "\ERRORES\" & ListView1(L).Tag & "\*.*"
        If Dir(Cad) <> "" Then Kill Cad
        ListView1(L).ListItems.Clear
    Next L
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdVer_Click()
'Dim Cad As String
'    Cad = "C:\WINDOWS\NOTEPAD.EXE"
'    If Dir(Cad, vbArchive) = "" Then Cad = "C:\WINNT\NOTEPAD.EXE"
'    If Dir(Cad, vbArchive) = "" Then
'        MsgBox "No se encuentra el NOTEPAD. ( " & Cad & ")", vbExclamation
'        Exit Sub
'    End If
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

'Recorrera las carpetas para añadir los archivos que tienen
Private Sub Form_Load()

    Me.Icon = frmPpal.Icon

    If Errores Then
        Label1(4).Caption = "Archivos con errores"
        Label1(4).ForeColor = &HC0&
    Else
        Label1(4).Caption = "Archivos pendientes de integrar"
        Label1(4).ForeColor = &H4000&
    End If
    For L = 0 To 3
        ListView1(L).SmallIcons = frmPpal.ImageList1
        ListView1(L).Icons = frmPpal.ImageList1
    Next L
    cmdEli.Visible = Errores
    cmdVer.Visible = Errores
    CargaList
End Sub




Private Sub CargaList()
Dim Cad As String
Dim Carpeta As String
Dim path As String
Dim ItmX As ListItem


    For L = 0 To 3
        ListView1(L).ListItems.Clear
    Next L

    If Errores Then
        path = vConfig.Integraciones & "\ERRORES\"
    Else
        path = vConfig.Integraciones & "\INTEGRA\"
    End If
    
    'ASientos
    L = 0
    Carpeta = path & ListView1(L).Tag & "\*.Z*"
    Cad = Dir(Carpeta, vbArchive)   ' Recupera la primera entrada.
    Do While Cad <> ""   ' Inicia el bucle.
       Set ItmX = ListView1(L).ListItems.Add(, , Cad)
        ItmX.SmallIcon = 1
       Cad = Dir   ' Obtiene siguiente entrada.
    Loop


    'ASientos a historico
    L = 1
    Carpeta = path & ListView1(L).Tag & "\*.Z*"
    Cad = Dir(Carpeta, vbArchive)   ' Recupera la primera entrada.
    Do While Cad <> ""   ' Inicia el bucle.
       Set ItmX = ListView1(L).ListItems.Add(, , Cad)
        ItmX.SmallIcon = 1
       Cad = Dir   ' Obtiene siguiente entrada.
    Loop
    
    
    'Facturas clientes
    L = 2
    Carpeta = path & ListView1(L).Tag & "\*.Z*"
    Cad = Dir(Carpeta, vbArchive)   ' Recupera la primera entrada.
    Do While Cad <> ""   ' Inicia el bucle.
       Set ItmX = ListView1(L).ListItems.Add(, , Cad)
        ItmX.SmallIcon = 1
       Cad = Dir   ' Obtiene siguiente entrada.
    Loop
    
    'Facturas
    L = 3
    Carpeta = path & ListView1(L).Tag & "\*.Z*"
    Cad = Dir(Carpeta, vbArchive)   ' Recupera la primera entrada.
    Do While Cad <> ""   ' Inicia el bucle.
       Set ItmX = ListView1(L).ListItems.Add(, , Cad)
        ItmX.SmallIcon = 1
       Cad = Dir   ' Obtiene siguiente entrada.
    Loop
    
    
End Sub



'Vamos a tratar de integrar el asiento
'para ello:
'       .-Crearemos la cadena de parametros con la que llamaremos a integrar
'       .-Borramos el archivo .Z
'       .-Si lo borramos, es k estamos en posesion de el para lanzarlo
'       .- Cuando acabe, no existe el fichero FLAG, entonces vemos
'       .- El archivo no existe
'                    - Todo correcto
'                   - EXISTE: Ha producido errores
'                           Habra que revisarlos
Private Sub ListView1_DblClick(Index As Integer)
Dim Cad As String
Dim Bien As Boolean
    If ListView1(Index).ListItems.Count < 1 Then Exit Sub
    If ListView1(Index).SelectedItem Is Nothing Then Exit Sub
    
    'Doble clik, integramos el archivo seleccionado
    Cad = ListView1(Index).SelectedItem.Text
    Cad = vConfig.Integraciones & "\INTEGRA\" & ListView1(Index).Tag & "\" & Cad
    Screen.MousePointer = vbHourglass
    If Not Errores Then
        Bien = IntegrarArchivo(Cad, Index)
        If Bien Then
            ListView1(Index).ListItems.Remove ListView1(Index).SelectedItem.Index
        Else
            'Mostraremos archvios con errores
            
        End If
    Else
    
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Function EliminarArchivo(Archi As String) As Boolean
On Error Resume Next
    Kill Archi
    If Err.Number <> 0 Then
        Err.Clear
        EliminarArchivo = False
    Else
        EliminarArchivo = True
    End If
End Function


Private Function IntegrarArchivo(NombreArchivo As String, Index As Integer) As Boolean
Dim Cad As String
Dim Parametros As String
Dim T1 As Single
    On Error GoTo EDob

    IntegrarArchivo = False
    'Opcion
    Parametros = "/A:N /O:" & Index
    Cad = NombreArchivo
    
    If EliminarArchivo(Cad) Then
        'Obtenemos el nombre
        Cad = ListView1(Index).SelectedItem.Text
        L = InStr(1, Cad, ".")
        Cad = Mid(Cad, 1, L) & "0" & Mid(Cad, L + 2)
        Cad = vConfig.Integraciones & "\INTEGRA\" & ListView1(Index).Tag & "\" & Cad
        Parametros = Parametros & " /F:" & Cad
        
        'Ejecutamos el integrar
        Shell App.path & "\integcon.exe " & Parametros, vbNormalFocus
        
        T1 = Timer
        Do
            Parametros = Dir(App.path & "\FLAG.txt")
            If Parametros <> "" Then
                If Timer - T1 > 60 Then
                    MsgBox "Transcurridos 60 segundos el proceso no finaliza.", vbExclamation
                    Parametros = ""
                End If
            End If
        Loop Until Parametros = ""
        Me.Refresh
        Me.SetFocus
        
        'Ya ha acabado, comprobamos si ha sido integrado o no
       'Comprobamos a ver si esta el archivo
        If Dir(NombreArchivo) = "" Then IntegrarArchivo = True
        
    End If
    
    
    Exit Function
EDob:
    MuestraError Err.Number, "Integra archivo: " & Cad
End Function
