Attribute VB_Name = "myODBC"
Option Explicit
'declaracion de constantes
Private Const ODBC_ADD_DSN = 1 ' Add data source
Private Const ODBC_CONFIG_DSN = 2 ' Configure (edit) data source
Private Const ODBC_REMOVE_DSN = 3 ' Remove data source
Private Const vbAPINull As Long = 0& ' NULL Pointer
'delcaracion de funciones
#If Win32 Then
    Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
            (ByVal hwndParent As Long, ByVal fRequest As Long, _
             ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
#Else
    Private Declare Function SQLConfigDataSource Lib "ODBCINST.DLL" _
            (ByVal hwndParent As Integer, ByVal fRequest As Integer, ByVal _
            lpszDriver As String, ByVal lpszAttributes As String) As Integer
#End If


Public Sub Crea_DSN()
'ModificaDSN
End Sub

'Para crear un DSN :
Private Sub CrearDSN()


#If Win32 Then
    Dim intRet As Long
#Else
    Dim intRet As Integer
#End If
Dim strDriver As String
Dim strAttributes As String

'usamos el driver de SQL Server porque es el mas comun
strDriver = "MySQL ODBC 3.51 Driver"
'Asignamos los parametros separados por null.
strAttributes = "SERVER=SRVCENTRAL" & Chr$(0)
strAttributes = strAttributes & "DESCRIPTION=Temp DSN" & Chr$(0)
strAttributes = strAttributes & "DSN=DSNTEMP" & Chr$(0)
strAttributes = strAttributes & "DATABASE=Conta2" & Chr$(0)
strAttributes = strAttributes & "UID=root" & Chr$(0)
strAttributes = strAttributes & "PWD=aritel" & Chr$(0)
'Para mostrar el diálogo usar Form1.Hwnd en vez de vbAPINull.
intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, strDriver, strAttributes)
If intRet Then
    MsgBox "DSN Creado"
Else
    MsgBox "Fallo en la creación"
End If
End Sub




'Para borrarlo:
Private Sub BorrarDSN()


#If Win32 Then
    Dim intRet As Long
#Else
    Dim intRet As Integer
#End If
Dim strDriver As String
Dim strAttributes As String
'usamos el driver de SQL Server porque es el mas comun
strDriver = "SQL Server"
'Asignamos los parametros separados por null.
strAttributes = "DSN=DSNTEMP" & Chr$(0)
'Para mostrar el diálogo usar Form1.Hwnd en vez de vbAPINull.
intRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_DSN, strDriver, strAttributes)
If intRet Then
    MsgBox "DSN Eliminado"
Else
    MsgBox "Fallo en el borrado"
End If
End Sub



'Para modificarlo:
Private Sub ModificaDSN()

#If Win32 Then
    Dim intRet As Long
#Else
    Dim intRet As Integer
#End If
Dim strDriver As String
Dim strAttributes As String

'usamos el driver de SQL Server porque es el mas comun
strDriver = "MySQL ODBC 3.51 Driver"
'Asignamos los parametros separados por null.
strAttributes = "SERVER=SRVCENTRAL" & Chr$(0)
strAttributes = strAttributes & "DESCRIPTION=Temp DSN modificado" & Chr$(0)
strAttributes = strAttributes & "DSN=DSNTEMP" & Chr$(0)
strAttributes = strAttributes & "DATABASE=Conta" & Chr$(0)
strAttributes = strAttributes & "UID=root" & Chr$(0)
strAttributes = strAttributes & "PWD=aritel" & Chr$(0)

''Para mostrar el diálogo usar Form1.Hwnd en vez de vbAPINull.
intRet = SQLConfigDataSource(vbAPINull, ODBC_CONFIG_DSN, strDriver, strAttributes)
If intRet Then
    MsgBox "DSN Modificado"
Else
    MsgBox "Fallo en la modificacion"
End If
End Sub

