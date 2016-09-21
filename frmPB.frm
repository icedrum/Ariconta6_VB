VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPB 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   900
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Maximo As Integer

Private Sub Form_Load()
Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = Maximo
End Sub



Public Sub Incrementa()
    On Error Resume Next
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
    If Err.Number <> 0 Then Err.Clear
    Me.Refresh
End Sub
