VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPunteoBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punteo bancario"
   ClientHeight    =   10020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18315
   Icon            =   "frmPunteoBanco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   18315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDatosCobroPago 
      Height          =   855
      Left            =   120
      TabIndex        =   58
      Top             =   720
      Visible         =   0   'False
      Width           =   18015
      Begin VB.TextBox txtVto 
         BackColor       =   &H80000018&
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
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   360
         Width           =   12555
      End
      Begin VB.TextBox txtVto 
         BackColor       =   &H80000018&
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
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox txtVto 
         BackColor       =   &H80000018&
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   11
         Left            =   2640
         TabIndex        =   61
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
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
         Index           =   10
         Left            =   240
         TabIndex        =   59
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   3480
      TabIndex        =   44
      Top             =   120
      Width           =   14145
      Begin VB.CommandButton cmdImportar 
         Caption         =   "Importar"
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
         Left            =   12600
         TabIndex        =   47
         Top             =   180
         Width           =   1035
      End
      Begin VB.TextBox Text12 
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
         Left            =   1500
         TabIndex        =   46
         Top             =   240
         Width           =   8895
      End
      Begin VB.CheckBox chkElimmFich 
         Caption         =   "Eliminar fichero "
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
         Left            =   10560
         TabIndex        =   45
         Top             =   240
         Value           =   1  'Checked
         Width           =   1965
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   1740
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   1200
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
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
         Index           =   7
         Left            =   180
         TabIndex        =   48
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7575
      Left            =   3480
      TabIndex        =   49
      Top             =   1560
      Width           =   12135
      Begin VB.CommandButton cmdReplace 
         Height          =   255
         Left            =   240
         Picture         =   "frmPunteoBanco.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Sustituir cadena"
         Top             =   7080
         Width           =   255
      End
      Begin VB.TextBox txtSald 
         BackColor       =   &H80000018&
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
         Left            =   10560
         TabIndex        =   56
         Top             =   307
         Width           =   1275
      End
      Begin VB.TextBox txtSald 
         BackColor       =   &H80000018&
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
         Left            =   2160
         TabIndex        =   54
         Top             =   307
         Width           =   1275
      End
      Begin VB.TextBox txtDatos 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5955
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   52
         Text            =   "frmPunteoBanco.frx":0A0E
         Top             =   840
         Width           =   11715
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Integrar"
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
         Left            =   9120
         TabIndex        =   51
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Volver"
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
         Left            =   10560
         TabIndex        =   50
         Top             =   6960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldo inicial fichero:"
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
         Index           =   9
         Left            =   7560
         TabIndex        =   55
         Top             =   360
         Width           =   2835
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo final BBDD:"
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
         Index           =   8
         Left            =   240
         TabIndex        =   53
         Top             =   360
         Width           =   1875
      End
   End
   Begin VB.Frame FrameGenera 
      Height          =   4935
      Left            =   6000
      TabIndex        =   20
      Top             =   2760
      Width           =   7665
      Begin VB.TextBox Text11 
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
         Left            =   240
         MaxLength       =   15
         TabIndex        =   25
         Text            =   "000000000000000"
         Top             =   2880
         Width           =   1965
      End
      Begin VB.CommandButton cmdAtoCancelar 
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
         Left            =   6240
         TabIndex        =   29
         Top             =   4440
         Width           =   1155
      End
      Begin VB.CommandButton cmdAstoAceptar 
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
         Left            =   4980
         TabIndex        =   28
         Top             =   4440
         Width           =   1155
      End
      Begin VB.TextBox txtFec 
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
         Left            =   240
         TabIndex        =   22
         Text            =   "99/99/9999"
         Top             =   780
         Width           =   1245
      End
      Begin VB.TextBox Text10 
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
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   1440
         Width           =   6225
      End
      Begin VB.TextBox Text9 
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
         Left            =   240
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1440
         Width           =   885
      End
      Begin VB.TextBox Text8 
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
         Left            =   2280
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   2880
         Width           =   5085
      End
      Begin VB.TextBox Text7 
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
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   2160
         Width           =   6225
      End
      Begin VB.TextBox Text6 
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
         Left            =   240
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   2160
         Width           =   885
      End
      Begin VB.TextBox Text5 
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
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   3600
         Width           =   5625
      End
      Begin VB.TextBox Text4 
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   3600
         Width           =   1405
      End
      Begin VB.Label Label10 
         Caption         =   "Documento"
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
         Left            =   240
         TabIndex        =   39
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   1680
         Top             =   3330
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   1230
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   900
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Si no pone contrapartida podrá añadir más de una línea en el asiento"
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
         Left            =   240
         TabIndex        =   38
         Top             =   3960
         Width           =   7275
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
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
         Left            =   240
         TabIndex        =   37
         Top             =   540
         Width           =   705
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   2
         Left            =   1200
         Picture         =   "frmPunteoBanco.frx":0A14
         Top             =   540
         Width           =   240
      End
      Begin VB.Label Label8 
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
         Height          =   225
         Left            =   2280
         TabIndex        =   36
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label Label7 
         Caption         =   "Diario"
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
         Left            =   240
         TabIndex        =   35
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Concepto"
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
         Left            =   240
         TabIndex        =   33
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Label Label5 
         Caption         =   "Contrapartida"
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
         Left            =   240
         TabIndex        =   32
         Top             =   3360
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Index           =   5
         Left            =   90
         TabIndex        =   21
         Top             =   180
         Width           =   5505
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   150
      TabIndex        =   41
      Top             =   120
      Width           =   3015
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   42
         Top             =   180
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Crear Asiento"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver asiento"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver vencimientos"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "V1"
                     Text            =   "Cobros    "
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "V2"
                     Text            =   "Pagos"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "G1"
                     Text            =   "Realizar cobro Ctrl+C"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "G2"
                     Text            =   "Realizar pago  Ctrl+B"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "G4"
                     Text            =   "Pendiente cliente  Ctrl+V"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameIntro 
      Height          =   885
      Left            =   150
      TabIndex        =   4
      Top             =   1080
      Width           =   18075
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar punteados"
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
         Left            =   14760
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   2340
      End
      Begin VB.TextBox Text1 
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
         Left            =   1560
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   330
         Width           =   1575
      End
      Begin VB.TextBox Text2 
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
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   330
         Width           =   5175
      End
      Begin VB.TextBox txtFec 
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
         Left            =   12930
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   330
         Width           =   1275
      End
      Begin VB.TextBox txtFec 
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
         Left            =   10170
         TabIndex        =   1
         Text            =   "99/99/9999"
         Top             =   330
         Width           =   1275
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   12630
         Picture         =   "frmPunteoBanco.frx":0A9F
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   9870
         Picture         =   "frmPunteoBanco.frx":0B2A
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Left            =   1140
         Top             =   330
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fin"
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
         Left            =   11550
         TabIndex        =   6
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha inicio"
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
         Left            =   8550
         TabIndex        =   5
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta"
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
         Left            =   240
         TabIndex        =   8
         Top             =   330
         Width           =   915
      End
   End
   Begin VB.Frame FrameDatos 
      Height          =   7875
      Left            =   180
      TabIndex        =   9
      Top             =   2040
      Width           =   18105
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   15390
         TabIndex        =   16
         Top             =   7290
         Width           =   2200
      End
      Begin VB.TextBox Text3 
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
         Index           =   1
         Left            =   13050
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   7290
         Width           =   2200
      End
      Begin VB.TextBox Text3 
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
         Left            =   10770
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   7290
         Width           =   2200
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6300
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   11113
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Importe"
            Object.Width           =   2471
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "D/H"
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Concepto"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6300
         Left            =   8490
         TabIndex        =   11
         Top             =   600
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   11113
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Asiento"
            Object.Width           =   1872
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "D/H"
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Saldo"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Ampliacion"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   17100
         Picture         =   "frmPunteoBanco.frx":0BB5
         ToolTipText     =   "Quitar punteados"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   17550
         Picture         =   "frmPunteoBanco.frx":0CFF
         ToolTipText     =   "Puntear"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   7680
         Picture         =   "frmPunteoBanco.frx":0E49
         ToolTipText     =   "Quitar punteados"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   8130
         Picture         =   "frmPunteoBanco.frx":0F93
         ToolTipText     =   "Puntear"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label11 
         Caption         =   "z"
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
         Left            =   330
         TabIndex        =   40
         Tag             =   "Doble click busca el importe en el otro lado"
         Top             =   7320
         Width           =   4815
      End
      Begin VB.Label Label4 
         Caption         =   "DIFERENCIA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   15390
         TabIndex        =   19
         Top             =   7020
         Width           =   1845
      End
      Begin VB.Label Label3 
         Caption         =   "CONTABILIDAD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13050
         TabIndex        =   18
         Top             =   7020
         Width           =   1635
      End
      Begin VB.Label Label2 
         Caption         =   "BANCO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10770
         TabIndex        =   17
         Top             =   7020
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Extracto bancario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   13
         Top             =   180
         Width           =   4575
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   4
         Left            =   8490
         TabIndex        =   12
         Top             =   180
         Width           =   5775
      End
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   17640
      TabIndex        =   43
      Top             =   120
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
End
Attribute VB_Name = "frmPunteoBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 314

Public VerCobrosPagos As Boolean

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCta As frmBasico2
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmCo As frmConceptos
Attribute frmCo.VB_VarHelpID = -1
Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmCC As frmColCtas
Attribute frmCC.VB_VarHelpID = -1
Private frmTESVerCobPag As frmTESVerCobrosPagos
Attribute frmTESVerCobPag.VB_VarHelpID = -1


Dim Sql As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Importe As Currency
Dim I As Long
Dim PrimeraSeleccion As Boolean
Dim ClickAnterior As Byte '0 Empezar 1.-Debe 2.-Haber
    
'Con estas dos variables
Dim ContadorBus As Integer
Dim Checkear As Boolean
Dim De As Currency
Dim Ha As Currency
Dim EstaLW1 As Boolean

Dim CuentaAnterior As String
Dim FechaAnterior As String

Dim NF As Integer
Dim FicheroPpal As String
Dim Cta As String
Dim Saldo As Currency
Dim cad As String


Private Sub Check1_Click()
    CuentaAnterior = Text1.Text
    ConfirmarDatos False
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    CuentaAnterior = Text1.Text
    ConfirmarDatos False
    KEYpress KeyAscii
End Sub

Private Sub ConfirmarDatos(DesdeCuenta As Boolean)
    Screen.MousePointer = vbHourglass
    If Text1.Text <> "" Then
        If CuentaAnterior <> "" Then BloqueoManual False, "PUNTEOB", CuentaAnterior
    
        'Tiene cta.
        'Veamos si la cuenta esta definida en ctas bancarias o no
        Sql = DevuelveDesdeBD("codmacta", "bancos", "codmacta", Text1.Text, "T")
        If Sql <> "" Then
            'Bloqueamos manualamente la tabla, con esa cuenta
            If Not BloqueoManual(True, "PUNTEOB", Text1.Text) Then
                MsgBox "Imposible acceder a puntear la cuenta. Esta bloqueada", vbExclamation
            Else
                Text3(0).Text = "": Text3(1).Text = "": Text3(2).Text = ""
                'Datos ok. Vamos a ver los resultados
                Label1(4).Caption = Text1.Text & " - " & Text2.Text
   '             PonerTamanyo True
                Me.Refresh
                CargarDatosLw True
            End If
        Else
            MsgBox "La cuenta no esta asociada a una cuenta bancaria.", vbExclamation
        End If
    Else
        MsgBox "Introduzca la cuenta ", vbExclamation
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargarDatosLw(BorrarImportes As Boolean)

    If txtFec(0).Text = "" Or txtFec(1).Text = "" Then Exit Sub



       'Resetamos importes punteados
       If BorrarImportes Then
            De = 0
            Ha = 0
            Text3(0).Text = "": Text3(1).Text = "": Text3(2).Text = ""
        End If
        PrimeraSeleccion = True
                    
        'Cargamos los datos
        Sql = "DELETE from tmpconextcab where codusu= " & vUsu.Codigo
        Conn.Execute Sql
            
        Sql = "DELETE from tmpconext where codusu= " & vUsu.Codigo
        Conn.Execute Sql
        
        Sql = "fechaent >= '" & Format(txtFec(0).Text, FormatoFecha)
        Sql = Sql & "' AND fechaent <= '" & Format(txtFec(1).Text, FormatoFecha) & "'"
        
        CargaDatosConExt Text1.Text, txtFec(0).Text, txtFec(1).Text, Sql, Text2.Text

                    
                    
                    
        Me.Refresh
        CargaBancario
        Me.Refresh
        CargaLineaApuntes
        
        FrameBotonGnral.Enabled = True
        
        Me.Refresh
End Sub



Private Sub cmdAstoAceptar_Click()
Dim NA As Long
Dim Sql As String
Dim Sql1 As String


    If txtFec(2).Text = "" Or Text9.Text = "" Or Text7.Text = "" Then
        MsgBox "Todos los campos, excepto la contrapartida, son obligados", vbExclamation
        Exit Sub
    End If
    
   
    'Generamos el asiento en errores
    If Not IsDate(txtFec(2).Text) Then
        MsgBox "Fecha incorrecta", vbExclamation
        Exit Sub
    End If
    
    varFecOk = FechaCorrecta2(CDate(txtFec(2).Text))
    If varFecOk > 1 Then
        If varFecOk = 2 Then
            Sql = varTxtFec
        Else
            Sql = "Fechas fuera de ejercicio actual/siguiente"
        End If
        MsgBox Sql, vbExclamation
        Exit Sub
    End If
    
    ' cogemos el nro de asiento dependiendo de la fecha
    Dim Mc As Contadores
    
    Set Mc = New Contadores
    If Mc.ConseguirContador(0, txtFec(2).Text <= vParam.fechafin, False) = 0 Then
        NA = Mc.Contador
    Else
        MsgBox "Error al obtener contador", vbExclamation
        Set Mc = Nothing
        Exit Sub
    End If
    Set Mc = Nothing
    
    'Ahora generemos la cabecera de apunte
    Screen.MousePointer = vbHourglass
    If GenerarCabecera(NA) Then
        CadenaDesdeOtroForm = ""
        If Text4.Text <> "" Then
            frmAsientosHco.DesdeNorma43 = 2
        Else
            frmAsientosHco.DesdeNorma43 = 1
        End If
        frmAsientosHco.Asiento = Text9.Text & "|" & txtFec(2).Text & "|" & NA & "|"
        frmAsientosHco.Show vbModal
    End If
    
    ' si el asiento está descuadrado hemos de eliminarlo
    Sql = "select sum(coalesce(timported,0) - coalesce(timporteh,0)) from hlinapu where numasien = " & DBSet(NA, "N") & " and numdiari = " & DBSet(Text9.Text, "N") & " and fechaent = " & DBSet(txtFec(2).Text, "F")
    Sql1 = "select count(*) from hlinapu where numasien = " & NA & " and numdiari = " & DBSet(Text9.Text, "N") & " and fechaent = " & DBSet(txtFec(2).Text, "F")
    If DevuelveValor(Sql) <> 0 Or DevuelveValor(Sql1) = 0 Then
        'Borramos las lineas del apunte
        Screen.MousePointer = vbHourglass
        Sql = "Delete from hlinapu where numasien = " & NA & " and numdiari = " & DBSet(Text9.Text, "N") & " and fechaent = " & DBSet(txtFec(2).Text, "F")
        Conn.Execute Sql
        Sql = "Delete from hcabapu where numasien = " & NA & " and numdiari = " & DBSet(Text9.Text, "N") & " and fechaent = " & DBSet(txtFec(2).Text, "F")
        Conn.Execute Sql
    
        'devolvemos el contador
        Set Mc = New Contadores
        Mc.DevolverContador 0, txtFec(2).Text <= vParam.fechafin, NA
        Set Mc = Nothing
    
    Else
    
        'Actualiz importes y demas
        
        HaGeneradoAsiento
        
    End If
    Me.FrameGenera.visible = False
    Me.FrameDatos.Enabled = True
    Frame1.Enabled = True
    Me.FrameIntro.Enabled = True
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub HaGeneradoAsiento()
    
        'Aumentamos los importes punteados
        Importe = CCur(ListView1.SelectedItem.SubItems(1))
        De = De + Importe
        Ha = Ha + Importe
        PonerImportes
    
        'Puntemos el extracto
        Sql = "UPDATE norma43 SET punteada= 1 WHERE codigo=" & ListView1.SelectedItem.Tag
        Conn.Execute Sql
    
        'Para buscarlo
        NumRegElim = ListView1.SelectedItem.Tag
        'Volvemos a cargar todo
        Screen.MousePointer = vbHourglass
        CargarDatosLw False
        'Volvemos a siutar el select item
        For I = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(I).Tag = NumRegElim Then
                
                Set ListView1.SelectedItem = ListView1.ListItems(I)
                            
                ListView1.SelectedItem.EnsureVisible
                ListView1_DblClick
                    
                Exit For
            End If
        Next I
        
        If ListView1.SelectedItem.Index < ListView1.ListItems.Count - 3 Then ListView1.ListItems(ListView1.SelectedItem.Index + 2).EnsureVisible
        
        
End Sub


Private Sub cmdAtoCancelar_Click()
    Me.FrameGenera.visible = False
    Me.FrameDatos.Enabled = True
    Frame1.Enabled = True
    Me.FrameIntro.Enabled = True
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub



Private Sub PonerTamanyo(Punteo As Boolean)
    Me.FrameDatos.visible = Punteo
    Me.FrameIntro.visible = Not Punteo
    If Punteo Then
        Me.Height = FrameDatos.Height + 400
        Me.Width = FrameDatos.Width + 100
        If Screen.Width > 12300 Then
            Me.top = 800
            Me.Left = 800
        Else
            Me.top = 0
            Me.Left = 0
        End If
    
    Else
        Me.Height = FrameIntro.Height + 400
        Me.Width = FrameIntro.Width + 100
        If Screen.Width > 12300 Then
            Me.top = 4000
            Me.Left = 4000
        Else
            Me.top = 1000
            Me.Left = 1000
        End If
    End If
          
End Sub





Private Sub CrearAsiento()
    'Crear asiento
    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
    If ListView1.SelectedItem.Checked Then
        MsgBox "Extracto ya esta punteado", vbExclamation
        Exit Sub
    End If
    
     
    If Text1.Text = "" Then
        MsgBoxA "Cuenta banco VACIA", vbExclamation
        PonFoco Text4
        Exit Sub
    End If
    
    
    
    
    'Deshabilitamos
    Me.FrameDatos.Enabled = False
    Frame1.Enabled = False
    Me.FrameIntro.Enabled = False
    
    'Limpiamos y ponemos datos
    Me.txtFec(2).Text = Format(ListView1.SelectedItem.Text, "dd/mm/yyyy")
    
    'dIARIO POR DEFECTO DE PARAMETROS
    'Veremos si hay parametros
    Sql = DevuelveDesdeBD("diario43", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
    Text9.Text = Sql
    If Text9.Text <> "" Then Sql = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text9.Text, "N")
    Text10.Text = Sql
    
    'Concepto por defecto desde parametros
    Sql = DevuelveDesdeBD("conce43", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
    Text6.Text = Sql
    If Text6.Text <> "" Then Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text6.Text, "N")
    Text7.Text = Sql
    
    'La ampliacion del concepto viene del extracto bancario
    Text8.Text = ListView1.SelectedItem.SubItems(4)
    
    Text4.Text = "": Text5.Text = ""
    Text11.Text = ""
    Label1(5).Caption = Label1(4).Caption
    'Ponemos visible
    Me.FrameGenera.visible = True
    'Ponemos el foco en doc
    Text11.SetFocus
    
End Sub




Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub Command1_Click()
    If EstaLW1 Then
        ListView1_DblClick
    Else
        ListView2_DblClick
    End If
End Sub

Private Sub cmdImportar_Click(Index As Integer)
    Text12.Text = Trim(Text12.Text)
    If Text12.Text = "" Then
        MsgBox "Debes indicar un archivo", vbExclamation
        Exit Sub
    End If
    If Dir(Text12.Text, vbArchive) = "" Then
        MsgBox "Archivo NO existe", vbExclamation
        Exit Sub
    End If
    'Borramos los temporales
    Sql = "Delete from tmpnorma43 where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    Screen.MousePointer = vbHourglass
    If ProcesarFichero Then
        NumRegElim = 1
        'Ahora procesamos los datos
        ProcesarDatos
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdReplace_Click()
            
            
    cad = InputBox("Sustituir:", "")
    If cad = "" Then Exit Sub
    Sql = Trim(InputBox(".. por:", ""))
        
    If MsgBox("Sustituir " & cad & " por " & IIf(Sql = "", "##VACIO##", Sql) & vbCrLf & "¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    txtDatos.Text = Replace(txtDatos, cad, Sql)
    cad = "UPDATE tmpnorma43 set concepto=replace(concepto," & DBSet(cad, "T") & "," & DBSet(Sql, "T") & ") where codusu =" & vUsu.Codigo
    Conn.Execute cad
End Sub



Private Sub Form_Load()

    Me.Icon = frmppal.Icon
    
    'La toolbar
    With Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 44
        .Buttons(2).Image = 1
        .Buttons(4).Image = 37
        .Buttons(4).Enabled = VerCobrosPagos
        .Buttons(4).visible = VerCobrosPagos
        
        .Buttons(5).Image = 38
        .Buttons(5).Enabled = VerCobrosPagos
        .Buttons(5).visible = VerCobrosPagos
        
        
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    imgCuentas.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Image1.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Image2.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Image3.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Image4.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    

    FrameGenera.visible = False
    FrameIntro.Enabled = True
    FrameBotonGnral.Enabled = False
    Frame1.Enabled = True
    Frame2.visible = False
    
    cmdReplace.visible = vUsu.Login = "root"
    Text1.Text = ""
    Text2.Text = ""
    txtFec(0).Text = ""
    txtFec(1).Text = ""
    Label11.Caption = Label11.Tag
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then
        If FrameGenera.visible Then
            cmdAtoCancelar_Click
        Else
            Unload Me
        End If
    End If
End Sub





Private Sub Form_Unload(Cancel As Integer)
    'Desbloqueamos
    BloqueoManual False, "PUNTEOB", Text1.Text
End Sub



Private Sub frmC_Selec(vFecha As Date)
    txtFec(CInt(txtFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    Text4.Text = RecuperaValor(CadenaSeleccion, 1)
    Text5.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCo_DatoSeleccionado(CadenaSeleccion As String)
    Text6.Text = RecuperaValor(CadenaSeleccion, 1)
    Text7.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    Text1.Text = RecuperaValor(CadenaSeleccion, 1)
    Text2.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    Text9.Text = RecuperaValor(CadenaSeleccion, 1)
    Text10.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click()
    Set frmD = New frmTiposDiario
    frmD.DatosADevolverBusqueda = "0|1|"
    frmD.Show vbModal
    Set frmD = Nothing
End Sub

Private Sub Image2_Click()
    Set frmCo = New frmConceptos
    frmCo.DatosADevolverBusqueda = "0|1|"
    frmCo.Show vbModal
    Set frmCo = Nothing
End Sub

Private Sub Image3_Click()
    Set frmCC = New frmColCtas
    frmCC.DatosADevolverBusqueda = "0|1"
    frmCC.ConfigurarBalances = 3  'NUEVO
    frmCC.Show vbModal
    Set frmCC = Nothing
End Sub

Private Sub Image4_Click()

    cd1.CancelError = False
    cd1.DialogTitle = "Archivo banco NORMA 43"
    cd1.ShowOpen
    If cd1.FileName <> "" Then Text12.Text = cd1.FileName
    
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim Puntear As Boolean

    cad = ""
    If Index < 2 Then
        'En apuntes
        If Me.ListView2.ListItems.Count = 0 Then cad = "No hay apuntes para realizar la accion"
    
    Else
        If Me.ListView1.ListItems.Count = 0 Then cad = "No hay datos en el extracto de banco para realizar la accion"
    End If
    If cad <> "" Then
        MsgBox cad, vbExclamation
        Exit Sub
    End If
    
    Puntear = (Index = 0 Or Index = 2)
    cad = IIf(Puntear, "puntear", "quitar el punteo a")
    
    If Index < 2 Then
        cad = "¿Desaea  " & cad & " las lineas de apuntes visualizadas?"
    Else
        cad = "¿Desaea  " & cad & " los extractos del banco visualizados?"
    End If
    
    NF = CInt(MsgBox(cad, vbQuestion + vbYesNoCancel))
    If NF <> vbYes Then Exit Sub
    
    
    
    'AUQI empieza la fiesta
    Screen.MousePointer = vbHourglass
    If Index >= 2 Then
        'En extracto bancario
        HacerPunteoAutomatico Me.ListView1, True, Puntear
    Else
        HacerPunteoAutomatico Me.ListView2, False, Puntear
    End If
    
    Label11.Caption = "Cargando datos"
    Label11.Refresh
    
    CargarDatosLw True
        
    FicheroPpal = ""
    Label11.Caption = Label11.Tag
    
    
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerPunteoAutomatico(ByRef LW As ListView, N43 As Boolean, Puntear As Boolean)

    NF = IIf(Puntear, 1, 0)

    For I = 1 To LW.ListItems.Count
        Label11.Caption = LW.ListItems(I).SubItems(4)
        Label11.Refresh
        cad = ""
        If Puntear Then
            'Si ya esta punteado no hago nada
            If LW.ListItems(I).Checked Then cad = "N"
        Else
            'Si NO esta punteado no hago nada
            If Not LW.ListItems(I).Checked Then cad = "N"
        End If
        
        If cad = "" Then
            If N43 Then
            
                cad = "UPDATE norma43 SET punteada= " & NF & " WHERE codigo=" & LW.ListItems(I).Tag
            
            
            Else
                cad = "UPDATE hlinapu SET "
                cad = cad & " punteada = " & NF
                cad = cad & " WHERE fechaent='" & Format(LW.ListItems(I).Text, FormatoFecha) & "'"
                cad = cad & " AND numasien=" & RecuperaValor(LW.ListItems(I).Tag, 1)
                cad = cad & " AND numdiari =" & RecuperaValor(LW.ListItems(I).Tag, 2)
                cad = cad & " AND linliapu =" & RecuperaValor(LW.ListItems(I).Tag, 3)
            End If
            
            Ejecuta cad, False
            
        End If
        
    Next
End Sub


Private Sub imgCuentas_Click()
    
    Set frmCta = New frmBasico2
    AyudaCuentasBancarias frmCta
    Set frmCta = Nothing
    
    PonerFoco Text1
End Sub

Private Sub imgppal_Click(Index As Integer)
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    txtFec(0).Tag = Index
    If txtFec(Index).Text <> "" Then
        If IsDate(txtFec(Index).Text) Then frmC.Fecha = CDate(txtFec(Index).Text)
    End If
    frmC.Show vbModal
    Set frmC = Nothing
End Sub


Private Sub ListView1_Click()
    EstaLW1 = True
End Sub

Private Sub ListView1_DblClick()
Dim J As Integer
Dim Find As Boolean
Dim Fin As Long

    EstaLW1 = True
    If ListView2.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    J = ListView2.SelectedItem.Index
    Find = False
    Fin = ListView2.ListItems.Count
    Do
        For I = J To Fin
            If ListView2.ListItems(I).SubItems(2) = ListView1.SelectedItem.SubItems(1) Then
                If ListView2.ListItems(I).SubItems(3) <> ListView1.SelectedItem.SubItems(2) Then
                    'Ha encontrado con el mismo importe y signos distintos D-H
                    Set ListView2.SelectedItem = ListView2.ListItems(I)
                    ListView2.SelectedItem.EnsureVisible
                    Find = True
                    Exit For
                End If
            End If
        Next I
        If Not Find Then
            If J > 1 Then
                Fin = J
                J = 1
            Else
                Find = True
            End If
        End If
                
    Loop Until Find
End Sub


Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        If VerCobrosPagos Then
            'C: cobro
            'V: Pago
            'B: Desde cuenta
            If KeyCode = vbKeyC Or KeyCode = vbKeyB Then VerVencimiento KeyCode = vbKeyC, False
            If KeyCode = vbKeyV Then VerVencimiento True, True
        End If
    End If
End Sub

Private Sub ListView2_Click()
    EstaLW1 = False
End Sub

Private Sub ListView2_DblClick()
Dim J As Integer
Dim Find As Boolean
Dim Fin As Long

    EstaLW1 = False
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView2.SelectedItem Is Nothing Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then
        J = 0
    Else
        J = ListView1.SelectedItem.Index + 1
    End If
    Find = False
    Fin = ListView1.ListItems.Count
    Do
        For I = J To Fin
            If ListView1.ListItems(I).SubItems(1) = ListView2.SelectedItem.SubItems(2) Then
                If ListView1.ListItems(I).SubItems(2) <> ListView2.SelectedItem.SubItems(3) Then
                    'Ha encontrado con el mismo importe y signos distintos D-H
                    Set ListView1.SelectedItem = ListView1.ListItems(I)
                    ListView1.SelectedItem.EnsureVisible
                    Find = True
                    Exit For
                End If
            End If
        Next I
        If Not Find Then
            If J > 1 Then
                Fin = J
                J = 1
            Else
                Find = True
            End If
        End If
                
    Loop Until Find

End Sub

Private Sub Text1_GotFocus()
    PonFoco Text1
    CuentaAnterior = Text1.Text
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 107 Or KeyCode = 187 Then
        KeyCode = 0
        Text1.Text = ""
        imgCuentas_Click
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus()
Dim RC As String
    
    If FrameGenera.visible Then Exit Sub
    Text1.Text = Trim(Text1.Text)
    If Text1.Text = "+" Then Text1.Text = ""
    If Text1.Text = "" Then
        Text2.Text = ""
        Exit Sub
    Else
         RC = Text1.Text
         If CuentaCorrectaUltimoNivel(RC, Sql) Then
             Text1.Text = RC
             Text2.Text = Sql
             
             ConfirmarDatos True
             CuentaAnterior = Text1.Text
         Else
             MsgBox Sql, vbExclamation
             Text2.Text = ""
         End If
         If Text2.Text = "" Then PonerFoco Text1
         
    End If
             
End Sub


Private Sub PonerFoco(Obj As Object)
    On Error Resume Next
    Obj.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub Text11_GotFocus()
    PonFoco Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub



Private Sub Text4_GotFocus()
    PonFoco Text4
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 107 Or KeyCode = 187 Then
        KeyCode = 0
        Text4.Text = ""
        Image3_Click
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text4_LostFocus()
Dim RC As String

    Text4.Text = Trim(Text4.Text)
    If Text4.Text = "+" Then Text4.Text = ""
    If Text4.Text = "" Then
        Text5.Text = ""
    Else
        RC = Text4.Text
        If CuentaCorrectaUltimoNivel(RC, Sql) Then
            Text4.Text = RC
            Text5.Text = Sql
        Else
            MsgBox Sql, vbExclamation
            Text5.Text = ""
            Text4.Text = ""
            Text4.SetFocus
        End If
    End If
End Sub



Private Sub Text6_GotFocus()
    PonFoco Text6
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text6_LostFocus()
   With Text6
        .Text = Trim(.Text)
        I = 1
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "El valor debe ser numérico: " & .Text, vbExclamation
            Else
                 If Val(.Text) >= 900 Then
                    MsgBox "Los conceptos superiores a 900 se los reserva la aplicación.", vbExclamation
                Else
                    Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", .Text, "N")
                    If Sql = "" Then
                        MsgBox "Concepto NO encontrado: " & .Text, vbExclamation
                    Else
                        Text7.Text = Sql
                        I = 0
                    End If
                End If
            End If
        Else
            'Igual a "" luego pasamos a otro campo en la tabulacion
            I = 2
        End If
        If I > 0 Then
            .Text = ""
            Text7.Text = ""
            If I = 1 Then Text6.SetFocus
        End If
    End With
End Sub

Private Sub Text8_GotFocus()
    PonFoco Text8
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text9_GotFocus()
    PonFoco Text9
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text9_LostFocus()
    With Text9
        .Text = Trim(.Text)
        I = 1
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "El valor debe ser numérico: " & .Text, vbExclamation
            Else
                Sql = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", .Text, "N")
                If Sql = "" Then
                    MsgBox "Concepto NO encontrado: " & .Text, vbExclamation
                Else
                    Text10.Text = Sql
                    I = 0
                End If
            End If
        Else
            'Igual a "" luego pasamos a otro campo
            I = 2
        End If
        If I > 0 Then
            .Text = ""
            Text10.Text = ""
            If I = 1 Then Text9.SetFocus
        End If
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Index = 1 Then
        CrearAsiento
    Else
        If Button.Index = 2 Then
           VerAsiento
        Else
            'Tiene submenu
           
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
   Dim Opci As Byte
    
    If Mid(ButtonMenu.Key, 1, 1) = "G" Then
    
        Opci = Val(Mid(ButtonMenu.Key, 2, 1))
        If Opci = 1 Or Opci = 2 Then
            VerVencimiento Mid(ButtonMenu.Key, 2, 1) <> "2", False
        Else
            If Opci = 4 Then VerVencimiento True, True
        End If
              
        
    Else
        
            
        If Mid(ButtonMenu.Key, 2, 1) <> "2" Then
            frmTESCobros.Show vbModal
        Else
            frmTESPagos.Show vbModal
        End If
            
    End If
    
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select

End Sub

Private Sub txtfec_GotFocus(Index As Integer)
    PonFoco txtFec(Index)
    FechaAnterior = txtFec(Index).Text
End Sub
'++
Private Sub txtfec_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYFecha KeyAscii, 0
            Case 1: KEYFecha KeyAscii, 1
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgppal_Click (Indice)
End Sub

'++

Private Sub txtfec_LostFocus(Index As Integer)
Dim Mal As Boolean
    txtFec(Index).Text = Trim(txtFec(Index).Text)
    Mal = True

    If txtFec(Index).Text = "" Then Exit Sub

        If Not EsFechaOK(txtFec(Index)) Then
            MsgBox "No es una fecha correcta", vbExclamation
        Else
            Mal = False
        End If
    If Mal Then
        PonerFoco txtFec(Index)
    Else
        If txtFec(Index).Text <> FechaAnterior Then ConfirmarDatos True
    End If
    
End Sub



Private Sub CargaBancario()

    ListView1.ListItems.Clear
    Sql = "Select * from norma43 where"
    Sql = Sql & " codmacta ='" & Text1.Text & "'"
    Sql = Sql & " AND fecopera >='" & Format(txtFec(0).Text, FormatoFecha) & "'"
    Sql = Sql & " AND fecopera <='" & Format(txtFec(1).Text, FormatoFecha) & "'"
    'OCultar/mostrar punteados
    If Check1.Value = 0 Then
        'Ocultar los ya puntedos
        Sql = Sql & " AND Punteada = 0 "
    End If
    Sql = Sql & " ORDER BY fecopera,codigo"

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add()
        ItmX.Text = Format(Rs!fecopera, "dd/mm/yyyy")
        'Importe Debe
        If Not IsNull(Rs!ImporteD) Then
            Importe = Rs!ImporteD
            Sql = "D"
        Else
            'Importe HABER
            If Not IsNull(Rs!ImporteH) Then
                Importe = Rs!ImporteH
                Sql = "H"
            Else
                Sql = "XX"
            End If
        End If
        ItmX.SubItems(1) = Format(Importe, FormatoImporte)
        ItmX.SubItems(2) = Sql
        ItmX.SubItems(3) = Format(Rs!Saldo, FormatoImporte)
        ItmX.SubItems(4) = Rs!Concepto
        ItmX.ListSubItems(4).ToolTipText = DBLet(Rs!Concepto, "T")
        
        ItmX.Tag = Rs!Codigo
        ItmX.Checked = (Rs!punteada = 1)
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
End Sub


Private Sub CargaLineaApuntes()

    ListView2.ListItems.Clear
    Sql = "Select numasien,fechaent,numdiari,linliapu,ampconce,timported,timporteh,punteada,saldo FROM tmpconext"
    Sql = Sql & " WHERE codusu = " & vUsu.Codigo
    
    If Check1.Value = 0 Then
        'Ocultar los ya puntedos
        Sql = Sql & " AND Punteada = '' "
    End If
    Sql = Sql & " ORDER BY pos"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Set ItmX = ListView2.ListItems.Add()
        ItmX.Text = Format(Rs!FechaEnt, "dd/mm/yyyy")
        'Importe Debe
        Sql = " "
        If Not IsNull(Rs!timported) Then
            Importe = Format(Rs!timported, FormatoImporte)
            Sql = "D"
        Else
            'Importe HABER
            If Not IsNull(Rs!timporteH) Then
                Importe = Rs!timporteH
                Sql = "H"
            Else
                Importe = 0
                Sql = "XX"
            End If
        End If
        ItmX.SubItems(1) = Format(Rs!NumAsien, "0000")
        ItmX.SubItems(2) = Format(Importe, FormatoImporte)
        ItmX.SubItems(3) = Sql
        ItmX.SubItems(4) = Format(Rs!Saldo, FormatoImporte)
        ItmX.SubItems(5) = DBLet(Rs!Ampconce, "T")
        ItmX.ListSubItems(5).ToolTipText = DBLet(Rs!Ampconce, "T")

        
        ItmX.Tag = Rs!NumAsien & "|" & Rs!NumDiari & "|" & Rs!Linliapu & "|"
        ItmX.Checked = (Rs!punteada <> "")
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub




'----------------- PUNTEOS

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
EstaLW1 = True
Screen.MousePointer = vbHourglass
    Set ListView1.SelectedItem = Item
    'Ponemos a true o a false
    PunteaEnBD Item, True
    'Comprobamos primero que esta checkeado
    If Item.Checked = False Then
        PrimeraSeleccion = True
        ClickAnterior = 0
    Else
        If ClickAnterior <> 1 Then
            If PrimeraSeleccion Then
                BusquedaEnHaber
                PrimeraSeleccion = False
                ClickAnterior = 1
            Else
                PrimeraSeleccion = True
                ClickAnterior = 0
            End If
        Else
            PrimeraSeleccion = True
        End If
    
    End If
    Screen.MousePointer = vbDefault

End Sub


Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Screen.MousePointer = vbHourglass
    EstaLW1 = False
    Set ListView2.SelectedItem = Item
    'Ponemos a true o a false
    PunteaEnBD Item, False
    'Comprobamos primero que esta checkeado
    If Item.Checked = False Then
        PrimeraSeleccion = True
        ClickAnterior = 0
    Else
        If ClickAnterior <> 2 Then
            If PrimeraSeleccion Then
                BusquedaEnDebe
                PrimeraSeleccion = False
                ClickAnterior = 2
            Else
                PrimeraSeleccion = True
                ClickAnterior = 0
            End If
        Else
            PrimeraSeleccion = True
        End If
    
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Sub BusquedaEnHaber()
    ContadorBus = 1
    Checkear = False
    Do
        I = 1
        While I <= ListView2.ListItems.Count
            'Comprobamos k no esta chekeado
            If Not ListView2.ListItems(I).Checked Then
                'K tiene el mismo importe
                If ListView1.SelectedItem.SubItems(1) = ListView2.ListItems(I).SubItems(2) Then
                    'K no sean DEBE o HABER los dos
                    Checkear = (ListView1.SelectedItem.SubItems(2) <> ListView2.ListItems(I).SubItems(3))

                    If Checkear Then
                        'Tiene el mismo importe y no esta chequeado
                        Set ListView2.SelectedItem = ListView2.ListItems(I)
                        ListView2.SelectedItem.EnsureVisible
                        ListView2.SetFocus
                        Beep
                        Exit Sub
                    End If
                End If
            End If
            I = I + 1
        Wend
        ContadorBus = ContadorBus + 1
        Loop Until ContadorBus > 2
End Sub



Private Sub BusquedaEnDebe()
    ContadorBus = 1
    Checkear = False
    Do
        I = 1
        While I <= ListView1.ListItems.Count
            If ListView2.SelectedItem.SubItems(2) = ListView1.ListItems(I).SubItems(1) Then
                'Lo hemos encontrado. Comprobamos que no esta chequeado
                If Not ListView1.ListItems(I).Checked Then
                    'Tiene el mismo importe y no son debe o haber
                    Checkear = (ListView2.SelectedItem.SubItems(3) <> ListView1.ListItems(I).SubItems(2))

                    If Checkear Then
                        Set ListView1.SelectedItem = ListView1.ListItems(I)
                        ListView1.SelectedItem.EnsureVisible
                        ListView1.SetFocus
                        Beep
                        Exit Sub
                    End If
                End If
            End If
            I = I + 1
        Wend
        ContadorBus = ContadorBus + 1
    Loop Until ContadorBus > 2
End Sub



Private Sub PunteaEnBD(ByRef IT As ListItem, EnDEBE As Boolean)
Dim RC As String
On Error GoTo EPuntea
    
    
    If Not EnDEBE Then
        'ASientos
        'Actualizamos en DOS tablas, en la tmp y en la hcoapuntes
        Sql = "UPDATE hlinapu SET "
        If IT.Checked Then
            RC = "1"
            Importe = 1
            Else
            RC = "0"
            Importe = -1
        End If
        Importe = Importe * CSng(IT.SubItems(2))
        If EnDEBE Then
            De = De + Importe
        Else
            Ha = Ha + Importe
        End If
        Sql = Sql & " punteada = " & RC
        Sql = Sql & " WHERE fechaent='" & Format(IT.Text, FormatoFecha) & "'"
        Sql = Sql & " AND numasien="
        RC = RecuperaValor(IT.Tag, 1)
        Sql = Sql & RC & " AND numdiari ="
        RC = RecuperaValor(IT.Tag, 2)
        Sql = Sql & RC & " AND linliapu ="
        RC = RecuperaValor(IT.Tag, 3)
        Sql = Sql & RC
        
        
        
        
    Else
        'En Norma 43
        
        If IT.Checked Then
            RC = "1"
            Importe = 1
            Else
            RC = "0"
            Importe = -1
        End If
        Importe = Importe * CSng(IT.SubItems(1))
        If EnDEBE Then
            De = De + Importe
        Else
            Ha = Ha + Importe
        End If
        Sql = "UPDATE norma43 SET punteada= " & RC & " WHERE codigo=" & IT.Tag
        
    End If
    
    Conn.Execute Sql
    
    'Ponemos los importes
    PonerImportes

    
    Exit Sub
EPuntea:
    MuestraError Err.Number, "Accediendo BD para puntear", Err.Description
End Sub

Private Sub PonFoco(ByRef T As TextBox)
    T.SelStart = 0
    T.SelLength = Len(T.Text)
End Sub

Private Function GenerarCabecera(NumAsi As Long) As Boolean
Dim cad As String

    On Error GoTo EGenerarCabecera
    GenerarCabecera = False
    
    '-------------------------------------------------------------------------
    'Insertamos cabecera
    Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES ("
    'Ejemplo
    ' 1, '2003-11-25', 1, 1, NULL, 'misobs')
    Sql = Sql & Text9.Text & ",'" & Format(CDate(txtFec(2).Text), FormatoFecha) & "'," & NumAsi & ","
    'Observaciones
    Sql = Sql & "'Asiento generado desde punteo bancario por " & vUsu.Nombre & " el " & Format(Now, "dd/mm/yyyy") & "',"
    '
    Sql = Sql & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Punteo Bancario')"
    Conn.Execute Sql
    
    '-----------------------------------------------------------------------------
    'La linea del asiento
    'Hemos puesto hlinapu mas atras para poder cambiarla
    Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, numdocum,"
    Sql = Sql & " ampconce, codconce, linliapu, codmacta, timporteD, timporteH, ctacontr, codccost, idcontab, punteada) VALUES ("
    
    'Ejemplo valores
    '1, '2001-01-20', 0, 0, '0', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0)"
    Sql = Sql & Text9.Text & ",'" & Format(CDate(txtFec(2).Text), FormatoFecha) & "'," & NumAsi & ","
    '          dcumento
    Sql = Sql & DBSet(Text11.Text, "T") & ","
    
    'Ampliacion concepto
    cad = Mid(Text7.Text & " " & Text8.Text, 1, 30)
    Sql = Sql & DBSet(cad, "T") & ","
    
    'Concepto
    Sql = Sql & Text6.Text & ","
    
    'El importe
    Importe = CCur(ListView1.SelectedItem.SubItems(1))
    cad = "1,'" & Text1.Text & "',"
    If ListView1.SelectedItem.SubItems(2) = "H" Then
        'Va al debe
        cad = cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
    Else
        cad = cad & "NULL," & TransformaComasPuntos(CStr(Importe))
    End If
    
    'Contrapartida
    If Text4.Text <> "" Then
        cad = cad & ",'" & Text4.Text & "'"
    Else
        cad = cad & ",NULL"
    End If
    
    'y la punteamos
    cad = Sql & cad & ",NULL,'CONTAB',1)"
    Conn.Execute cad
    
    'Si tiene contrapartida entonces genero la segunda linea de apuntes
    ' k sera la de la contrapartida, con el importe el mismo al lado contrario
    ' el mismo concepto
    If Text4.Text <> "" Then
        'SI TIENE
            cad = "2,'" & Text4.Text & "',"
            'En la de arriba es igual a H
            If ListView1.SelectedItem.SubItems(2) = "D" Then
                'Va al debe
                cad = cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
            Else
                cad = cad & "NULL," & TransformaComasPuntos(CStr(Importe))
            End If
            
            'Contrapartida es la del banco
            cad = cad & ",'" & Text1.Text & "'"
            
            'y NO la punteamos
            cad = Sql & cad & ",NULL,'CONTAB',0)"
            Conn.Execute cad
    End If
    GenerarCabecera = True
    Exit Function
EGenerarCabecera:
    MuestraError Err.Number, Err.Description
End Function



Private Sub PonerImportes()

    If De <> 0 Then
        Text3(0).Text = Format(De, FormatoImporte)
        Else
        Text3(0).Text = ""
    End If
    If Ha <> 0 Then
        Text3(1).Text = Format(Ha, FormatoImporte)
        Else
        Text3(1).Text = ""
    End If
    Importe = De - Ha
    If Importe <> 0 Then
        Text3(2).Text = Format(Importe, FormatoImporte)
        Else
        Text3(2).Text = ""
    End If
End Sub

'############################################################
'  PARTE CORRESPONDIENTE A LA IMPORTACION DE DATOS NORMA 34
'############################################################

Private Function ProcesarFichero() As Boolean
Dim Fin As Boolean
Dim cad As String

On Error GoTo EProcesarFichero
    'Abrimos el fichero para lectura
    ProcesarFichero = False
    NF = FreeFile
    FicheroPpal = "|"
    Open Text12.Text For Input As #NF
    While Not EOF(NF)
        Line Input #NF, Sql
        
        If Len(FicheroPpal) = 1 Then
            'Primera linea.
            'A veces, el fichhero esta grabado en UTF8
            'Par ello quitaremos los tres primeros caracteres, o hasta llegar a un NUMERO
            K = 0
            cad = "OK"
            For I = 1 To Len(Sql)
                J = Asc(Mid(Sql, I, 1))
                If J > 90 Then
                    'nada
                    cad = "N"
                Else
                    cad = ""
                    K = I
                    Exit For
                End If
                
            Next
            
            If K > 1 Then Sql = Mid(Sql, K)
                
   
        End If
        
        If InStr(1, Sql, vbCrLf) = 0 Then
            'No hay salto de linea y return
            If InStr(1, Sql, vbLf) > 0 Then
                If K = 1 Then
                    'Solo tiene una linea, y viene con saltos de linea vbLF
                    Sql = Replace(Sql, vbLf, "|")
                End If
            End If
        End If
        
        If Sql <> "" Then
                                        'Separador de lineas
            FicheroPpal = FicheroPpal & Sql & "|"
        End If
    Wend
    Close #NF
    ProcesarFichero = True
    Exit Function
EProcesarFichero:
    MuestraError Err.Number
End Function


Private Sub ProcesarDatos()
Dim I As Long
Dim CONT As Long
Dim NF As Long
Dim Linea As String
Dim Fichero As String
Dim Primer23 As Boolean
Dim Num22 As Integer  'Para conrolar los asientos k se han realizado
Dim Ampliacion As String
Dim RegistroInsertado As Boolean
Dim Comienzo As Long   'Para cuando vienen varios bancos
Dim Fecha As String   'Fecha importacion datos

Dim ContadorMYSQL As Integer
Dim ContadorRegistrosBanco As Integer
Dim Mas_Observaciones As Boolean
Dim MiAux As String



    'Vemos cuantas cuentas trae el extracto
    I = 0
    CONT = 0
    Do
        NF = I + 1
        I = InStr(NF, FicheroPpal, "|11")  'los registros empiezan por 11 para las cuentas
        If I > 0 Then CONT = CONT + 1
    Loop Until I = 0
        
    If CONT = 0 Then
        MsgBox "Error en el fichero. No se ha encontrado registro 11", vbExclamation
        Exit Sub
    End If

    
    
    txtDatos.Text = ""
    txtSald(0).Text = ""
    txtSald(1).Text = ""
    txtSald(1).BackColor = Me.txtSald(0).BackColor
    Comienzo = 2
    ContadorMYSQL = 1
    ContadorRegistrosBanco = 0
    Cta = ""
    'Ya sabemos cuantas cont hay k tratar
    For I = 1 To CONT
        If I <> CONT Then
            Linea = "|11"
            'Hay mas de un |11 o cuenta bancaria
        Else
            'Una unica cta bancaria en este fichero
            Linea = "|88"
        End If
        
        NF = InStr(Comienzo, FicheroPpal, Linea)
        If NF = 0 Then
            MsgBox "Imposible situar datos.", vbExclamation
            Exit Sub
        End If
        
        Fichero = Mid(FicheroPpal, Comienzo, NF - 1)
        
        Comienzo = NF + 1
                
        'Fecha
        Fecha = ""
        Linea = Mid(Fichero, 31, 2) & "/" & Mid(Fichero, 29, 2) & "/" & Mid(Fichero, 27, 2)
        If IsDate(Linea) Then
            Fecha = "Fecha: " & Space(18) & Format(Linea, "dd/mm/yyyy")
        Else
            Fecha = "Fecha: " & Space(18) & "Error obteniendo fecha"
        End If
        Fecha = Fecha & vbCrLf
                
        'ANTES
        NF = InStr(1, Fichero, "|") 'Es el fin de la primera linea
        
        'Primara linea, la de la cuenta
        Linea = Mid(Fichero, 1, NF - 1) 'pq quitamos el pipe del principio y del final
        
        'De la primera linea obtenemos el numero de cuenta
        Ampliacion = Cta
        FijarCtaContable (Linea)
        If Ampliacion <> Cta Then
            If Ampliacion <> "" Then
                'HA CAMBIADO DE CUENTA DEEEENTRO DEL MISMO Fichero
                ContadorRegistrosBanco = 0
            End If
        End If
        
        If Cta = "" Then
            
            MsgBox "Error obteniendo la cuenta contable asociada. Linea: " & Linea, vbExclamation
            Exit Sub
        Else
            Sql = ""
            If ContadorRegistrosBanco = 0 Then
                If txtDatos.Text <> "" Then txtDatos.Text = txtDatos.Text & Sql & vbCrLf
                For NF = 1 To 98
                    Sql = Sql & "="
                Next NF
                txtDatos.Text = txtDatos.Text & Sql & vbCrLf
                Sql = Mid(Linea, 3, 4) & " " & Mid(Linea, 7, 4) & " ** " & Mid(Linea, 11, 10)
                txtDatos.Text = txtDatos.Text & "Cuenta bancaria: " & Sql & vbCrLf
                Fecha = Fecha & "Cuenta bancaria:   " & Sql & vbCrLf
                txtDatos.Text = txtDatos.Text & "Cuenta contable:   " & Cta & vbCrLf
                Fecha = Fecha & "Cuenta contable:    " & Cta & vbCrLf
                txtDatos.Text = txtDatos.Text & "Linea  F.Opercion   F.Valor         Debe            Haber          Concepto" & vbCrLf
                Sql = ""
                For NF = 1 To 98
                    Sql = Sql & "-"
                Next NF
                txtDatos.Text = txtDatos.Text & Sql & vbCrLf
            Else
                'Es otro trozo de fichero 11| pero de la misma cuenta
                txtDatos.Text = txtDatos.Text & String(98, "=") & vbCrLf
            End If
        End If
        
        'Fijaremos el saldo incial
        Sql = Mid(Linea, 34, 14)
        If Not IsNumeric(Sql) Then
            MsgBox "Error. Se esperaba un importe: " & Sql, vbExclamation
            Exit Sub
        End If
        Saldo = Val(Sql) / 100
        txtSald(1).Text = Format(Saldo, FormatoImporte)
        
        
        'ANTES 25 Noviembre
        'Se trabaja al reves
        'Signo del saldo
        If Mid(Linea, 33, 1) = "1" Then Saldo = Saldo * -1
        
        NF = InStr(1, Fichero, "|") 'Es el fin de la primera linea
        Fichero = Mid(Fichero, NF + 1) '+1 y le quito el pipe
        
        RegistroInsertado = False
        Ampliacion = ""
        Num22 = 0
        'Ya tenemos los primeros datos. Ahora a por los apuntes
        Do
            NF = InStr(1, Fichero, "|")
            Linea = Mid(Fichero, 1, NF - 1)
            Fichero = Mid(Fichero, NF + 1)
            
            Sql = Mid(Linea, 1, 2)
          
            
            If Sql = "22" Then
                If Num22 > 0 Then
                    If Not RegistroInsertado Then
                        If Not InsertarRegistro(Ampliacion, ContadorMYSQL, ContadorRegistrosBanco) Then Exit Sub
                    End If
                End If
            
                'Primera parte de la linea de apunte
                If Not ProcesaLineaASiento(Linea, Ampliacion) Then Exit Sub
                RegistroInsertado = False
                Primer23 = True
                Num22 = Num22 + 1
            Else
                If Sql = "23" Then
                    If Primer23 Then
                        Primer23 = False
                        
                        'Mayo 2022. Veo si la siguiente linea es 23 tambien, es decir , observaciones transferencia.
                        Mas_Observaciones = False
                        Do
                            If Mid(Fichero, 1, 2) = "23" Then
                                'OK.  Es una observacion Van desde 2301, hasta 2305
                                
                                NF = InStr(1, Fichero, "|")
                                MiAux = Trim(Mid(Fichero, 1, NF - 1))
                                MiAux = Mid(MiAux, 5) 'quito el 230?
                                Linea = Linea & MiAux
                                Fichero = Mid(Fichero, NF + 1)
                                Mas_Observaciones = True
                            Else
                                Mas_Observaciones = False
                            End If
                        Loop Until Not Mas_Observaciones
                        
                        'Insertaremos
                        Ampliacion = ProcesaAmpliacion2(Linea)
                        If Not InsertarRegistro(Ampliacion, ContadorMYSQL, ContadorRegistrosBanco) Then Exit Sub
                        RegistroInsertado = True
                    End If
                    
                    
                Else
                    If Sql = "33" Then
                        If Not RegistroInsertado Then
                            If Num22 > 0 Then
                                If Not InsertarRegistro(Ampliacion, ContadorMYSQL, ContadorRegistrosBanco) Then Exit Sub
                            End If
                        End If
                        'Fin CTA. Hacer comprobaciones
                        
                        If Not HacerComprobaciones(Linea, ContadorRegistrosBanco, ContadorMYSQL) Then
                            Exit Sub
                        End If
                        Fichero = ""
                       
                    Else
                        'Cualquier otro caso no esta tratado
                        Fichero = ""
                    End If
                End If
            End If
        Loop Until Fichero = ""
        'Kitamos de ppal el valor
    Next I
    
    'Si llega aqui es k ha ido bien.Si no inserta nada, NO muestro los datos
    If ContadorMYSQL > 1 Then
        DoEvents
        Screen.MousePointer = vbHourglass
        
        Sql = "codmacta='" & Me.Text1.Text & "' and 1"
        Sql = DevuelveDesdeBD("saldo", "norma43", Sql, "1 order by codigo desc")
        If Sql <> "" Then
            txtSald(0).Text = Format(CCur(Sql), FormatoImporte)
            
            
            If txtSald(0).Text <> txtSald(1).Text Then Me.txtSald(1).BackColor = &HD6D9FE
            
        End If
        PonerModo 1
        Screen.MousePointer = vbHourglass
    End If
        
        
End Sub

Private Sub FijarCtaContable(ByRef Lin As String)
    Sql = "Select codmacta from bancos"
    Sql = Sql & " where mid(iban,5,4) = " & Mid(Lin, 3, 4) ' entidad
    Sql = Sql & " AND mid(iban,9,4) = " & Mid(Lin, 7, 4) ' oficina
    Sql = Sql & " AND mid(iban,15,10) = '" & Mid(Lin, 11, 10) & "'" ' cuentaba
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cta = ""
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then Cta = Rs.Fields(0)
    End If
    Rs.Close
    Set Rs = Nothing
    If Cta = "" Then
        Sql = "Fichero pertenece a la cuenta bancaria:  " & Mid(Lin, 3, 4) & "  " & Mid(Lin, 7, 4) & " ** " & Mid(Lin, 11, 10) & vbCrLf
        Sql = Sql & vbCrLf & "No esta asociada a ninguna cuenta contable."
        MsgBox Sql, vbExclamation
    End If
End Sub

Private Sub Command2_Click(Index As Integer)
Dim SQ As String

    If Index = 1 Then
        PonerModo 0
        Exit Sub
    End If
    
    'Comprobaremos que hay datos para traspasar
    If txtDatos.Text = "" Then
        MsgBox "Datos vacios", vbExclamation
        Exit Sub
    End If
    
    'COntamos los saltos de linea
    NumRegElim = 1
    SQ = txtDatos.Text
    NF = 0
    Do
        NumRegElim = InStr(1, SQ, vbCrLf)
        If NumRegElim > 0 Then
            SQ = Mid(SQ, NumRegElim + 2)  'vbcrlf son DOS caracteres
            NF = NF + 1
            If NF > 5 Then NumRegElim = 0 'Hay mas lineas que las del encabezado
        End If
    Loop Until NumRegElim = 0
    'Fichero comprobacion de saldos
    If NF <= 5 Then
        txtDatos.Text = ""
        If chkElimmFich.Value = 1 Then
            If Dir(Text12.Text, vbArchive) <> "" Then Kill Text12.Text
        End If
        Exit Sub
    End If
    'Comprobamos que no existen datos entre las fechas
    Screen.MousePointer = vbHourglass
    SQ = ""
    Set Rs = New ADODB.Recordset
    Sql = "Select min(fecopera) from tmpnorma43 where codusu = " & vUsu.Codigo
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then SQ = " fecopera >='" & Format(Rs.Fields(0), FormatoFecha) & "'"
    End If
    Rs.Close
    Sql = "Select max(fecopera) from tmpnorma43 where codusu = " & vUsu.Codigo
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then SQ = SQ & " and fecopera <='" & Format(Rs.Fields(0), FormatoFecha) & "'"
    End If
    Rs.Close
    Sql = "Select count(*) from norma43 where " & SQ
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NF = 0
    If Not Rs.EOF Then
        NF = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close
    Set Rs = Nothing
    
    If NF > 0 Then
        Sql = "Se han encontrado datos entre las fechas importadas." & vbCrLf
        Sql = Sql & "( " & SQ & " )" & vbCrLf & vbCrLf
        Sql = Sql & "Puede duplicar los datos. ¿ Desea continuar ? " & vbCrLf
        If MsgBox(Sql, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        If MsgBox("¿Los datos serán importados. ¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
    End If
    
    'Haremos la insercion del registro del banco
    If BloqueoManual(True, "norma43", "clave") Then
        InsertarHcoBanco
        BloqueoManual False, "norma43", ""
        PonerModo 0
        Text1_LostFocus
    Else
        MsgBox "Tabla bloqueada por otro usuario.", vbExclamation
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Function InsertarRegistro(Ampliacion As String, ByRef ContadorMYSQL As Integer, ByRef ContadorRegistrosDeUnBanco As Integer) As Boolean
Dim vSql As String
Dim L As String

    On Error GoTo EProcesaAmpliacion
    InsertarRegistro = False
        
    vSql = "INSERT INTO tmpnorma43 (codusu,orden, codmacta, fecopera,"
    vSql = vSql & "fecvalor, importeD, importeH,  concepto,"
    vSql = vSql & "numdocum, saldo) VALUES (" & vUsu.Codigo & "," & ContadorMYSQL & ",'"
    'Numero de apunte
    txtDatos.Text = txtDatos.Text & Right("     " & NumRegElim, 5)
    'Fecha operacion
    L = RecuperaValor(CadenaDesdeOtroForm, 1)
    txtDatos.Text = txtDatos.Text & "  " & Format(L, "dd/mm/yyyy")
    vSql = vSql & Cta & "','" & L
    'Fc Valor
    L = RecuperaValor(CadenaDesdeOtroForm, 2)
    txtDatos.Text = txtDatos.Text & " " & Format(L, "dd/mm/yyyy")
    vSql = vSql & "','" & L
    'Importe DEBE/HABER
    vSql = vSql & "'," & RecuperaValor(CadenaDesdeOtroForm, 3)
    L = RecuperaValor(CadenaDesdeOtroForm, 3)
    NF = 0
    If L = "NULL" Then
        NF = 1
        L = RecuperaValor(CadenaDesdeOtroForm, 4)
    End If
    
    L = TransformaPuntosComas(L)
    L = Format(L, FormatoImporte)
    cad = "              "
    If NF = 0 Then
        'Debe
        txtDatos.Text = txtDatos.Text & "  " & Right("              " & L, 14) & "    " & cad
    Else
        txtDatos.Text = txtDatos.Text & "  " & cad & "    " & Right("              " & L, 14)
    End If
    vSql = vSql & "," & RecuperaValor(CadenaDesdeOtroForm, 4)
    
    'El concepto lo saco de la linea de aqui
    cad = DevNombreSQL(Trim(Ampliacion))  '30 como mucho
    vSql = vSql & ",'" & cad & "',"
    txtDatos.Text = txtDatos.Text & "    " & Ampliacion & vbCrLf
        
    'NumDocum
    vSql = vSql & "'" & RecuperaValor(CadenaDesdeOtroForm, 5) & "'"
    Saldo = Saldo - Importe
    cad = TransformaComasPuntos(CStr(Saldo))
    vSql = vSql & "," & cad & ")"
    'Para la BD
    ContadorMYSQL = ContadorMYSQL + 1
    
    'Para comprobar los regisitros
    ContadorRegistrosDeUnBanco = ContadorRegistrosDeUnBanco + 1
    'El que habia.
    NumRegElim = NumRegElim + 1 'Contador mas uno
    Conn.Execute vSql
    
    InsertarRegistro = True
    Exit Function
EProcesaAmpliacion:
    MuestraError Err.Number, Err.Description & vbCrLf & vSql
       
End Function

'Metere en CadenaDesdeOtroForm, empipado
' Fecha operacion, fecha valor, importeDebe, importe haber, numdocum
Private Function ProcesaLineaASiento(ByRef Lin As String, vAmpliacion As String) As Boolean
Dim Debe As Boolean


    ProcesaLineaASiento = False
    CadenaDesdeOtroForm = ""
    'Fecha operacion
    cad = Mid(Lin, 11, 6)
    cad = "20" & Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5, 2)
    If Not IsDate(cad) Then
        MsgBox "Formato fecha incorrecto", vbExclamation
        Exit Function
    End If
    CadenaDesdeOtroForm = Format(cad, FormatoFecha) & "|"
    
    'Fecha valor
    cad = Mid(Lin, 17, 6)
    cad = "20" & Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5, 2)
    If Not IsDate(cad) Then
        MsgBox "Formato fecha incorrecto", vbExclamation
        Exit Function
    End If
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Format(cad, FormatoFecha) & "|"
    
    
    'Importe
    cad = Mid(Lin, 28, 1)
    Debe = cad = "1"
    cad = Mid(Lin, 29, 14)
    If Not IsNumeric(cad) Then
        MsgBox "Importe registro 22 incorrecto: " & cad, vbExclamation
        Exit Function
    End If
    Importe = Val(cad) / 100
    cad = TransformaComasPuntos(CStr(Importe))
    
    'Importe debe / haber
    If Debe Then
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & cad & "|NULL|"
    Else
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "NULL|" & cad & "|"
    End If
    
    
    'Posible ampliacion
    If Len(Lin) > 53 Then
        vAmpliacion = Trim(Mid(Lin, 53))
        If Len(vAmpliacion) > 30 Then vAmpliacion = Mid(vAmpliacion, 1, 30)
    Else
        vAmpliacion = ""
    End If
    
  '  'Para el arrastrado
  '  'Esto va al reves de la contbiliad, ya k trabajamos con la cuenta del banoc
  '  'ANTES del 25 de Novi
    If Not Debe Then Importe = Importe * -1
  '  If Debe Then Importe = Importe * -1
    'Num docum
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Mid(Lin, 43, 10)
    ProcesaLineaASiento = True
End Function

Private Function ProcesaAmpliacion2(miLinea As String) As String
Dim Cadena As String
Dim C2 As String
Dim Blanco As Boolean
Dim I As Integer

    Cadena = ""
    Blanco = False
    For I = 5 To Len(miLinea)
        C2 = Mid(miLinea, I, 1)
        If C2 = " " Then
             If Not Blanco Then
                Cadena = Cadena & C2
                Blanco = True
            End If
        Else
            Blanco = False
            Cadena = Cadena & C2
        End If
    Next I
    If Len(Cadena) > 100 Then Cadena = Mid(Cadena, 1, 100)
    ProcesaAmpliacion2 = Cadena
End Function

Private Function HacerComprobaciones(ByRef Lin As String, ContadorRegistrosBanco As Integer, TotalRegistrosInsertados As Integer) As Boolean
Dim Ok As Boolean
Dim InsercionesActuales As Integer
    Set Rs = New ADODB.Recordset
    HacerComprobaciones = False
    InsercionesActuales = NumRegElim - 1
    cad = "Select max(orden) from tmpnorma43 where codusu =" & vUsu.Codigo
    cad = cad & " AND codmacta ='" & Cta & "'"
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NF = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then NF = Rs.Fields(0)
    End If
    Rs.Close
    
    'Numero de lineas insertadas
    Ok = False
    'Total registros en BD
    If NF = ContadorRegistrosBanco Then
        'Coinciden los contadores de insercion parcial
        
        NF = Val(Mid(Lin, 21, 5)) + Val(Mid(Lin, 40, 5))
        If NF = NumRegElim - 1 Then Ok = True
    End If
    If Not Ok Then
        'Error en contadores de registros
        MsgBox "Error en contadores de registo", vbExclamation
        NumRegElim = 0
    End If
    
    
    
    If NumRegElim > 0 Then
        'Obtengo la suma de importes
        cad = "Select sum(importeD)as debe,sum(importeH) as haber,sum(importeD)-sum(importeH) from tmpnorma43 where codusu = " & vUsu.Codigo
        cad = cad & " AND codmacta ='" & Cta & "'"
        'Enero 2009.
        'Estamos admitiendo ficheros que , aun siendo de la misma cuenta, tran mas de una entrada 11| (cabecera de cuenta
        NF = ContadorRegistrosBanco - InsercionesActuales
        cad = cad & " AND orden >" & NF
        Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not Rs.EOF Then
            cad = CStr(Val(Mid(Lin, 26, 14)) / 100)
            CadenaDesdeOtroForm = DBLet(Rs.Fields(0), "N")
            Ok = (cad = CadenaDesdeOtroForm)
            If Ok Then
                cad = CStr(Val(Mid(Lin, 45, 14)) / 100)
                CadenaDesdeOtroForm = DBLet(Rs.Fields(1), "N")
                Ok = (cad = CadenaDesdeOtroForm)
            End If
            If Ok Then
                Importe = Val(Mid(Lin, 60, 14)) / 100
                If Mid(Lin, 59, 1) = "2" Then Importe = Importe * -1
                
                If ContadorRegistrosBanco = 0 Then
                    cad = "Fichero de comprobación de saldos: " & vbCrLf & vbCrLf
                    cad = cad & "Saldo: " & CStr(Importe)
                    cad = cad & vbCrLf & vbCrLf & vbCrLf
                    cad = cad & "¿Desea eliminar el archivo?"
                    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
                        If Dir(Text12.Text, vbArchive) <> "" Then
                            Kill Text12.Text
                            Text12.Text = ""
                        End If
                    End If
                End If
                
            End If
        End If
        Rs.Close
        If Ok Then
            NumRegElim = 1
        Else
            NumRegElim = 0
        End If
    End If
    
    'Si llegamos aqui y numregelim>0 esta bien
    If NumRegElim > 0 Then HacerComprobaciones = True
    Set Rs = Nothing
    
End Function


Private Sub PonerModo(vModo As Byte)
    Select Case vModo
    Case 0
        'Primer frame
        Frame1.Enabled = True
        Frame2.visible = False
    Case 1
        Frame2.visible = True
        Frame1.Enabled = False
    End Select
    Me.Height = Me.Height + 420
    Me.Width = Me.Width + 150
    Me.Refresh
End Sub


Private Sub InsertarHcoBanco()
Dim Codigo As Long
    
    Set Rs = New ADODB.Recordset
    Codigo = 0
    Sql = "Select max(codigo) from norma43"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Codigo = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close
    Codigo = Codigo + 2   'Dejare un salto con el objetivo de poder cuadrar mas adelante
    
    Sql = "Select * from tmpnorma43 where codusu = " & vUsu.Codigo & " ORDER By Orden"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Cadena de insercion
    Sql = "INSERT INTO norma43 (codigo, codmacta, fecopera, fecvalor, importeD,"
    Sql = Sql & "importeH, concepto, numdocum, saldo, punteada) VALUES ("
    While Not Rs.EOF
        cad = Codigo & ",'" & Rs!codmacta & "','" & Format(Rs!fecopera, FormatoFecha)
        cad = cad & "','" & Format(Rs!fecvalor, FormatoFecha) & "',"
        If IsNull(Rs!ImporteD) Then
            cad = cad & "NULL," & TransformaComasPuntos(CStr(Rs!ImporteH))
        Else
            cad = cad & TransformaComasPuntos(CStr(Rs!ImporteD)) & ",NULL"
        End If
        cad = cad & ",'" & DevNombreSQL(DBLet(Rs!Concepto)) & "','" & Rs!Numdocum & "',"
        cad = cad & TransformaComasPuntos(CStr(Rs!Saldo)) & ",0);"
        cad = Sql & cad
        'Ejecutamos SQL
        Conn.Execute cad
        Codigo = Codigo + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    'Ahora deberiamos eliminar el archivo
    If chkElimmFich.Value = 1 Then
        If Dir(Text12.Text, vbArchive) <> "" Then Kill Text12.Text
         MsgBox "Importación finalizada", vbInformation
    Else
        MsgBox "Proceso finalizado. El fichero NO será eliminado", vbExclamation
    End If
End Sub






Private Sub VerAsiento()

    If Me.ListView2.ListItems.Count = 0 Then Exit Sub
    If ListView2.SelectedItem Is Nothing Then Exit Sub
    
    'Rs!NumAsien & "|" & Rs!NumDiari & "|" & Rs!Linliapu & "|"
    
    frmAsientosHco.Asiento = DBLet(RecuperaValor(ListView2.SelectedItem.Tag, 2), "N") & "|" & ListView2.SelectedItem.Text & "|" & ListView2.SelectedItem.SubItems(1) & "|"
    frmAsientosHco.SoloImprimir = True
    frmAsientosHco.Show vbModal
            
    
    

End Sub



Private Sub VerVencimiento(Cobro As Boolean, VerCuenta As Boolean)

    CadenaDesdeOtroForm = ""
    
    
    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
    If ListView1.SelectedItem.Checked Then
        MsgBox "Extracto ya esta punteado", vbExclamation
        Exit Sub
    End If
    
     
    If Text1.Text = "" Then
        MsgBoxA "Cuenta banco VACIA", vbExclamation
        PonFoco Text4
        Exit Sub
    End If
    
    If FechaCorrecta2(CDate(ListView1.SelectedItem.Text), True) > 1 Then Exit Sub
    
    If Not VerCuenta Then
    
        If Cobro Then
            Sql = "situacion=0 and (coalesce(Cobros.ImpVenci, 0) + coalesce(Cobros.Gastos, 0) - coalesce(Cobros.impcobro, 0) = " & DBSet(ListView1.SelectedItem.SubItems(1), "N") & " )"
            Cta = DevuelveDesdeBD("count(*)", "cobros", Sql & " AND 1", "1")
        Else
            Sql = "situacion=0 and (coalesce(pagos.impefect, 0)  - coalesce(pagos.imppagad, 0) = " & DBSet(ListView1.SelectedItem.SubItems(1), "N") & " )"
            Cta = DevuelveDesdeBD("count(*)", "pagos", Sql & " AND 1", "1")
        End If
        If Val(Cta) = 0 Then
            MsgBox "Ningun " & IIf(Cobro, "cobro", "pago") & " pendiente con este importe : " & ListView1.SelectedItem.SubItems(1), vbExclamation
            Cta = ""
        End If
        If Cta = "" Then Exit Sub
        Cta = ""
    End If
    
    FrameDatosCobroPago.top = 0
    txtVto(0).Text = ListView1.SelectedItem.Text
    txtVto(1).Text = ListView1.SelectedItem.SubItems(1)
    txtVto(2).Text = ListView1.SelectedItem.SubItems(4)
    FrameDatosCobroPago.visible = True
    
    
    If VerCuenta Then
        
        frmPunteoCobrosCta.ImporteBanco = ImporteFormateado(ListView1.SelectedItem.SubItems(1))
        frmPunteoCobrosCta.Show vbModal
    
    Else
        
        Set frmTESVerCobPag = New frmTESVerCobrosPagos
        frmTESVerCobPag.Situacion = 1
        frmTESVerCobPag.vSql = Sql
        frmTESVerCobPag.OrdenarEfecto = False
        frmTESVerCobPag.Regresar = True
        frmTESVerCobPag.Cobros = Cobro
        frmTESVerCobPag.Show vbModal
        
        Set frmTESVerCobPag = Nothing
    
    End If
    FrameDatosCobroPago.visible = False
    
    
    If CadenaDesdeOtroForm = "" Then Exit Sub
      
    Screen.MousePointer = vbHourglass
    
    Set miRsAux = New ADODB.Recordset
    If VerCuenta Then
        Sql = "  WHERE  " & CadenaDesdeOtroForm
    Else
        Sql = "  WHERE  numserie = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 1), "T")
        Sql = Sql & " AND numfactu = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "T")
        Sql = Sql & " AND fecfactu = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "F")
        Sql = Sql & " AND numorden = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 4), "N")
    End If
    
    
    If Cobro Then
        Sql = " left join formapago on cobros.codforpa=formapago.codforpa" & Sql
        Sql = " FROM cobros  left join cuentas on cobros.codmacta=cuentas.codmacta " & Sql
        Sql = "Select numserie,numfactu,fecfactu,numorden,cobros.codmacta,cobros.codforpa,fecvenci,impvenci,gastos,if(coalesce(nomclien,'')<>'',nomclien,nommacta ) nommacta,tipforpa,ccostedef codccost" & Sql
    Else
        Sql = " left join formapago on pagos.codforpa=formapago.codforpa" & Sql
        Sql = " FROM pagos  left join cuentas on pagos.codmacta=cuentas.codmacta " & Sql
        Sql = "Select numserie,numfactu,fecfactu,numorden,pagos.codmacta,pagos.codforpa,fecefect as fecvenci ,impefect impvenci,0 gastos,if(coalesce(nomprove,'')<>'',nomprove,nommacta ) nommacta,tipforpa,ccostedef codccost" & Sql
    End If
    'Abro en modo KEYSET
    miRsAux.Open Sql, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBoxA "No se ha encontrado el vencimiento: " & CadenaDesdeOtroForm, vbExclamation
    
    Else
        Sql = IIf(Cobro, "COBRO", "PAGO") & vbCrLf & String(45, "-") & vbCrLf
        Sql = Sql & IIf(Cobro, "Cliente", "Proveedor") & ": " & miRsAux!Nommacta & vbCrLf
        Sql = Sql & "Vencimiento: " & miRsAux!NUmSerie & " " & Format(miRsAux!numfactu, "000000") & " de fecha " & miRsAux!Fecfactu & vbCrLf & vbCrLf
        Sql = Sql & "Importe pendiente vencimiento: " & Format(miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N"), FormatoImporte) & vbCrLf
        Sql = Sql & "Importe BANCO: " & ListView1.SelectedItem.SubItems(1) & vbCrLf & vbCrLf & "¿Continuar?"
        If VerCuenta Then
            Sql = ""
        Else
            If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then Sql = ""
        End If
        If Sql = "" Then
            Conn.BeginTrans
            If Contabilizar(CStr(CadenaDesdeOtroForm), Cobro) Then
                Conn.CommitTrans
                HaGeneradoAsiento
            Else
                Conn.RollbackTrans
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
      
      
      
      


End Sub






Private Function Contabilizar(ByVal ElVto As String, Cobro As Boolean) As Boolean
Dim Mc As Contadores
Dim FP As Ctipoformapago
Dim Sql As String
Dim Ampliacion As String
Dim Numdocum As String
Dim Conce As Integer
Dim LlevaContr As Boolean
Dim Im As Currency
Dim Debe As Boolean
Dim ElConcepto As Integer
Dim vNumDiari As Integer
Dim Situacion As Integer


Dim CtaBancoGastos As String
Dim DescuentaImporteDevolucion As Boolean
Dim Sql5 As String
Dim Fecha As Date
Dim Linea As Integer

Dim ImporteCtaCliente As Currency
Dim MaDeUnCobro As Boolean

    On Error GoTo ECon
    Contabilizar = False
    
    Fecha = CDate(ListView1.SelectedItem.Text)
     
    
    Set Mc = New Contadores
    If Mc.ConseguirContador("0", Fecha <= vParam.fechafin, True) = 1 Then Err.Raise 513, , "Fechas contables"

    
    
    
    
    Sql5 = DBLet(miRsAux!TipForpa, "N")
    Set FP = New Ctipoformapago
    If FP.Leer(CInt(Sql5)) Then Err.Raise 513, , "Forma pago"
    
    

   
    'Inserto cabecera de apunte
    Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
    If Cobro Then
        Sql = Sql & FP.diaricli
        vNumDiari = FP.diaricli
    Else
        Sql = Sql & FP.diaripro
        vNumDiari = FP.diaripro
    End If
    Sql = Sql & ",'" & Format(Fecha, FormatoFecha) & "'," & Mc.Contador & ","
    
    Numdocum = "Generado desde Tesorería el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & DevNombreSQL(vUsu.Nombre)
    If Im < 0 Then Numdocum = Numdocum & "  (ABONO)"
   
    
    MaDeUnCobro = False
    Linea = 0
    Do
        Linea = Linea + 1
        miRsAux.MoveNext
    Loop Until miRsAux.EOF
    If Linea > 1 Then
        MaDeUnCobro = True
        Numdocum = Numdocum & vbCrLf & "Cobros: " & Linea & "    Total banco: " & Me.ListView1.SelectedItem.SubItems(1)
    End If
    miRsAux.MoveFirst
    
    Numdocum = Numdocum & vbCrLf & ListView1.SelectedItem.SubItems(4)
    Sql = Sql & DBSet(Numdocum, "T") & ","
    If Cobro Then
        Sql = Sql & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizar Cobros punteo banco')"
    Else
        Sql = Sql & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizar Pagos punteo banco')"
    End If
    
    Conn.Execute Sql
        
        
        
        
        
        
        
        
        
    '**********************  Lineas de apunte
    Linea = 0
        
    Do
      
        Linea = Linea + 1
            
        'Inserto en las lineas de apuntes
        Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
        Sql = Sql & "codmacta, numdocum, codconce, ampconce,timporteD,"
        Sql = Sql & " timporteH, codccost, ctacontr, idcontab, punteada,"
        
        'campos añadidos en hlinapu
        If Cobro Then
            Sql = Sql & "numserie,numfaccl,fecfactu,numorden,tipforpa,reftalonpag,bancotalonpag) VALUES ("
        Else
            Sql = Sql & "numserie,numfacpr,fecfactu,numorden,tipforpa,reftalonpag,bancotalonpag) VALUES ("
        End If
        
        If Cobro Then
            Sql = Sql & FP.diaricli
        Else
            Sql = Sql & FP.diaripro
        End If
        Sql = Sql & ",'" & Format(Fecha, FormatoFecha) & "'," & Mc.Contador & ","
        
        
        'numdocum
        Numdocum = DevNombreSQL(miRsAux!numfactu)
        If Cobro Then
            Numdocum = miRsAux!NUmSerie & Format(miRsAux!numfactu, "0000000")
        Else
            
            If vParam.CodiNume = 1 Then
                'Quiero el numero de registro. Intento buscar
                Sql5 = "numserie = " & DBSet(RecuperaValor(miRsAux!NUmSerie, 1), "T")
                Sql5 = Sql5 & " AND numfactu = " & DBSet(miRsAux!numfactu, "T")
                Sql5 = Sql5 & " AND fecfactu = " & DBSet(miRsAux!Fecfactu, "F") & " AND 1"
                Sql5 = DevuelveDesdeBD("numregis", "factpro", Sql5, "1")
                If Sql5 <> "" Then
                    Sql5 = Right("0000000000" & Sql5, 10)
                    Numdocum = Sql5
                End If
           
            End If
            
            
            
            Sql5 = ""
        End If
        
        Im = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N")
        'Concepto y ampliacion del apunte
        Ampliacion = ""
        ImporteCtaCliente = Im - DBLet(miRsAux!Gastos, "N")
        If Cobro Then
            
            'CLIENTES
            Debe = False
            If ImporteCtaCliente < 0 Then
                If Not vParam.abononeg Then Debe = True
            End If
            If Debe Then
                Conce = FP.ampdecli
                LlevaContr = FP.ctrdecli = 1
                ElConcepto = FP.condecli
            Else
                ElConcepto = FP.conhacli
                Conce = FP.amphacli
                LlevaContr = FP.ctrhacli = 1
            End If
        Else
            'PAGOS
            Debe = True
            If Im < 0 Then
                If Not vParam.abononeg Then Debe = False
            End If
            If Debe Then
                Conce = FP.ampdepro
                LlevaContr = FP.ctrdepro = 1
                ElConcepto = FP.condepro
            Else
                ElConcepto = FP.conhapro
                Conce = FP.amphapro
                LlevaContr = FP.ctrhapro = 1
            End If
    
        End If
               
        'Si el importe es negativo y no permite abonos negativos
        'como ya lo ha cambiado de lado (dbe <-> haber)
        If ImporteCtaCliente < 0 Then
            If Not vParam.abononeg Then ImporteCtaCliente = Abs(ImporteCtaCliente)
        End If
           
               
        If Conce = 2 Then
           Sql5 = IIf(Cobro, miRsAux!FecVenci, miRsAux!fecefec)
           Ampliacion = Ampliacion & Sql5 'Fecha vto
           Sql5 = ""
        ElseIf Conce = 4 Then
            'Contra partida
            Ampliacion = DevNombreSQL(DBLet(miRsAux!Nommacta, "T")) 'falta
        ElseIf Conce = 6 Then
            'Cuenta
            Sql5 = miRsAux!NUmSerie
            
            If Cobro Then
                MiVariableAuxiliar = Sql5 & Format(miRsAux!numfactu, "0000000")
                Sql5 = ""
            Else
                If Sql5 = "1" Then Sql5 = ""
                MiVariableAuxiliar = Sql5 & miRsAux!numfacut
                Sql5 = ""
            End If
            If Len(MiVariableAuxiliar) > 23 Then MiVariableAuxiliar = Mid(MiVariableAuxiliar, 1, 20)
            Ampliacion = Mid(miRsAux!Nommacta, 1, 39 - Len(MiVariableAuxiliar))
            Ampliacion = Sql5 & Ampliacion & " " & MiVariableAuxiliar
            Sql5 = ""
        Else
            
           If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
           If Cobro Then
                Ampliacion = Ampliacion & miRsAux!NUmSerie & Format(miRsAux!numfactu, "0000000")
           Else
                Ampliacion = Ampliacion & Mid(miRsAux!numfactu, 1, 15)
           End If
        End If
        
        'Fijo en concepto el codconce
        If Conce <> 6 Then
            Conce = ElConcepto
            cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
        Else
            cad = ""
            Conce = ElConcepto
        End If
        Ampliacion = Trim(cad & " " & Ampliacion)
        Ampliacion = Mid(Ampliacion, 1, 45)
        
        
        
        
        'Ahora ponemos linliapu codmacta numdocum codconce ampconce timported timporte codccost ctacontr idcontab punteada
        'Cuenta Cliente/proveedor
        
        cad = Linea & ",'" & Text1.Text & "','" & Numdocum & "'," & Conce & ",'" & DevNombreSQL(Ampliacion) & "',"
        'Si el cliente-prov va al debe el banoc va al haber
        If Debe Then
            cad = cad & "NULL," & TransformaComasPuntos(CStr(ImporteCtaCliente))
        Else
            cad = cad & TransformaComasPuntos(CStr(ImporteCtaCliente)) & ",NULL"
        End If
        'Codccost
        cad = cad & ",NULL,"
        'cntrapar
        cad = cad & "'" & miRsAux!codmacta & "'"
        
        If Cobro Then
            cad = cad & ",'COBROS',1,"
            cad = cad & DBSet(miRsAux!NUmSerie, "T") & "," '& RecuperaValor(Vto, 2) & ","
        Else
            cad = cad & ",'PAGOS',1,"
            cad = cad & DBSet(RecuperaValor(miRsAux!NUmSerie, 1), "T") & ","
        End If
        
        cad = cad & DBSet(miRsAux!numfactu, "T") & "," & DBSet(miRsAux!Fecfactu, "F") & ","
        cad = cad & DBSet(miRsAux!numorden, "N") & "," & FP.tipoformapago & "," & ValorNulo & "," & ValorNulo & ")"
        
        cad = Sql & cad
        Conn.Execute cad
        
        If DBLet(miRsAux!Gastos, "N") <> 0 Then
            CtaBancoGastos = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", Text1.Text, "T")
            If CtaBancoGastos = "" Then Err.Raise 513, , "Cuenta gastos sin configurar"
            
            'Cuenta Cliente/proveedor
            Linea = Linea + 1
            cad = Linea & ",'" & CtaBancoGastos & "','" & Numdocum & "'," & Conce & ",'" & DevNombreSQL(Ampliacion) & "',"
            'Importe cobro-pago
            ' nos lo dire "debe"
            If Not Debe Then
                cad = cad & "NULL," & TransformaComasPuntos(CStr(miRsAux!Gastos))
            Else
                cad = cad & TransformaComasPuntos(CStr(miRsAux!Gastos)) & ",NULL"
            End If
            'Codccost
            Sql5 = ""
            If vParam.autocoste Then Sql5 = DBLet(miRsAux!CodCCost, "T")
            If Sql5 <> "" Then
                Sql5 = DBSet(Sql5, "T")
            Else
                Sql5 = "NULL"
            End If
            cad = cad & "," & Sql5 & ","
            cad = cad & "'" & miRsAux!codmacta & "'"
            
            If Cobro Then
                cad = cad & ",'COBROS',0,"
                cad = cad & DBSet(miRsAux!NUmSerie, "T") & ","
            Else
                cad = cad & ",'PAGOS',0,"
                cad = cad & DBSet(miRsAux!NUmSerie, "T") & ","
            End If
            
            cad = cad & DBSet(miRsAux!numfactu, "T") & "," & DBSet(miRsAux!Fecfactu, "F") & ","
            cad = cad & DBSet(miRsAux!numorden, "N") & "," & FP.tipoformapago & "," & ValorNulo & "," & ValorNulo & ")"
            
            cad = Sql & cad
            Conn.Execute cad
        
        End If
           
        'El banco    *******************************************************************************
        '---------------------------------------------------------------------------------------------
        
        'Vuelvo a fijar los valores
         'Concepto y ampliacion del apunte
        Ampliacion = ""
        If Cobro Then
           'CLIENTES
            
            If Debe Then
                Conce = FP.ampdecli
                LlevaContr = FP.ctrdecli = 1
                ElConcepto = FP.condecli
            Else
                ElConcepto = FP.conhacli
                Conce = FP.amphacli
                LlevaContr = FP.ctrhacli = 1
            End If
        Else
            'PAGOS
            
            If Debe Then
                Conce = FP.ampdepro
                LlevaContr = FP.ctrdepro = 1
                ElConcepto = FP.condepro
            Else
                ElConcepto = FP.conhapro
                Conce = FP.amphapro
                LlevaContr = FP.ctrhapro = 1
            End If
        End If
               
               
        If Conce = 2 Then
           Sql5 = IIf(Cobro, miRsAux!FecVenci, miRsAux!fecefec)
           Ampliacion = Ampliacion & Sql5 'Fecha vto
           Sql5 = ""
           
        ElseIf Conce = 4 Then
            'Contra partida
            Ampliacion = DevNombreSQL(DBLet(miRsAux!Nommacta, "T"))
        ElseIf Conce = 6 Then
        
            Sql5 = RecuperaValor(miRsAux!NUmSerie, 1)
            
            If Cobro Then
                MiVariableAuxiliar = Sql5 & Format(miRsAux!numfactu, "0000000")
                Sql5 = ""
            Else
                If Sql5 = "1" Then Sql5 = ""
                MiVariableAuxiliar = Sql5 & miRsAux!numfactu
                Sql5 = "PAG "
            End If
            ' Como quiere ver toda la ampliacion en la ventana, vamos a suponore un maximo de 23 carcateres
            If Len(MiVariableAuxiliar) > 23 Then MiVariableAuxiliar = Mid(MiVariableAuxiliar, 1, 20)
            Ampliacion = Mid(RecuperaValor(Cta, 2), 1, 23 - Len(MiVariableAuxiliar))
            Ampliacion = Sql5 & Ampliacion & " " & MiVariableAuxiliar
            Sql5 = ""
        Else
        
            If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
            If Cobro Then
                 Ampliacion = Ampliacion & miRsAux!NUmSerie & Format(miRsAux!numfactu, "0000000")
            Else
                 Ampliacion = Ampliacion & Mid(miRsAux!numfactu, 1, 15)
            End If
        End If
        
        If Conce <> 6 Then
            Conce = ElConcepto
            cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
        Else
            cad = ""
            Conce = ElConcepto
        End If
        Ampliacion = Trim(cad & " " & Ampliacion)
        Ampliacion = Mid(Ampliacion, 1, 45)
        
        
        
        Linea = Linea + 1
        cad = Linea & ",'" & miRsAux!codmacta & "','" & Numdocum & "'," & Conce & ",'" & Ampliacion & "',"
        'Importe cliente
        'Si el cobro/pago va al debe el contrapunte ira al haber
        If Not vParamT.abononeg Then ImporteCtaCliente = Abs(ImporteCtaCliente)
        If Debe Then
            'al debe
            cad = cad & TransformaComasPuntos(CStr(ImporteCtaCliente)) & ",NULL"
        Else
            'al haber
            cad = cad & "NULL," & TransformaComasPuntos(CStr(ImporteCtaCliente))
        End If
        
        'Codccost
        cad = cad & ",NULL,"
        
        
        cad = cad & "'" & Text1.Text & "'"
        
        If Cobro Then
            cad = cad & ",'COBROS',0," ' idcontab
        Else
            cad = cad & ",'PAGOS',0," ' idcontab
        End If
        
        
        cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
        
        cad = Sql & cad
        Conn.Execute cad
        
            
        
        
        If Cobro Then
            Sql = FP.diaricli
        Else
            Sql = FP.diaripro
        End If
        
        
      
        If Cobro Then
            Sql = "cobros"
            Ampliacion = "fecultco"
            Numdocum = "impcobro"
           
        Else
            
            Sql = "pagos"
            Ampliacion = "fecultpa"
            Numdocum = "imppagad"
           
        End If
        
        
        
        
        If Cobro Then
            Sql = "update cobros set impcobro = " & DBSet(Im, "N")
            Sql = Sql & ", situacion=1, fecultco = " & DBSet(Fecha, "F")
            Sql = Sql & " where numserie = " & DBSet(miRsAux!NUmSerie, "T") & " and numfactu = " & DBSet(miRsAux!numfactu, "N")
            Sql = Sql & " and fecfactu = " & DBSet(miRsAux!Fecfactu, "F") & " and numorden = " & miRsAux!numorden
        
        Else
            
            Sql = "update pagos set imppagad = " & DBSet(Im, "N")
            Sql = Sql & ", situacion=1, fecultpa = " & DBSet(Fecha, "F")
            Sql = Sql & " where numserie = " & DBSet(miRsAux!NUmSerie, "T") & " and numfactu = " & DBSet(miRsAux!numfactu, "T")
            Sql = Sql & " and fecfactu = " & DBSet(miRsAux!Fecfactu, "F") & " and numorden = " & DBSet(miRsAux!numorde, "N")
            Sql = Sql & " and codmacta = " & DBSet(miRsAux!codmacta, "T")
            
        End If
        Conn.Execute Sql
        
        miRsAux.MoveNext
        If Not miRsAux.EOF Then Debug.Assert False
    Loop Until miRsAux.EOF
    miRsAux.Close
    
    
    Contabilizar = True

    
   
ECon:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Contabilizar anticipo", Err.Description
        
    End If
    Set Mc = Nothing
    Set FP = Nothing
    Set miRsAux = Nothing
End Function


