VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAVNAvnics 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A.V.N.I.C.S."
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   15270
   Icon            =   "frmAVNAvnics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3825
      TabIndex        =   68
      Top             =   0
      Width           =   3330
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   240
         TabIndex        =   69
         Top             =   180
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Renovación Avnics"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cálculo Intereses"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Contabilización intereses"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cancelación Avnics"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ayuda Modelo 123"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Grabacion modelo 193"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   12420
      TabIndex        =   66
      Top             =   180
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
      TabIndex        =   64
      Top             =   0
      Width           =   3510
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   65
         Top             =   180
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
               Object.Tag             =   "0"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   7260
      TabIndex        =   62
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   63
         Top             =   210
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Index           =   0
      Left            =   240
      TabIndex        =   34
      Top             =   765
      Width           =   14850
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frmAVNAvnics.frx":000C
         Left            =   9600
         List            =   "frmAVNAvnics.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Clave Alta|N|N|0|2|avnic|codialta|||"
         Top             =   240
         Width           =   2610
      End
      Begin VB.TextBox text1 
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
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "F.Alta|F|N|||avnic|fechalta|dd/mm/yyyy||"
         Top             =   225
         Width           =   1275
      End
      Begin VB.TextBox text1 
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
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Código Avnics|N|N|0|999999|avnic|codavnic|000000|S|"
         Text            =   "000000"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox text1 
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
         Left            =   4065
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Ejercicio|N|N|0|9999|avnic|anoejerc|0000|S|"
         Text            =   "2020"
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Clave Alta"
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
         Left            =   8085
         TabIndex        =   51
         Top             =   270
         Width           =   1560
      End
      Begin VB.Label Label20 
         Caption         =   "Fecha Alta"
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
         Left            =   5280
         TabIndex        =   49
         Top             =   255
         Width           =   615
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   6000
         Picture         =   "frmAVNAvnics.frx":002C
         ToolTipText     =   "Buscar fecha"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Ejercicio"
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
         Left            =   3000
         TabIndex        =   36
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Código Avnic"
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
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   31
      Top             =   8145
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   32
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   13950
      TabIndex        =   30
      Top             =   8235
      Width           =   1065
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   12630
      TabIndex        =   29
      Top             =   8220
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6405
      Left            =   240
      TabIndex        =   33
      Top             =   1635
      Width           =   14850
      _ExtentX        =   26194
      _ExtentY        =   11298
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmAVNAvnics.frx":00B7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(26)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label29"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "imgZoom(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgBuscar(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label6(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "text1(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "text1(5)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "text1(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "text1(7)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "text1(6)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "FrameDatosAlta"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "FrameDatosContacto"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "text1(26)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "text1(9)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "text2(9)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "text1(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Movimientos"
      TabPicture(1)   =   "frmAVNAvnics.frx":00D3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux0"
      Tab(1).ControlCount=   1
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   5850
         Left            =   -74865
         TabIndex        =   70
         Top             =   405
         Width           =   14400
         Begin VB.TextBox text2 
            Alignment       =   1  'Right Justify
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
            Index           =   2
            Left            =   11880
            TabIndex        =   88
            Top             =   225
            Width           =   1905
         End
         Begin VB.TextBox text2 
            Alignment       =   1  'Right Justify
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
            Index           =   1
            Left            =   9945
            TabIndex        =   87
            Top             =   225
            Width           =   1905
         End
         Begin VB.TextBox text2 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   8010
            TabIndex        =   85
            Top             =   225
            Width           =   1905
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
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
            Left            =   2880
            TabIndex        =   77
            ToolTipText     =   "Buscar fecha"
            Top             =   2160
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   855
            MaxLength       =   10
            TabIndex        =   76
            Tag             =   "Ejercicio|N|N|0|9999|movim|anoejerc|0000|S|"
            Text            =   "Año"
            Top             =   2145
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   105
            TabIndex        =   75
            Tag             =   "Código Avnics|N|N|0|999999|movim|codavnic|000000|S|"
            Text            =   "Codigo"
            Top             =   2145
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Frame FrameToolAux 
            Height          =   555
            Left            =   60
            TabIndex        =   73
            Top             =   120
            Width           =   1605
            Begin MSComctlLib.Toolbar ToolbarAux 
               Height          =   330
               Left            =   180
               TabIndex        =   74
               Top             =   150
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   3
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Insertar"
                  EndProperty
                  BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Modificar"
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Eliminar"
                  EndProperty
               EndProperty
            End
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   6480
            MaxLength       =   15
            TabIndex        =   82
            Tag             =   "Importe Bruto|N|N|||movim|timport1|###,###,##0.00||"
            Text            =   "Bruto"
            Top             =   2160
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   4725
            MaxLength       =   15
            TabIndex        =   81
            Tag             =   "Importe|N|N|||movim|timporte|###,###,##0.00||"
            Text            =   "Importe"
            Top             =   2160
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
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
            Index           =   2
            Left            =   15420
            TabIndex        =   72
            ToolTipText     =   "Buscar concepto"
            Top             =   2130
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   3240
            MaxLength       =   35
            TabIndex        =   80
            Tag             =   "Concepto|T|N|||movim|concepto|||"
            Text            =   "concepto"
            Top             =   2160
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   8730
            MaxLength       =   15
            TabIndex        =   83
            Tag             =   "Importe Ret|N|N|||movim|timport2|###,###,##0.00||"
            Text            =   "ImpRetencion"
            Top             =   2160
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CheckBox chkAux 
            BackColor       =   &H80000005&
            Height          =   255
            Index           =   0
            Left            =   10035
            TabIndex        =   84
            Tag             =   "IntConta|N|N|0|1|movim|intconta|||"
            Top             =   2190
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Index           =   12
            Left            =   15630
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   71
            Text            =   "Nombre cuenta"
            Top             =   2160
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   2190
            TabIndex        =   78
            Tag             =   "F.Movim|F|N|||movim|fechamov|dd/mm/yyyy|S|"
            Text            =   "fechamov"
            Top             =   2160
            Visible         =   0   'False
            Width           =   645
         End
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   375
            Index           =   0
            Left            =   3720
            Top             =   480
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
            Caption         =   "AdoAux(0)"
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
         Begin MSDataGridLib.DataGrid DataGridAux 
            Height          =   4905
            Index           =   0
            Left            =   45
            TabIndex        =   79
            Top             =   765
            Width           =   14385
            _ExtentX        =   25374
            _ExtentY        =   8652
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               AllowFocus      =   0   'False
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label14 
            Caption         =   "TOTALES"
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
            Left            =   7020
            TabIndex        =   86
            Top             =   270
            Width           =   1050
         End
      End
      Begin VB.TextBox text1 
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
         Index           =   4
         Left            =   1440
         MaxLength       =   60
         TabIndex        =   5
         Tag             =   "Nombre|T|N|||avnic|nombrper|||"
         Top             =   1005
         Width           =   6180
      End
      Begin VB.TextBox text2 
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
         Index           =   9
         Left            =   2790
         TabIndex        =   46
         Top             =   2805
         Width           =   4830
      End
      Begin VB.TextBox text1 
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
         Index           =   9
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Cta.Contable|T|S|||avnic|codmacta|||"
         Top             =   2805
         Width           =   1305
      End
      Begin VB.TextBox text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Index           =   26
         Left            =   7995
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Tag             =   "Observaciones|T|S|||avnic|observac|||"
         Top             =   4065
         Width           =   6690
      End
      Begin VB.Frame FrameDatosContacto 
         Caption         =   "Datos Segundo Titular"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2850
         Left            =   135
         TabIndex        =   44
         Top             =   3420
         Width           =   7590
         Begin VB.TextBox text1 
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
            Index           =   11
            Left            =   5625
            MaxLength       =   15
            TabIndex        =   12
            Tag             =   "NIF / CIF|T|S|||avnic|nifrepre|||"
            Text            =   "12345678912345"
            Top             =   360
            Width           =   1785
         End
         Begin VB.TextBox text1 
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
            Index           =   10
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   11
            Tag             =   "NIF / CIF|T|S|||avnic|nifpers1|||"
            Top             =   360
            Width           =   1785
         End
         Begin VB.TextBox text1 
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
            Index           =   13
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   14
            Tag             =   "Domicilio|T|S|||avnic|nomcall1|||"
            Top             =   1320
            Width           =   6090
         End
         Begin VB.TextBox text1 
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
            Index           =   16
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   17
            Tag             =   "Provincia|T|S|||avnic|provinc1|||"
            Top             =   2205
            Width           =   6090
         End
         Begin VB.TextBox text1 
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
            Index           =   15
            Left            =   2100
            MaxLength       =   50
            TabIndex        =   16
            Tag             =   "Población|T|S|||avnic|poblaci1|||"
            Top             =   1770
            Width           =   5310
         End
         Begin VB.TextBox text1 
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
            Index           =   14
            Left            =   1320
            MaxLength       =   6
            TabIndex        =   15
            Tag             =   "C.Postal|T|S|||avnic|codpost1|||"
            Top             =   1770
            Width           =   735
         End
         Begin VB.TextBox text1 
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
            Index           =   12
            Left            =   1320
            MaxLength       =   60
            TabIndex        =   13
            Tag             =   "Nombre|T|S|||avnic|nombper1|||"
            Top             =   840
            Width           =   6090
         End
         Begin VB.Label Label12 
            Caption         =   "NIF Representante"
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
            Left            =   3570
            TabIndex        =   61
            Top             =   360
            Width           =   2040
         End
         Begin VB.Label Label10 
            Caption         =   "NIF"
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
            TabIndex        =   60
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Domicilio"
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
            Index           =   3
            Left            =   120
            TabIndex        =   59
            Top             =   1335
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
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
            Left            =   120
            TabIndex        =   58
            Top             =   1785
            Width           =   960
         End
         Begin VB.Label Label9 
            Caption         =   "Provincia"
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
            TabIndex        =   57
            Top             =   2235
            Width           =   960
         End
         Begin VB.Label Label6 
            Caption         =   "Nombre"
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
            Left            =   120
            TabIndex        =   56
            Top             =   855
            Width           =   735
         End
      End
      Begin VB.Frame FrameDatosAlta 
         Caption         =   "Datos Financieros"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3195
         Left            =   7920
         TabIndex        =   41
         Top             =   420
         Width           =   6780
         Begin VB.TextBox text1 
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
            Index           =   27
            Left            =   1695
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "IBAN|T|S|||avnic|iban|||"
            Top             =   855
            Width           =   600
         End
         Begin VB.TextBox text1 
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
            Index           =   25
            Left            =   1695
            MaxLength       =   6
            TabIndex        =   27
            Tag             =   "% Int.|N|N|0|999.99|avnic|porcinte|##0.00||"
            Top             =   2670
            Width           =   600
         End
         Begin VB.TextBox text1 
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
            Index           =   24
            Left            =   1695
            MaxLength       =   12
            TabIndex        =   26
            Tag             =   "Importe Ret.|N|S|0|9999999.99|avnic|imporret|#,###,##0.00||"
            Top             =   2220
            Width           =   1320
         End
         Begin VB.TextBox text1 
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
            Index           =   23
            Left            =   1695
            MaxLength       =   12
            TabIndex        =   25
            Tag             =   "Importe Per.|N|S|0|9999999.99|avnic|imporper|#,###,##0.00||"
            Top             =   1770
            Width           =   1320
         End
         Begin VB.TextBox text1 
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
            Index           =   22
            Left            =   1695
            MaxLength       =   12
            TabIndex        =   24
            Tag             =   "Importe|N|N|0|9999999.99|avnic|importes|#,###,##0.00||"
            Top             =   1320
            Width           =   1320
         End
         Begin VB.TextBox text1 
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
            Index           =   21
            Left            =   4335
            MaxLength       =   10
            TabIndex        =   23
            Tag             =   "Cuenta|T|N|||avnic|cuentaba|||"
            Top             =   855
            Width           =   1320
         End
         Begin VB.TextBox text1 
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
            Index           =   20
            Left            =   3795
            MaxLength       =   2
            TabIndex        =   22
            Tag             =   "D.C.|T|N|||avnic|digcontr|||"
            Top             =   855
            Width           =   480
         End
         Begin VB.TextBox text1 
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
            Index           =   19
            Left            =   3075
            MaxLength       =   4
            TabIndex        =   21
            Tag             =   "Sucursal|N|N|0|9999|avnic|codsucur|0000||"
            Top             =   855
            Width           =   600
         End
         Begin VB.TextBox text1 
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
            Left            =   2355
            MaxLength       =   4
            TabIndex        =   20
            Tag             =   "Banco|N|N|0|9999|avnic|codbanco|0000||"
            Top             =   855
            Width           =   600
         End
         Begin VB.TextBox text1 
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
            Index           =   17
            Left            =   1695
            MaxLength       =   10
            TabIndex        =   18
            Tag             =   "F.Vto.|F|N|||avnic|fechavto|dd/mm/yyyy||"
            Top             =   420
            Width           =   1275
         End
         Begin VB.Label Label15 
            Caption         =   "% Interes"
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
            TabIndex        =   55
            Top             =   2700
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Imp.Retención"
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
            TabIndex        =   54
            Top             =   2250
            Width           =   1470
         End
         Begin VB.Label Label11 
            Caption         =   "Imp.Percepción"
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
            TabIndex        =   53
            Top             =   1800
            Width           =   1515
         End
         Begin VB.Label Label3 
            Caption         =   "Importe Avnic"
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
            TabIndex        =   52
            Top             =   1350
            Width           =   1515
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1380
            Picture         =   "frmAVNAvnics.frx":00EF
            ToolTipText     =   "Buscar fecha"
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label17 
            Caption         =   "IBAN Avnic"
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
            TabIndex        =   48
            Top             =   885
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Fecha Vto."
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
            TabIndex        =   43
            Top             =   435
            Width           =   1110
         End
      End
      Begin VB.TextBox text1 
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
         Index           =   6
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   7
         Tag             =   "C.Postal|T|N|||avnic|codposta|||"
         Top             =   1935
         Width           =   735
      End
      Begin VB.TextBox text1 
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
         Left            =   2220
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "Población|T|N|||avnic|poblacio|||"
         Top             =   1935
         Width           =   5400
      End
      Begin VB.TextBox text1 
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
         Index           =   8
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   9
         Tag             =   "Provincia|T|N|||avnic|provinci|||"
         Top             =   2370
         Width           =   6180
      End
      Begin VB.TextBox text1 
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
         Index           =   5
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "Domicilio|T|N|||avnic|nomcalle|||"
         Top             =   1485
         Width           =   6180
      End
      Begin VB.TextBox text1 
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
         Index           =   3
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   4
         Tag             =   "NIF / CIF|T|N|||avnic|nifperso|||"
         Top             =   520
         Width           =   1785
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre"
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
         Left            =   240
         TabIndex        =   50
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label8 
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
         Left            =   240
         TabIndex        =   47
         Top             =   2805
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1170
         ToolTipText     =   "Buscar Cta.Contable"
         Top             =   2805
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   9540
         ToolTipText     =   "Zoom descripción"
         Top             =   3735
         Width           =   240
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
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
         Left            =   7995
         TabIndex        =   45
         Top             =   3750
         Width           =   1440
      End
      Begin VB.Label Label7 
         Caption         =   "Provincia"
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
         TabIndex        =   40
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
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
         Index           =   26
         Left            =   240
         TabIndex        =   39
         Top             =   1950
         Width           =   1050
      End
      Begin VB.Label Label6 
         Caption         =   "Domicilio"
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
         Left            =   240
         TabIndex        =   38
         Top             =   1500
         Width           =   1050
      End
      Begin VB.Label Label5 
         Caption         =   "NIF"
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
         TabIndex        =   37
         Top             =   525
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4200
      Top             =   6960
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   13950
      TabIndex        =   42
      Top             =   8190
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   14580
      TabIndex        =   67
      Top             =   135
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
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
      Begin VB.Menu mnBuscarTarjeta 
         Caption         =   "Buscar &Tarjeta"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmAVNAvnics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Menú: AVNICS                    -+-+
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

'Dim T1 As Single
Private Const IdPrograma = 1416  '*=*=


Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

' *** declarar els formularis als que vaig a cridar ***
'Private WithEvents frmB As frmBuscaGrid
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal 'calendario fecha
Attribute frmC1.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
' *****************************************************


Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim BuscaChekc As String

Dim Sql As String

Private DevfrmCCtas As String


Private Sub cmbAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkAux_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAux(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux(" & Index & ")|"
    End If
End Sub

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOK Then
                If InsertarDesdeForm2(Me, 1) Then
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOK Then
                If ModificaDesdeFormulario2(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
        
        ' *** si n'hi han llínies ***
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    If ModificarLinea Then
                        PosicionarData
                    Else
                        PonFoco txtaux(12)
                    End If
            End Select
        ' **************************
            
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdAux_Click(Index As Integer)
    Set frmC1 = New frmCal
    frmC1.Fecha = Now
    If txtaux(2).Text <> "" Then frmC1.Fecha = CDate(txtaux(2).Text)
    frmC1.Show vbModal
    Set frmC1 = Nothing
    PonFoco txtaux(2)
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then PrimeraVez = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
    
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With
    
    ' Botonera Principal 2
    With Me.Toolbar2
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 41
        .Buttons(2).Image = 44
        .Buttons(3).Image = 47
        .Buttons(4).Image = 42
        .Buttons(5).Image = 36
        .Buttons(6).Image = 40
        
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
    With Me.ToolbarAux
        .HotImageList = frmppal.imgListComun_OM16
        .DisabledImageList = frmppal.imgListComun_BN16
        .ImageList = frmppal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With
    
    
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture 'frmppal.imgListImages16.ListImages(1).Picture
    Next i
    
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture 'frmppal.imgListImages16.ListImages(3).Picture
    Next i
    
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "avnic"
    Ordenacion = " ORDER BY avnic.codavnic, avnic.anoejerc "
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = Conn
    '***** cambiar el nombre de la PK de la cabecera *************
'    Data1.RecordSource = "Select * from " & NombreTabla & " left join movim on avnic.codavnic = movim.codavnic and avnic.anoejerc = movim.anoejerc where false"
    Data1.RecordSource = "Select * from " & NombreTabla & " where false"
    Data1.Refresh
       
    ModoLineas = 0
       
    ' *** si n'hi han combos (capçalera o llínies) ***
    CargaCombo 0
    
    CargaGrid 0, False
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
        text1(0).BackColor = vbLightBlue 'codclien
        ' ****************************************************************************
    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    Limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    ' *** si n'hi han combos a la capçalera ***
    Me.Combo1(0).ListIndex = -1
    Me.chkAux(0).Value = 0

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    b = (Modo = 2)
    If Not Data1.Recordset Is Nothing Then
        DespalzamientoVisible b And (Data1.Recordset.RecordCount > 1)
    End If
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
'    ' ********************************************************
    b = Modo = 2 Or Modo = 0
    
    For i = 0 To text1.Count - 1
        text1(i).Locked = b
        If Modo <> 1 Then
            text1(i).BackColor = vbWhite
        End If
    Next i
    For i = 0 To Combo1.Count - 1
        Combo1(i).Locked = b
    Next i
    
    For i = 0 To imgBuscar.Count - 1
        imgBuscar(i).Enabled = Not b
    Next i
    imgFec(0).Enabled = Not b
    imgFec(1).Enabled = Not b
    
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

      
    ' ****** si n'hi han combos a la capçalera ***********************
    If (Modo = 0) Or (Modo = 2) Or (Modo = 5) Then
        Combo1(0).Enabled = False
        Combo1(0).BackColor = &H80000018 'groc
    ElseIf (Modo = 1) Or (Modo = 3) Or (Modo = 4) Then
        Combo1(0).Enabled = True
        Combo1(0).BackColor = &H80000005 'blanc
    End If
    ' ****************************************************************
    
    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
    End If
    
    b = (Modo = 5) And (NumTabMto = 1) 'And (ModoLineas <> 3)
    
    
    For i = 2 To txtaux.Count - 1
        txtaux(i).Enabled = (Modo = 5) Or (Modo = 1)
    Next i
    'BloqueaTXT txtaux(1), (Modo = 5 And ModoLineas = 2)
    BloqueaTXT txtaux(2), (Modo = 5 And ModoLineas = 2)
    
    Dim anc As Single
    Dim jj As Integer
    anc = DataGridAux(0).top + 240
    For jj = 2 To 6
        txtaux(jj).visible = (Modo = 1)
        txtaux(jj).top = anc
    Next jj
    Me.chkAux(0).visible = (Modo = 1)
    Me.chkAux(0).Enabled = (Modo = 1)
    Me.chkAux(0).top = anc
    
    Me.cmdAux(0).visible = (Modo = 1) Or (Modo = 5 And ModoLineas = 1)
    Me.cmdAux(0).Enabled = (Modo = 1) Or (Modo = 5 And ModoLineas = 1)
    Me.cmdAux(0).top = anc
    
    
    
'    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
'    PonerOpcionesMenu   'Activar opcions de menú según nivell
'                        'de permisos de l'usuari

    PonerModoUsuarioGnral Modo, "ariconta"


EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim i As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(4).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(7).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(12).Enabled = True And Not DeConsulta
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
'    For i = 0 To ToolAux.Count - 1
'        ToolAux(i).Buttons(1).Enabled = b
'        If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
'        ToolAux(i).Buttons(2).Enabled = bAux
'        ToolAux(i).Buttons(3).Enabled = bAux
'    Next i
    
End Sub

Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim b As Boolean
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub


Private Sub Desplazamiento(Index As Integer)
    If Data1.Recordset.EOF Then Exit Sub
    
    Select Case Index
        Case 1
            Data1.Recordset.MoveFirst
        Case 2
            Data1.Recordset.MovePrevious
            If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
        Case 3
            Data1.Recordset.MoveNext
            If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
        Case 4
            Data1.Recordset.MoveLast
    End Select
    PonerCampos
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabem quins camps son els que mos torna
        'Creem una cadena consulta i posem els datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(text1(0), CadenaDevuelta, 1)
        CadB = Aux
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmTra_Actualizar(vValor As Integer)
'Mantenimiento de Colectivos
    
    LimpiarCampos
    text1(0).Text = vValor
    
    FormateaCampo text1(0)
        Modo = 1
        cmdAceptar_Click
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     text1(indice).Text = vCampo
End Sub

' *** si n'hi ha buscar data, posar a les <=== el menor index de les imagens de buscar data ***
' NOTA: ha de coincidir l'index de la image en el del camp a on va a parar el valor
Private Sub imgFec_Click(Index As Integer)

    Set frmC = New frmCal
    frmC.Fecha = Now
    indice = 2
    If Index = 1 Then indice = 17
    If text1(indice).Text <> "" Then frmC.Fecha = CDate(text1(indice).Text)
    Sql = ""
    frmC.Show vbModal
    Set frmC = Nothing
    If Sql <> "" Then
        text1(indice).Text = Sql
        PonFoco text1(indice)
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    'text1(2).Text = Format(vFecha, "dd/mm/yyyy") '<===
    Sql = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC1_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtaux(2).Text = Format(vFecha, "dd/mm/yyyy") '<===
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 26
        frmZ.pTitulo = "Observaciones del Avnic."
        frmZ.pValor = text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonFoco text1(indice)
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
    Combo1(0).ListIndex = -1 'quan busque, per defecte no seleccione res.
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    frmAVNInformes.OpcionListado = 0
    frmAVNInformes.Show vbModal
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1  'Nou
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Búscar
            mnBuscar_Click
        Case 6  'Tots
            mnVerTodos_Click
        Case 8 'Imprimir
            mnImprimir_Click
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonFoco text1(0) ' <===
        text1(0).BackColor = vbLightBlue ' <===
        ' *** si n'hi han combos a la capçalera ***
        For i = 0 To Combo1.Count - 1
            Combo1(i).ListIndex = -1
        Next i
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            text1(kCampo).Text = ""
            text1(kCampo).BackColor = vbLightBlue
            PonFoco text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()
Dim CadB1 As String

    CadB = ObtenerBusqueda2(Me)
    CadB1 = ObtenerBusqueda2(Me, BuscaChekc, 2, "FrameAux0")
    
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Or CadB1 <> "" Then
        If CadB1 <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " LEFT JOIN movim on avnic.codavnic = movim.codavnic and avnic.anoejerc = movim.anoejerc WHERE 1=1 "
        Else
            CadenaConsulta = "select * from " & NombreTabla & " WHERE 1=1 "
        End If
        
        If CadB <> "" Then CadenaConsulta = CadenaConsulta & " and " & CadB
        If CadB1 <> "" Then CadenaConsulta = CadenaConsulta & " and " & CadB1
        
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonFoco text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    Dim Cad As String
        
'    'Cridem al form
'    ' **************** arreglar-ho per a vore lo que es desije ****************
'    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
'    cad = ""
'    cad = cad & ParaGrid(text1(0), 15, "Cód.")
'    cad = cad & ParaGrid(text1(3), 25, "N.I.F.")
'    cad = cad & ParaGrid(text1(4), 60, "Nombre")
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = NombreTabla
'        frmB.vSql = CadB
'        HaDevueltoDatos = False
'        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
'        frmB.vTitulo = "Avnics" ' ***** repasa açò: títol de BuscaGrid *****
'        frmB.vSelElem = 1
'
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha posat valors i tenim que es formulari de búsqueda llavors
'        'tindrem que tancar el form llançant l'event
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha retornat datos, es a decir NO ha retornat datos
'            PonerFoco text1(kCampo)
'        End If
'    End If
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            Cad = Cad & text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonVerTodos()
'Vore tots
    LimpiarCampos 'Neteja els Text1
    CadB = ""
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
'        CadenaConsulta = "Select * from avnic left join movim on avnic.codavnic = movim.codavnic and avnic.anoejerc = movim.anoejerc " & Ordenacion
        CadenaConsulta = "Select * from avnic  " & Ordenacion
        
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
    text1(0).Text = SugerirCodigoSiguienteStr("avnic", "codavnic")
    FormateaCampo text1(0)
    
    text1(1).Text = Format(Now, "yyyy") '
    text1(2).Text = Format(Now, "dd/mm/yyyy") ' Quan afegixc pose en F.Alta i F.Modificación la data actual
    PosicionarCombo Combo1(0), 1
        
    PonFoco text1(0) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
End Sub

Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloqueaTXT text1(0), True
    BloqueaTXT text1(1), True
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonFoco text1(2)
End Sub

Private Sub BotonEliminar()
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar el Avnics?"
    Cad = Cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), "000000")
    Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Avnics", Err.Description
End Sub

Private Sub PonerCampos()
Dim i As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
    text2(9).Text = PonerNombreDeCod(text1(9), "cuentas", "nommacta", "codmacta", "T")
    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    CargaGrid 0, True
    
'    PonerModoOpcionesMenu (Modo)
'    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
Dim v

    Select Case Modo
        Case 1, 3 'Búsqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                ' *** foco al primer camp visible de la capçalera ***
                PonFoco text1(0)

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonFoco text1(0)
        
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
                        'txtAux2(2).text = ""

                    End If


                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0

                    ' *** si n'hi han tabs ***
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        v = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & v)
                        ' ***************************************************************
                    End If

           End Select
            
           PosicionarData
           
    End Select
End Sub

Private Function DatosOK() As Boolean
Dim b As Boolean
'Dim Datos As String
Dim Cta As String
Dim cadMen As String


    On Error GoTo EDatosOK

    DatosOK = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(text1(0)) Then b = False
    End If
    
    
        '[Monica]22/08/2013: añadida la comprobacion de que la cuenta contable sea correcta
        If text1(18).Text = "" Or text1(19).Text = "" Or text1(21).Text = "" Or text1(21).Text = "" Then
            '[Monica]20/11/2013: añadido el codigo de iban
            text1(27).Text = ""
            text1(18).Text = ""
            text1(19).Text = ""
            text1(20).Text = ""
            text1(21).Text = ""
        Else
            Cta = Format(text1(18).Text, "0000") & Format(text1(19).Text, "0000") & Format(text1(20).Text, "00") & Format(text1(21).Text, "0000000000")
            If Val(ComprobarCero(Cta)) = 0 Then
                cadMen = "El avnic no tiene asignada cuenta bancaria."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(Cta) Then
                cadMen = "La cuenta bancaria del avnic no es correcta. ¿ Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonFoco text1(19)
                    b = False
                End If
            Else

'       sustituido por lo de David
                BuscaChekc = ""
                If Me.text1(27).Text <> "" Then BuscaChekc = Mid(text1(27).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, Cta, Cta) Then
                    If Me.text1(27).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.text1(27).Text = BuscaChekc & Cta
                    Else
                        If Mid(text1(27).Text, 3) <> Cta Then
                            Cta = "Calculado : " & BuscaChekc & Cta
                            Cta = "Introducido: " & Me.text1(27).Text & vbCrLf & Cta & vbCrLf
                            Cta = "Error en codigo IBAN" & vbCrLf & Cta & "Continuar?"
                            If MsgBox(Cta, vbQuestion + vbYesNo) = vbNo Then
                                PonFoco text1(27)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
    
    
    ' ************************************************************************************
    DatosOK = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "(codavnic=" & text1(0).Text & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, Cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    Conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE codavnic=" & Data1.Recordset!codavnic
        
    'Eliminar la CAPÇALERA
    vWhere = " WHERE codavnic=" & Data1.Recordset!codavnic & " and anoejerc=" & Data1.Recordset!anoejerc
    Conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Conn.RollbackTrans
        Eliminar = False
    Else
        Conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco text1(Index), Modo
End Sub

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbLightBlue
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbLightBlue Then Combo1(Index).BackColor = vbWhite
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0 'Cod.Avnic
            PonerFormatoEntero text1(0)

        Case 4, 12 'NOMBRE
            text1(Index).Text = UCase(text1(Index).Text)
        
        Case 3, 10, 11 'NIF
            text1(Index).Text = UCase(text1(Index).Text)
'            ValidarNIF text1(Index).Text
                
        Case 2, 17 'Fechas
            PonerFormatoFecha text1(Index)
            
        Case 9 'cuenta contable
            DevfrmCCtas = text1(Index).Text
            If CuentaCorrectaUltimoNivel(DevfrmCCtas, Sql) Then
                text1(Index).Text = DevfrmCCtas
                text2(Index).Text = Sql
            Else
                MsgBox Sql, vbExclamation
                text1(Index).Text = ""
                text2(Index).Text = ""
                PonFoco text1(Index)
            End If
            DevfrmCCtas = ""
            
        Case 22, 23, 24 'IMPORTES
            cadMen = TransformaPuntosComas(text1(Index).Text)
            text1(Index).Text = Format(cadMen, "#,###,##0.00")
        
        Case 25 '% INTERES
            cadMen = TransformaPuntosComas(text1(Index).Text)
            text1(Index).Text = Format(cadMen, "##0.00")
            
        Case 27 ' codigo de iban
            text1(Index).Text = UCase(text1(Index).Text)
            
    End Select
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 18 Or Index = 19 Or Index = 20 Or Index = 21 Then
        Dim Cta As String
        Dim CC As String
        If text1(18).Text <> "" And text1(19).Text <> "" And text1(20).Text <> "" And text1(21).Text <> "" Then
            
            Cta = Format(text1(18).Text, "0000") & Format(text1(19).Text, "0000") & Format(text1(20).Text, "00") & Format(text1(21).Text, "0000000000")
            If Len(Cta) = 20 Then
    '        Text1(42).Text = Calculo_CC_IBAN(cta, Text1(42).Text)
    
                If text1(27).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", Cta, Cta) Then text1(27).Text = "ES" & Cta
                Else
                    CC = CStr(Mid(text1(27).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), Cta, Cta) Then
                        If Mid(text1(27).Text, 3) <> Cta Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & Cta & "]", vbExclamation
                        End If
                    End If
                End If
                
                
            End If
        End If
    End If
    ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYFecha KeyAscii, 15 'fecha de alta
                Case 9: KEYBusqueda KeyAscii, 2 'cuenta contable
                Case 17: KEYFecha KeyAscii, 16 'fecha de vencimiento
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Unload Me
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo(Index As Integer)
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    Combo1(0).Clear
    
    Combo1(0).AddItem "Antigua"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Alta Ejercicio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Cancelada Ejercicio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
End Sub

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
    ' ****************************************************
    
    SepuedeBorrar = True
End Function

Private Sub imgBuscar_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    Set frmCtas = New frmColCtas
    DevfrmCCtas = ""
    frmCtas.DatosADevolverBusqueda = "0"
    frmCtas.Show vbModal
    Set frmCtas = Nothing
    If DevfrmCCtas <> "" Then
        text1(Index + 9).Text = RecuperaValor(DevfrmCCtas, 1)
        text2(Index + 9).Text = RecuperaValor(DevfrmCCtas, 2)
    End If

End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
''Cuentas contables de la Contabilidad
    DevfrmCCtas = CadenaSeleccion
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    SSTab1.Tab = numTab
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************


' *** si n'hi han tabs sense datagrids ***
Private Sub NetejaFrameAux(nom_frame As String)
Dim Control As Object
    
    For Each Control In Me.Controls
        If (Control.Tag <> "") Then
            If (Control.Container.Name = nom_frame) Then
                If TypeOf Control Is TextBox Then
                    Control.Text = ""
                ElseIf TypeOf Control Is ComboBox Then
                    Control.ListIndex = -1
                End If
            End If
        End If
    Next Control

End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " avnic.codavnic=" & Val(text1(0).Text) & " and avnic.anoejerc = " & DBSet(text1(1), "N")
    
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    
        Case 1 'Renovacion Avnics
            frmAVNInformes.OpcionListado = 1
            frmAVNInformes.Show vbModal
        Case 2 'Calculo de intereses
            frmAVNIntereses.Show vbModal
        Case 3 'Contabilizar intereses
            frmAVNVarios.OpcionListado = 0
            frmAVNVarios.Show vbModal
        Case 4 'Cancelacion Avnics
            frmAVNVarios.OpcionListado = 1
            frmAVNVarios.Show vbModal
        Case 5 ' Ayuda modelo 123
            frmAVNVarios.OpcionListado = 2
            frmAVNVarios.Show vbModal
        Case 6 ' Grabacion modelo 193
            frmAVNModelo193.Show vbModal
        
    End Select

End Sub

Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea 0
        Case 2
            BotonModificarLinea 0
        Case 3
            BotonEliminarLinea 0
        Case Else
    End Select
End Sub

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub CargaGrid(Index As Integer, Enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, Enlaza)

    b = DataGridAux(Index).Enabled
    DataGridAux(Index).Enabled = False
    
    AdoAux(Index).ConnectionString = Conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, Enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    DataGridAux(Index).ScrollBars = dbgNone
    AdoAux(Index).Refresh
    Set DataGridAux(Index).DataSource = AdoAux(Index)
    
    DataGridAux(Index).AllowRowSizing = False
    DataGridAux(Index).RowHeight = 350
    
    If PrimeraVez Then
        DataGridAux(Index).ClearFields
        DataGridAux(Index).ReBind
        DataGridAux(Index).Refresh
    End If

    For i = 0 To DataGridAux(Index).Columns.Count - 1
        DataGridAux(Index).Columns(i).AllowSizing = False
    Next i
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    
    'DataGridAux(Index).Enabled = b
'    PrimeraVez = False
    
    Select Case Index
        
        Case 0 'lineas de movimientos
            
            tots = "N||||0|;N|txtaux(1)|T|Año|805|;S|txtaux(2)|T|Fecha Mov.|1400|;S|cmdAux(0)|B|||;S|txtaux(3)|T|Concepto|6300|;"
            tots = tots & "S|txtaux(4)|T|Importe|1905|;S|txtaux(5)|T|Bruto|1905|;S|txtaux(6)|T|Retencion|1905|;"
            tots = tots & "N||||0|;S|chkAux(0)|CB|Int|400|;"
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(3).Alignment = dbgLeft
            DataGridAux(Index).Columns(4).Alignment = dbgRight
            DataGridAux(Index).Columns(5).Alignment = dbgRight
            DataGridAux(Index).Columns(6).Alignment = dbgRight
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

            If (Enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
            
            Else
                For i = 0 To 6
                    txtaux(i).Text = ""
                Next i
            End If
    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
    If Not AdoAux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
        CalcularTotales
    Else
        LimpiarCamposFrame Index
    End If
    ' **********************************************************
      
    
    PonerModoUsuarioGnral Modo, "ariconta"

      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Function CalcularTotales()
Dim Sql As String
Dim Rs As ADODB.Recordset

    text2(0).Text = ""
    text2(1).Text = ""
    text2(2).Text = ""

    Sql = "select sum(coalesce(timporte,0)), sum(coalesce(timport1,0)),sum(coalesce(timport2,0)) "
    Sql = Sql & " from movim where codavnic = " & DBSet(text1(0), "N") & " and anoejerc = " & DBSet(text1(1), "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0).Value, "N") <> 0 Then text2(0).Text = Format(DBLet(Rs.Fields(0).Value, "N"), "###,###,##0.00")
        If DBLet(Rs.Fields(1).Value, "N") <> 0 Then text2(1).Text = Format(DBLet(Rs.Fields(1).Value, "N"), "###,###,##0.00")
        If DBLet(Rs.Fields(2).Value, "N") <> 0 Then text2(2).Text = Format(DBLet(Rs.Fields(2).Value, "N"), "###,###,##0.00")
    End If

End Function


Private Function MontaSQLCarga(Index As Integer, Enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el data1
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 0 ' lineas de movimientos
            tabla = "movim"
            Sql = "SELECT movim.codavnic, movim.anoejerc, movim.fechamov, movim.concepto, movim.timporte, movim.timport1, movim.timport2, movim.intconta, IF(movim.intconta=1,'*','') as dintconta"
            Sql = Sql & " FROM avnic left join movim on avnic.codavnic = movim.codavnic and avnic.anoejerc = movim.anoejerc "
            If Enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE false "
            End If
            Sql = Sql & " ORDER BY 1,2,3"
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function


Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim Eliminar As Boolean
'*=*=
Dim Campo As String
Dim vTabla As String

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'movimientos
            Sql = "¿Seguro que desea eliminar el movimiento?"
            Sql = Sql & vbCrLf & "Código: " & AdoAux(Index).Recordset!codavnic
            Sql = Sql & vbCrLf & "Año   : " & AdoAux(Index).Recordset!anoejerc
            Sql = Sql & vbCrLf & "Fecha : " & Format(AdoAux(Index).Recordset!Fechamov, "dd/mm/yyyy")
            
            If MsgBoxA(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM movim "
                Sql = Sql & Replace(vWhere, "avnic.", "movim.") & " AND anoejerc= " & AdoAux(Index).Recordset!anoejerc
                Sql = Sql & " AND fechamov= " & DBSet(AdoAux(Index).Recordset!Fechamov, "F")
            End If
            
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        Conn.Execute Sql
        
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGrid Index, True
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
    End If
    
    ModoLineas = 0
    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim i As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloqueaTXT text1(0), True

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vTabla = "movim"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'movimientos de avnics
                    txtaux(0).Text = text1(0).Text 'codigo avnic
                    txtaux(1).Text = text1(1).Text 'ano ejercicio
                    
                    For i = 2 To 6
                        txtaux(i).Text = ""
                    Next i
                    PonFoco txtaux(2)
                    
            End Select
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
       
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la capçalera ***
    BloqueaTXT text1(0), True
  
    Select Case Index
        Case 0 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, i
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
    End Select
    
    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 0 'calidades
            For J = 0 To 6
                txtaux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            ' ***** canviar-ho pel nom del camp del combo *********
            
            Me.chkAux(0).Value = Me.AdoAux(Index).Recordset!intconta
            
            
'            For i = 0 To 2
'                BloqueaTXT txtaux(i), False
'            Next i
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'año ejercicio
            PonFoco txtaux(3)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'movimientos de avnics
            For jj = 2 To 6
                txtaux(jj).visible = b
                txtaux(jj).top = alto
            Next jj
            Me.chkAux(0).visible = b
            Me.chkAux(0).top = alto
            
'            Me.cmdAux(0).visible = b
            Me.cmdAux(0).top = alto
            
            
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
   If Not txtaux(Index).MultiLine Then ConseguirFocoLin txtaux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtaux(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtaux(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 1 Or Modo = 5 Then
                Select Case Index
                    Case 2: KEYFecha KeyAscii, 0 'fecha de alta
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'calibres
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            Select Case NumTabMto
                Case 0 ' *** els index de les llinies en grid (en o sense tab) ***
                     CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
            End Select
           
'            SituarTab (NumTabMto + 1)
        End If
    End If
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim v As Integer
'*=*=
Dim Campo As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'calibres
    End Select
    
    ModificarLinea = False
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ModoLineas = 0
            
            v = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
            CargaGrid NumTabMto, True
            
            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(NumTabMto)
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & v)
            
            LLamaLineas NumTabMto, 0
            ModificarLinea = True
        End If
    End If
End Function

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    ' ******************************************************************************
    DatosOkLlin = b
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Famia As String

    If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub
    
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 1
            PonerFormatoEntero txtaux(Index)
            
        Case 2 'FECHA
            PonerFormatoFecha txtaux(Index)
        
        Case 3 'CONCEPTO
            txtaux(Index).Text = UCase(txtaux(Index).Text)
            
        Case 4, 5, 6 'IMPORTE
            If txtaux(Index).Text <> "" Then PonerFormatoDecimal txtaux(Index), 3
            
    End Select
End Sub


