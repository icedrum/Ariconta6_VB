VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos cuentas"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11100
   Icon            =   "frmCuentas2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11100
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.Frame FrGranEmpresa 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   63
      Top             =   4920
      Width           =   8055
      Begin VB.CommandButton cmdCopiarDatos 
         Height          =   375
         Index           =   1
         Left            =   5250
         Picture         =   "frmCuentas2.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtRegularizacion 
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
         Left            =   5760
         TabIndex        =   39
         Top             =   300
         Width           =   1725
      End
      Begin VB.Label Label4 
         Caption         =   "Grandes empresas.   Regularizaci�n grupos 8 y 9"
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
         TabIndex        =   64
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   150
      TabIndex        =   43
      Top             =   960
      Width           =   10755
      Begin TabDlg.SSTab SSTab1 
         Height          =   6135
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   10821
         _Version        =   393216
         Tab             =   1
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
         TabCaption(0)   =   "Datos cuentas"
         TabPicture(0)   =   "frmCuentas2.frx":685E
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label1(11)"
         Tab(0).Control(1)=   "Label1(2)"
         Tab(0).Control(2)=   "Label1(7)"
         Tab(0).Control(3)=   "Label1(3)"
         Tab(0).Control(4)=   "Label1(4)"
         Tab(0).Control(5)=   "Label1(5)"
         Tab(0).Control(6)=   "Label1(6)"
         Tab(0).Control(7)=   "Label1(8)"
         Tab(0).Control(8)=   "Label1(9)"
         Tab(0).Control(9)=   "Label1(10)"
         Tab(0).Control(10)=   "Label1(22)"
         Tab(0).Control(11)=   "imgppal(2)"
         Tab(0).Control(12)=   "imgWeb(0)"
         Tab(0).Control(13)=   "Image1(2)"
         Tab(0).Control(14)=   "Image1(3)"
         Tab(0).Control(15)=   "Label1(14)"
         Tab(0).Control(16)=   "Image1(4)"
         Tab(0).Control(17)=   "Label1(15)"
         Tab(0).Control(18)=   "Image1(6)"
         Tab(0).Control(19)=   "ToolbarMail"
         Tab(0).Control(20)=   "Text1(12)"
         Tab(0).Control(21)=   "Check1"
         Tab(0).Control(22)=   "Text1(2)"
         Tab(0).Control(23)=   "Text1(3)"
         Tab(0).Control(24)=   "Text1(4)"
         Tab(0).Control(25)=   "Text1(5)"
         Tab(0).Control(26)=   "Text1(6)"
         Tab(0).Control(27)=   "Text1(7)"
         Tab(0).Control(28)=   "Text1(8)"
         Tab(0).Control(29)=   "Text1(9)"
         Tab(0).Control(30)=   "Text1(10)"
         Tab(0).Control(31)=   "Text1(23)"
         Tab(0).Control(32)=   "Text2(2)"
         Tab(0).Control(33)=   "Text2(3)"
         Tab(0).Control(34)=   "Text1(30)"
         Tab(0).Control(35)=   "Text1(27)"
         Tab(0).Control(36)=   "Text2(27)"
         Tab(0).Control(37)=   "Check3"
         Tab(0).ControlCount=   38
         TabCaption(1)   =   "Tesorer�a"
         TabPicture(1)   =   "frmCuentas2.frx":687A
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label1(21)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label1(24)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Image1(1)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Image1(0)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label1(26)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label1(27)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "imgppal(4)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label1(28)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Text1(29)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Text1(25)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Text1(26)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Text2(0)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Text2(1)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Text1(13)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "Text1(14)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "Text1(15)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "Text1(16)"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "Frame4"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "Text1(31)"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "Text1(32)"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "Text1(33)"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "Text1(34)"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).ControlCount=   22
         TabCaption(2)   =   "Departamentos"
         TabPicture(2)   =   "frmCuentas2.frx":6896
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FrameAux2"
         Tab(2).ControlCount=   1
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
            Index           =   34
            Left            =   6420
            MaxLength       =   4
            TabIndex        =   22
            Text            =   "Text"
            Top             =   720
            Width           =   720
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
            Index           =   33
            Left            =   5640
            MaxLength       =   4
            TabIndex        =   21
            Text            =   "Text"
            Top             =   720
            Width           =   720
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
            Index           =   32
            Left            =   8940
            TabIndex        =   27
            Tag             =   "Fecha Mandato|F|S|||cuentas|SEPA_FecFirma|dd/mm/yyyy||"
            Text            =   "0000000000"
            Top             =   2430
            Width           =   1305
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Cuenta M�ltiple"
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
            Left            =   -66720
            TabIndex        =   97
            Tag             =   "Cuenta M�ltiple|N|S|||cuentas|esctamultiple|||"
            Top             =   3930
            Width           =   1875
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
            Index           =   31
            Left            =   2430
            MaxLength       =   35
            TabIndex        =   26
            Tag             =   "Cta banco|T|S|||cuentas|SEPA_Refere|||"
            Top             =   2430
            Width           =   3645
         End
         Begin VB.Frame Frame4 
            Caption         =   "Operaciones Aseguradas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   90
            TabIndex        =   86
            Top             =   3330
            Width           =   10455
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
               Index           =   19
               Left            =   5610
               TabIndex        =   31
               Tag             =   "Imp1|N|S|||cuentas|credisol|#0.00||"
               Top             =   1050
               Width           =   1305
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
               Left            =   5610
               TabIndex        =   30
               Tag             =   "Fl|F|S|||cuentas|fecsolic|dd/mm/yyyy||"
               Top             =   540
               Width           =   1305
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
               Index           =   28
               Left            =   1860
               TabIndex        =   29
               Tag             =   "F. baja credito|F|S|||cuentas|fecbajcre|dd/mm/yyyy||"
               Top             =   1050
               Width           =   1305
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
               Index           =   17
               Left            =   1860
               MaxLength       =   10
               TabIndex        =   28
               Tag             =   "Raz�n social|T|S|||cuentas|numpoliz|||"
               Top             =   540
               Width           =   1305
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
               Height          =   795
               Index           =   22
               Left            =   1860
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   34
               Tag             =   "Raz�n social|T|S|||cuentas|observa|||"
               Text            =   "frmCuentas2.frx":68B2
               Top             =   1710
               Width           =   8235
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
               Index           =   21
               Left            =   8820
               TabIndex        =   33
               Tag             =   "lmpor1|N|S|||cuentas|credicon|#0.00||"
               Top             =   1020
               Width           =   1305
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
               Index           =   20
               Left            =   8820
               TabIndex        =   32
               Tag             =   "Fecha|F|S|||cuentas|fecconce|dd/mm/yyyy||"
               Text            =   "0000000000"
               Top             =   540
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "N� Poliza"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   16
               Left            =   150
               TabIndex        =   95
               Top             =   570
               Width           =   915
            End
            Begin VB.Image Image1 
               Height          =   240
               Index           =   5
               Left            =   1860
               Top             =   1440
               Width           =   240
            End
            Begin VB.Label Label1 
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
               Height          =   285
               Index           =   13
               Left            =   150
               TabIndex        =   94
               Top             =   1500
               Width           =   1665
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
               Height          =   285
               Index           =   17
               Left            =   4320
               TabIndex        =   93
               Top             =   570
               Width           =   915
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
               Height          =   285
               Index           =   18
               Left            =   4320
               TabIndex        =   92
               Top             =   1080
               Width           =   915
            End
            Begin VB.Image imgppal 
               Height          =   240
               Index           =   0
               Left            =   5310
               Picture         =   "frmCuentas2.frx":68BD
               Top             =   540
               Width           =   240
            End
            Begin VB.Image imgppal 
               Height          =   240
               Index           =   3
               Left            =   1530
               Picture         =   "frmCuentas2.frx":6948
               Top             =   1050
               Width           =   240
            End
            Begin VB.Label Label3 
               Caption         =   "CONCEDIDO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   2
               Left            =   8820
               TabIndex        =   91
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label Label3 
               Caption         =   "SOLICITADO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   1
               Left            =   5670
               TabIndex        =   90
               Top             =   180
               Width           =   1395
            End
            Begin VB.Image imgppal 
               Height          =   240
               Index           =   1
               Left            =   8520
               Picture         =   "frmCuentas2.frx":69D3
               Top             =   540
               Width           =   240
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
               Height          =   285
               Index           =   20
               Left            =   7560
               TabIndex        =   89
               Top             =   1050
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
               Height          =   195
               Index           =   19
               Left            =   7590
               TabIndex        =   88
               Top             =   570
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "Fecha BAJA"
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
               Index           =   25
               Left            =   150
               TabIndex        =   87
               Top             =   1080
               Width           =   915
            End
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
            Index           =   27
            Left            =   -73500
            TabIndex        =   84
            Top             =   4050
            Width           =   4245
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
            Index           =   27
            Left            =   -74850
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Contrapartida Habitual|T|S|||cuentas|codcontrhab|||"
            Text            =   "0000000000"
            Top             =   4050
            Width           =   1305
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
            Index           =   16
            Left            =   4836
            MaxLength       =   4
            TabIndex        =   20
            Text            =   "Text"
            Top             =   720
            Width           =   720
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
            Index           =   15
            Left            =   4034
            MaxLength       =   4
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   720
            Width           =   720
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
            Index           =   14
            Left            =   3232
            MaxLength       =   4
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   720
            Width           =   720
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
            Index           =   13
            Left            =   2430
            MaxLength       =   4
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   720
            Width           =   720
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
            Index           =   30
            Left            =   -74850
            MaxLength       =   2
            TabIndex        =   14
            Tag             =   "Iva|N|S|||cuentas|codigiva|00||"
            Text            =   "Text1"
            Top             =   4680
            Width           =   660
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
            Index           =   3
            Left            =   -74040
            TabIndex        =   81
            Top             =   4680
            Width           =   4785
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
            Index           =   2
            Left            =   -68550
            TabIndex        =   80
            Top             =   2670
            Width           =   3675
         End
         Begin VB.Frame FrameAux2 
            BorderStyle     =   0  'None
            Height          =   5160
            Left            =   -74880
            TabIndex        =   73
            Top             =   630
            Width           =   10320
            Begin VB.Frame FrameToolAux 
               Height          =   555
               Left            =   90
               TabIndex        =   78
               Top             =   0
               Width           =   1545
               Begin MSComctlLib.Toolbar ToolbarAux 
                  Height          =   330
                  Left            =   210
                  TabIndex        =   79
                  Top             =   150
                  Width           =   1125
                  _ExtentX        =   1984
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
            Begin VB.TextBox txtAux3 
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
               Height          =   350
               Index           =   1
               Left            =   750
               MaxLength       =   3
               TabIndex        =   75
               Tag             =   "Departamento|N|N|||departamentos|dpto|000|S|"
               Text            =   "dpto"
               Top             =   3405
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox txtAux3 
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
               Height          =   350
               Index           =   0
               Left            =   300
               MaxLength       =   10
               TabIndex        =   74
               Tag             =   "Cuenta|T|N|||departamentos|codmacta||S|"
               Text            =   "Cuenta"
               Top             =   3420
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox txtAux3 
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
               Height          =   350
               Index           =   2
               Left            =   1290
               MaxLength       =   30
               TabIndex        =   76
               Tag             =   "Descripcion|T|N|||departamentos|descripcion|||"
               Text            =   "descripci"
               Top             =   3420
               Visible         =   0   'False
               Width           =   5235
            End
            Begin MSAdodcLib.Adodc AdoAux 
               Height          =   375
               Index           =   2
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
               Bindings        =   "frmCuentas2.frx":6A5E
               Height          =   3225
               Index           =   2
               Left            =   90
               TabIndex        =   77
               Top             =   870
               Width           =   9930
               _ExtentX        =   17515
               _ExtentY        =   5689
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
            Index           =   1
            Left            =   3900
            TabIndex        =   70
            Top             =   1860
            Width           =   6345
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
            Index           =   0
            Left            =   3900
            TabIndex        =   69
            Top             =   1290
            Width           =   6345
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
            Index           =   26
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   25
            Tag             =   "Cta banco|T|S|||cuentas|ctabanco|||"
            Top             =   1860
            Width           =   1425
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
            Index           =   25
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   24
            Tag             =   "For. pago|N|S|||cuentas|forpa|||"
            Text            =   "123456789012345678901234567890"
            Top             =   1290
            Width           =   1425
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
            Index           =   23
            Left            =   -66630
            MaxLength       =   15
            TabIndex        =   15
            Tag             =   "NIF|F|S|||cuentas|fecbloq|||"
            Text            =   "Text1"
            Top             =   4710
            Width           =   1755
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
            Height          =   705
            Index           =   10
            Left            =   -74850
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Tag             =   "Observaciones|T|S|||cuentas|obsdatos|||"
            Text            =   "frmCuentas2.frx":6A76
            Top             =   5310
            Width           =   10005
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
            Left            =   -69120
            MaxLength       =   50
            TabIndex        =   12
            Tag             =   "Direccion web|T|S|||cuentas|webdatos|||"
            Text            =   "Text1"
            Top             =   3360
            Width           =   4260
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
            Index           =   8
            Left            =   -74850
            MaxLength       =   40
            TabIndex        =   11
            Tag             =   "E-Mail|T|S|||cuentas|maidatos|||"
            Text            =   "Text1"
            Top             =   3360
            Width           =   5625
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
            Index           =   7
            Left            =   -67500
            MaxLength       =   15
            TabIndex        =   4
            Tag             =   "NIF|T|S|||cuentas|nifdatos|||"
            Text            =   "Text1"
            Top             =   675
            Width           =   1845
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
            Index           =   6
            Left            =   -73470
            MaxLength       =   30
            TabIndex        =   9
            Tag             =   "Provincia|T|S|||cuentas|desprovi|||"
            Text            =   "Text1"
            Top             =   2670
            Width           =   4260
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
            Index           =   5
            Left            =   -74850
            MaxLength       =   50
            TabIndex        =   7
            Tag             =   "Poblaci�n|T|S|||cuentas|despobla|||"
            Text            =   "12345678901234567890123456789012345678901234567890"
            Top             =   1950
            Width           =   9990
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
            Index           =   4
            Left            =   -74835
            MaxLength       =   6
            TabIndex        =   8
            Tag             =   "Cod. Postal|T|S|||cuentas|codposta|||"
            Text            =   "Text1"
            Top             =   2670
            Width           =   1305
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
            Index           =   3
            Left            =   -74850
            MaxLength       =   50
            TabIndex        =   6
            Tag             =   "Domicilio|T|S|||cuentas|dirdatos|||"
            Text            =   "12345678901234567890123456789012345678901234567890"
            Top             =   1320
            Width           =   9990
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
            Index           =   2
            Left            =   -74850
            MaxLength       =   60
            TabIndex        =   3
            Tag             =   "Raz�n social|T|S|||cuentas|razosoci|||"
            Top             =   675
            Width           =   7305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "347"
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
            Left            =   -65580
            TabIndex        =   5
            Tag             =   "Modelo|N|S|||cuentas|model347|||"
            Top             =   720
            Width           =   1005
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
            Index           =   12
            Left            =   -69120
            MaxLength       =   2
            TabIndex        =   10
            Tag             =   "Pais|T|S|||cuentas|codpais|||"
            Text            =   "Text1"
            Top             =   2670
            Width           =   540
         End
         Begin MSComctlLib.Toolbar ToolbarMail 
            Height          =   390
            Left            =   -74070
            TabIndex        =   72
            Top             =   3000
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Envio Mail"
               EndProperty
            EndProperty
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
            Index           =   29
            Left            =   6480
            MaxLength       =   40
            TabIndex        =   23
            Tag             =   "IBAN|T|S|||cuentas|iban|||"
            Text            =   "Text1"
            Top             =   1290
            Width           =   3720
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha de mandato"
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
            Index           =   28
            Left            =   6570
            TabIndex        =   98
            Top             =   2460
            Width           =   1995
         End
         Begin VB.Image imgppal 
            Height          =   240
            Index           =   4
            Left            =   8640
            Picture         =   "frmCuentas2.frx":6A7C
            Top             =   2430
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "N� Referencia"
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
            Left            =   210
            TabIndex        =   96
            Top             =   2430
            Width           =   1455
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   6
            Left            =   -72540
            Top             =   3750
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Contrapartida habitual"
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
            Index           =   15
            Left            =   -74850
            TabIndex        =   85
            Top             =   3780
            Width           =   2355
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN"
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
            Left            =   210
            TabIndex        =   83
            Top             =   720
            Width           =   585
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   4
            Left            =   -73260
            Top             =   5040
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "IVA"
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
            Index           =   14
            Left            =   -74820
            TabIndex        =   82
            Top             =   4440
            Width           =   405
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   3
            Left            =   -74400
            Top             =   4440
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   2
            Left            =   -68640
            Top             =   2370
            Width           =   240
         End
         Begin VB.Image imgWeb 
            Height          =   240
            Index           =   0
            Left            =   -67620
            Picture         =   "frmCuentas2.frx":6B07
            Top             =   3060
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   2100
            Top             =   1320
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   2100
            Top             =   1890
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta banco"
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
            Index           =   24
            Left            =   210
            TabIndex        =   68
            Top             =   1860
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Forma pago"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   21
            Left            =   210
            TabIndex        =   67
            Top             =   1290
            Width           =   1425
         End
         Begin VB.Image imgppal 
            Height          =   240
            Index           =   2
            Left            =   -65130
            Picture         =   "frmCuentas2.frx":7091
            Top             =   4410
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Bloqueo"
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
            Index           =   22
            Left            =   -66630
            TabIndex        =   61
            Top             =   4440
            Width           =   1440
         End
         Begin VB.Label Label1 
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
            Index           =   10
            Left            =   -74835
            TabIndex        =   60
            Top             =   5040
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Direcci�n web"
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
            Index           =   9
            Left            =   -69135
            TabIndex        =   59
            Top             =   3060
            Width           =   1905
         End
         Begin VB.Label Label1 
            Caption         =   "e-MAIL"
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
            Left            =   -74850
            TabIndex        =   58
            Top             =   3060
            Width           =   855
         End
         Begin VB.Label Label1 
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
            Index           =   6
            Left            =   -73455
            TabIndex        =   57
            Top             =   2430
            Width           =   1905
         End
         Begin VB.Label Label1 
            Caption         =   "Poblaci�n"
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
            Index           =   5
            Left            =   -74850
            TabIndex        =   56
            Top             =   1710
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "C.Postal"
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
            Left            =   -74850
            TabIndex        =   55
            Top             =   2430
            Width           =   1065
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   3
            Left            =   -74865
            TabIndex        =   54
            Top             =   1080
            Width           =   1170
         End
         Begin VB.Label Label1 
            Caption         =   "N.I.F."
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
            Index           =   7
            Left            =   -67500
            TabIndex        =   53
            Top             =   420
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Raz�n social"
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
            Left            =   -74850
            TabIndex        =   52
            Top             =   390
            Width           =   1725
         End
         Begin VB.Label Label1 
            Caption         =   "Pa�s"
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
            Left            =   -69120
            TabIndex        =   51
            Top             =   2430
            Width           =   465
         End
      End
   End
   Begin VB.CommandButton cmdCopiarDatos 
      Height          =   300
      Index           =   2
      Left            =   1410
      Picture         =   "frmCuentas2.frx":711C
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   "copiar cuentas OTRA SECCION/EMPRESA"
      Top             =   90
      Width           =   300
   End
   Begin VB.CommandButton cmdCopiarDatos 
      Height          =   300
      Index           =   0
      Left            =   1020
      Picture         =   "frmCuentas2.frx":D96E
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Copiar cuenta"
      Top             =   90
      Width           =   300
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Left            =   8130
      TabIndex        =   46
      Top             =   210
      Width           =   1500
      Begin VB.CheckBox chkUltimo 
         Caption         =   "Ultimo nivel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   2
         Top             =   210
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   11
         Left            =   210
         MaxLength       =   30
         TabIndex        =   48
         Tag             =   "Ultimo nbivel|T|N|||cuentas|apudirec|||"
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   3900
      End
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
      Left            =   9810
      TabIndex        =   36
      Top             =   7140
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   8520
      TabIndex        =   35
      Top             =   7140
      Width           =   1035
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
      Left            =   9810
      TabIndex        =   40
      Top             =   7140
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   120
      TabIndex        =   44
      Top             =   7050
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
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
         TabIndex        =   45
         Top             =   180
         Width           =   2955
      End
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
      Index           =   0
      Left            =   120
      MaxLength       =   15
      TabIndex        =   0
      Tag             =   "Codigo cuenta|T|N|||cuentas|codmacta||S|"
      Top             =   390
      Width           =   1575
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
      Index           =   1
      Left            =   1770
      MaxLength       =   60
      TabIndex        =   1
      Tag             =   "Denominaci�n cuenta|T|N|||cuentas|nommacta|||"
      Top             =   390
      Width           =   5940
   End
   Begin VB.CheckBox Check2 
      Height          =   375
      Left            =   240
      TabIndex        =   37
      Top             =   2220
      Width           =   345
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
      Height          =   1635
      Index           =   24
      Left            =   1770
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   38
      Text            =   "frmCuentas2.frx":141C0
      Top             =   2970
      Width           =   6405
   End
   Begin VB.Label lbl347 
      Caption         =   "Ofertar la marca de 347 para las cuentas del subgrupo"
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
      Left            =   600
      TabIndex        =   66
      Top             =   2280
      Width           =   7350
   End
   Begin VB.Label Label1 
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
      Index           =   23
      Left            =   255
      TabIndex        =   62
      Top             =   3000
      Width           =   1905
   End
   Begin VB.Label Label2 
      Caption         =   "NO es cuenta �ltimo nivel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   47
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Denominaci�n"
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
      Left            =   1830
      TabIndex        =   42
      Top             =   120
      Width           =   3465
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
      Index           =   0
      Left            =   150
      TabIndex        =   41
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "frmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmCta As frmBasico2
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmFPag As frmBasico2
Attribute frmFPag.VB_VarHelpID = -1
Private WithEvents frmBan As frmBasico2
Attribute frmBan.VB_VarHelpID = -1
Private WithEvents frmPais As frmBasico2
Attribute frmPais.VB_VarHelpID = -1
Private WithEvents frmIVA As frmBasico2
Attribute frmIVA.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1

Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private Const IdPrograma = 201


Public CodCta As String
Public vModo As Byte
' 0.- Ver solo
' 1.- A�adir
' 2.- Modificar
' 3.- Buscar

' 5.- Lineas

Public Event DatoSeleccionado(CadenaSeleccion As String)
Private kCampo As Integer
Dim Sql As String


Dim ModoLineas As Byte
    ' 1 = insertar
    ' 2 = modificar
    ' 3 = eliminar


'Para saber si han bloquedao una cuenta, si tienen que avisar de
Private varBloqCta As String
Private PrimeraVez  As Boolean

Dim Modo As Byte
Dim Indice As Integer


Private Sub cboPais_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check3_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub cmdAceptar_Click()
    Dim I As Integer
    Dim B As Boolean
    Dim v As Long
    
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    If Modo = 5 Then
        Select Case ModoLineas
            Case 1 ' insertar
                If DatosOkLin("FrameAux2") Then
'                    TerminaBloquear
                    If InsertarDesdeForm2(Me, 2, "FrameAux2") Then
                        CargaGrid 2, True
                        BotonAnyadirLinea 2
                        
                    End If
                End If
                
            Case 2 ' modificar
                If DatosOkLin("FrameAux2") Then
                    If ModificaDesdeFormulario2(Me, 2, "FrameAux2") Then
                
                        ModoLineas = 0
            
                        v = AdoAux(2).Recordset.Fields(1) 'el 2 es el n� de departamento
                        CargaGrid 2, True
            
                        ' *** si n'hi han tabs ***
                        Me.SSTab1.Tab = 2
            
                        ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
                        DataGridAux(2).SetFocus
                        AdoAux(2).Recordset.Find (AdoAux(2).Recordset.Fields(1).Name & " =" & v)
                        ' ***********************************************************
            
                        LLamaLineas 2, 0
                        Modo = 2
                
                        'Vamos a ver los datos
                        PonerCampos ""
                        
                        lblIndicador.Caption = "Ver cuenta"
                        cmdCancelar.SetFocus
                    End If
                End If
            
            Case 3 ' eliminar
        
        End Select
    
        Screen.MousePointer = vbDefault
        
        Exit Sub
    End If
    
    Select Case vModo
    Case 1
        If DatosOK Then
            '-----------------------------------------
            'Hacemos insertar
            
            'estoy aqui, da problemas, creo que es el  chcek para indicar si es ultimomnivel o no
            If InsertarDesdeForm2(Me, 1) Then
                
                If Len(Text1(0).Text) = vEmpresa.DigitosUltimoNivel Then
                           
                    If vParam.EnlazaCtasMultibase <> "" Then
                        Screen.MousePointer = vbHourglass
                        lblIndicador.Caption = "ENLACE GESTION"
                        Me.Refresh
                        DoEvents
                               'Cta                     nomcta              NIF
                        Sql = Text1(0).Text & "|" & Text1(1).Text & "|" & Text1(7).Text & "|"
                        HacerEnlaceMultibase 0, Sql
                    
                    End If
                    
                    
                    If Text1(23).Text <> varBloqCta Then
                        'Siginifica que el bloqueo de cuenta ha sido modificado
                        Sql = "Hay conectados los siguientes PCs. Deberian reiniciar." & vbCrLf
                        If UsuariosConectados(Sql) Then
                        
                        End If
                        'Volvemos a leer las cuentas bloqueadas
                        vParam.ObtenerCuentasBloqueadas
                    End If
                    
''''                    'Si es cuenta de ultimo nivel. Compruebo si la insercion tiene que ver
''''                    'con la variable GRAN EMPRESA
''''                    If Val(Mid(Text1(0).Text, 1, 1)) >= 8 Then
''''                        If Not vEmpresa.GranEmpresa Then vEmpresa.GranEmpresa = True
''''                    End If
                    
                End If
                'Salimos
                CadenaDesdeOtroForm = Text1(0).Text
                Unload Me
               
               
            End If
        End If
    Case 2
            'Modificar
            If DatosOK Then
                '-----------------------------------------
                'Hacemos modificar
                
                'If ModificaDesdeFormulario2(Me, 1) Then
                If ModificarRegistro Then
                    'SOLO ACTAULZIAMOS CUENTAS DE ULTIMO NIVEL
                    If Len(Text1(0).Text) = vEmpresa.DigitosUltimoNivel Then
                        If vParam.EnlazaCtasMultibase <> "" Then
                            Screen.MousePointer = vbHourglass
                            lblIndicador.Caption = "ENLACE GESTION"
                            Me.Refresh
                            DoEvents
                                   'Cta                     nomcta              NIF
                            Sql = Text1(0).Text & "|" & Text1(1).Text & "|" & Text1(7).Text & "|"
                            HacerEnlaceMultibase 1, Sql
                        
                        End If
                    End If
                    
                    If Text1(23).Text <> varBloqCta Then
                        'Siginifica que el bloqueo de cuenta ha sido modificado
                        Sql = "Hay conectados los siguientes PCs. Deberian reiniciar." & vbCrLf
                        If UsuariosConectados(Sql) Then
                        
                        End If
                        'Volvemos a leer las cuentas bloqueadas
                        vParam.ObtenerCuentasBloqueadas
                    End If
                    CadenaDesdeOtroForm = Text1(0).Text
                    Unload Me
                End If
            End If
    Case 3
            'Si hay busqueda
            CadenaDesdeOtroForm = ""
            Sql = ObtenerBusqueda2(Me, , 1)
            
            If Text1(13).Text <> "" Then
                If Sql <> "" Then Sql = Sql & " and "
                Sql = Sql & "mid(iban,1,4) = " & DBSet(Text1(13).Text, "T")
            End If
            If Text1(14).Text <> "" Then
                If Sql <> "" Then Sql = Sql & " and "
                Sql = Sql & "mid(iban,5,4) = " & DBSet(Text1(14).Text, "T")
            End If
            If Text1(15).Text <> "" Then
                If Sql <> "" Then Sql = Sql & " and "
                Sql = Sql & "mid(iban,9,4) = " & DBSet(Text1(15).Text, "T")
            End If
            If Text1(16).Text <> "" Then
                If Sql <> "" Then Sql = Sql & " and "
                Sql = Sql & "mid(iban,13,4) = " & DBSet(Text1(16).Text, "T")
            End If
            If Text1(33).Text <> "" Then
                If Sql <> "" Then Sql = Sql & " and "
                Sql = Sql & "mid(iban,17,4) = " & DBSet(Text1(33).Text, "T")
            End If
            If Text1(34).Text <> "" Then
                If Sql <> "" Then Sql = Sql & " and "
                Sql = Sql & "mid(iban,21,4) = " & DBSet(Text1(34).Text, "T")
            End If

            
            Dim Sql2 As String
            Sql2 = ObtenerBusqueda2(Me, , 2, "FrameAux2")
            If Sql2 <> "" Then
                If Sql <> "" Then Sql = Sql & " and "
                
                Sql = Sql & " cuentas.codmacta in (select codmacta from departamentos where " & Sql2 & ")"
            End If
            
            If Sql <> "" Then
                CadenaDesdeOtroForm = Sql
                Unload Me
            Else
                MsgBox "Especifique algun campo de b�squeda", vbExclamation
            End If
            
    Case 5 ' a�adir lineas
            
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Function ModificarRegistro() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim B As Boolean

    ModificarRegistro = False
    
    Conn.BeginTrans
    
    B = ModificaDesdeFormulario2(Me, 1)
         
    If B Then
        If Check3.Value = 1 Then
            ' modificacion de facturas de clientes
            Sql = "update factcli set nommacta = " & DBSet(Text1(1).Text, "T")
            Sql = Sql & ", dirdatos = " & DBSet(Text1(3).Text, "T")
            Sql = Sql & ", codpobla = " & DBSet(Text1(4).Text, "T")
            Sql = Sql & ", despobla = " & DBSet(Text1(5).Text, "T")
            Sql = Sql & ", desprovi = " & DBSet(Text1(6).Text, "T")
            Sql = Sql & ", nifdatos = " & DBSet(Text1(7).Text, "T")
            Sql = Sql & ", codpais = " & DBSet(Text1(12).Text, "T")
            Sql = Sql & " where codmacta = " & DBSet(Text1(0).Text, "T")
            
            Conn.Execute Sql
            
            ' modificacion de facturas de proveedor
            Sql = "update factpro set nommacta = " & DBSet(Text1(1).Text, "T")
            Sql = Sql & ", dirdatos = " & DBSet(Text1(3).Text, "T")
            Sql = Sql & ", codpobla = " & DBSet(Text1(4).Text, "T")
            Sql = Sql & ", despobla = " & DBSet(Text1(5).Text, "T")
            Sql = Sql & ", desprovi = " & DBSet(Text1(6).Text, "T")
            Sql = Sql & ", nifdatos = " & DBSet(Text1(7).Text, "T")
            Sql = Sql & ", codpais = " & DBSet(Text1(12).Text, "T")
            Sql = Sql & " where codmacta = " & DBSet(Text1(0).Text, "T")
            
            Conn.Execute Sql
        End If
    End If
    
    ModificarRegistro = B
    Conn.CommitTrans
    Exit Function

eModificarRegistro:
    MuestraError Err.Number, "Modifica Registro", Err.Description
    Conn.RollbackTrans
End Function


Private Sub cmdCancelar_Click()
Unload Me
End Sub



'0.- Cuenta normal
'1.- Forpa
'2.- Cuenta bancaria
Private Sub AbrirSelCuentas2(vOpcion As Byte, OtraSeccion As String)

    Screen.MousePointer = vbHourglass
    Select Case vOpcion
    Case 0
        Set frmCta = New frmBasico2
        AyudaCuentas frmCta, , "cuentas.apudirec = ""S"""
        Set frmCta = Nothing
    End Select

End Sub


Private Sub cmdCopiarDatos_Click(Index As Integer)
Dim EmpresaSt As String

    If Index = 0 Or Index = 2 Then
       If Not Frame1.Visible Then
            MsgBox "Solo se pueden copiar datos para las cuentas a ultimo nivel", vbExclamation
            Exit Sub
        End If
    Else
        'Para poner contra que cuenta regularizan las 8 y 9
        
    End If
    
    EmpresaSt = ""
    
    If Index = 2 Then
        'Abrimos para que seleccione las empresas
            Sql = ""
            CadenaDesdeOtroForm = "NO"  'Para que no seleccione ninguna empresa por defecto
            frmMensajes.Opcion = 4
            frmMensajes.Show vbModal
            If CadenaDesdeOtroForm = "" Then Exit Sub
            NumRegElim = RecuperaValor(CadenaDesdeOtroForm, 1)
            If NumRegElim <> 1 Then
                Sql = "Seleccione una �nica empresa"
                
            Else
                EmpresaSt = RecuperaValor(CadenaDesdeOtroForm, 3)
                EmpresaSt = "ariconta" & EmpresaSt & "."
                
                CadenaDesdeOtroForm = DevuelveDesdeBD("numnivel", EmpresaSt & "empresa", "1", "1")
                If CadenaDesdeOtroForm = "" Then
                   Sql = "Error obteniendo datos empresa : " & EmpresaSt
                Else
                    CadenaDesdeOtroForm = "numdigi" & CadenaDesdeOtroForm
                    CadenaDesdeOtroForm = DevuelveDesdeBD(CadenaDesdeOtroForm, EmpresaSt & "empresa", "1", "1")
                    If CadenaDesdeOtroForm = "" Then
                        Sql = "Error obteniendo datos ultimo nivel: " & EmpresaSt
                    Else
                        If vEmpresa.DigitosUltimoNivel <> Val(CadenaDesdeOtroForm) Then
                            Sql = "Distintos digitos ultimo nivel"
                        End If
                    End If
                End If
            End If
            
            If Sql <> "" Then
                MsgBox Sql, vbExclamation
                Sql = ""
                Exit Sub
            End If
                
    
    End If
    AbrirSelCuentas2 0, EmpresaSt  '0. Cuentas normal
    
    If Sql <> "" Then
        Sql = RecuperaValor(Sql, 1)
        'Ha devuelto datos
        Me.Refresh
        DoEvents
        Screen.MousePointer = vbHourglass
        
            
        If Index = 0 Or Index = 2 Then
            PonerDatosDeOtraCuenta EmpresaSt
            'no nos traemos ni fecha de baja ni cuenta de contrapartida
            Text1(23).Text = ""
            Text1(27).Text = ""
        Else
            Me.txtRegularizacion.Text = Sql
        End If
        
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdRegresar_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub




Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        PrimeraVez = False
    End If
End Sub

Private Sub Form_Load()

    Me.Icon = frmPpal.Icon


    PrimeraVez = True

    SSTab1.Tab = 0
    Me.SSTab1.TabVisible(1) = vEmpresa.TieneTesoreria
    Text1(0).Enabled = True
    Text1(0).MaxLength = vEmpresa.DigitosUltimoNivel
    EnablarText (vModo <> 0)
    cmdCopiarDatos(0).Visible = vModo = 1
    cmdCopiarDatos(1).Visible = vModo = 1 Or vModo = 2
    
    For I = 0 To Me.imgppal.Count - 1
        Me.imgppal(I).Visible = vModo > 0
    Next I
    
    FrGranEmpresa.Visible = False
    
    ' La Ayuda
    With Me.ToolbarMail
        .ImageList = frmPpal.imgListComun16
        .Buttons(1).Image = 27
    End With
    
    With Me.ToolbarAux
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With
    
    For I = 0 To Image1.Count - 1
        Image1(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    
    
    Select Case vModo
    Case 0
            Modo = 2
    
            'Vamos a ver los datos
            PonerCampos ""
            
            lblIndicador.Caption = "Ver cuenta"
            
            CargaGrid 2, True
            
    Case 1
            Modo = 3
    
            LimpiarCampos
            If CodCta <> "" Then Text1(0).Text = CodCta
            '347
            Check1.Value = 1
            Frame1.Visible = True
            Frame1.Enabled = False
            lblIndicador.Caption = "INSERTAR"
            
            Me.cmdCopiarDatos(2).Visible = HayMasDeUnaEmpresa
            
            CargaGrid 2, False
            
    Case 2
    
            Modo = 4
    
            Text1(0).Enabled = False
            Text1(1).Enabled = True
            PonerCampos ""
            lblIndicador.Caption = "MODIFICAR"
            
            CargaGrid 2, True
            
   Case 3
            Modo = 1
    
            LimpiarCampos
            Frame1.Visible = True
            lblIndicador.Caption = "BUSQUEDA"
    
            CargaGrid 2, False
            
            Dim anc As Single
            anc = DataGridAux(2).Top
            If DataGridAux(2).Row < 0 Then
                anc = anc + 250
            Else
                anc = anc + DataGridAux(2).RowTop(DataGridAux(2).Row) + 5
            End If

            LLamaLineas 2, Modo, anc
    
    End Select
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
    If vModo = 0 Or vModo = 2 Then
        If Text1(11).Text = "S" Then
            kCampo = vModo
            vModo = 2
            Text1_LostFocus 25
            Text1_LostFocus 26
            Text1_LostFocus 12
            Text1_LostFocus 30
            If Text1(27).Text <> "" Then
                Text2(27).Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(27).Text, "T")
            End If
            vModo = kCampo
            kCampo = 0
        End If
    End If

    
    If vModo = 1 Or vModo = 0 Or (vModo = 2 And (Text1(11).Text = "S" Or chkUltimo.Value = 1)) Then
        Me.Text1(12).Enabled = True
        Me.Text1(30).Enabled = True
        For I = 2 To 3
            Me.Image1(I).Enabled = True
            Me.Image1(I).Visible = True
        Next I
        
    Else
        Me.Text1(12).Enabled = False
        Me.Text1(30).Enabled = False
        For I = 2 To 3
            Me.Image1(I).Enabled = False
            Me.Image1(I).Visible = False
        Next I
    End If
    
    If vModo = 2 Then
        Text1(0).BackColor = &H80000018
    Else
        Text1(0).BackColor = &H80000005
    End If
    
    
    ' solo podemos poner una cuenta habitual si es del grupo 4 o 5
    Dim B As Boolean
    B = ((Modo = 3 Or Modo = 4) And (Mid(Text1(0).Text, 1, 1) = "4" Or Mid(Text1(0).Text, 1, 1) = "5")) And chkUltimo.Value = 1
    Text1(27).Enabled = B
    Me.Image1(6).Enabled = B
    Me.Image1(6).Visible = B
   
    
    ' copiar de otra contabilidad solo puede ser si estamos insertando
    cmdCopiarDatos(0).Visible = (Modo = 3)
    cmdCopiarDatos(0).Enabled = (Modo = 3)
    cmdCopiarDatos(2).Visible = (Modo = 3)
    cmdCopiarDatos(2).Enabled = (Modo = 3)
    
    
End Sub

Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    'Aqui va el especifico de cada form es
    '### a mano
    chkUltimo.Value = 0
End Sub

Private Sub PonerCampos(QueEmpresa As String)
Dim Rs As ADODB.Recordset
Dim mTag As CTag
Dim I  As Integer
Dim T As Object
Dim Valor

    Set Rs = New ADODB.Recordset
    Sql = "Select * from " & QueEmpresa & "cuentas where codmacta='" & CodCta & "'"
    Rs.Open Sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    If Rs.EOF Then
        LimpiarCampos
        lblIndicador.Caption = "MODIFICAR"
    Else
        Set mTag = New CTag
        
        For I = 0 To Text1.Count - 1
            Set T = Text1(I)
            mTag.Cargar T
            If mTag.Cargado Then
                'Columna en la BD
                Sql = mTag.Columna
                If mTag.Vacio = "S" Then
                    Valor = DBLet(Rs.Fields(Sql))
                Else
                    Valor = Rs.Fields(Sql)
                End If
                If mTag.Formato <> "" Then Valor = Format(Valor, mTag.Formato)
                
                Text1(I).Text = Valor
            Else
                Text1(I).Text = ""
            End If
        Next I
        varBloqCta = ""
        If Rs.Fields!apudirec = "S" Then
            chkUltimo.Value = 1
            Text1(11).Text = "S"
            Me.Frame1.Visible = True
            varBloqCta = Text1(23).Text

            Else
            chkUltimo.Value = 0
            Frame1.Visible = False
            Text1(24).Text = Text1(10).Text
            Text1(11).Text = "N"
        End If
        Check1.Value = Rs!model347
        Check2.Value = Check1.Value
        Check2.Enabled = (vModo = 2)
        
        Check2.Visible = (Len(Text1(0).Text) = 3)
        lbl347.Visible = (Len(Text1(0).Text) = 3)
        
        Check3.Value = Rs!esctamultiple
        
        
        PonerFrameGranEmpresa
        
        If vModo = 2 And chkUltimo.Value = 1 Then
        End If
        Set mTag = Nothing


        Text1(13).Text = ""
        Text1(14).Text = ""
        Text1(15).Text = ""
        Text1(16).Text = ""
        Text1(33).Text = ""
        Text1(34).Text = ""
        
        Text1(13).ToolTipText = ""
        Text1(14).ToolTipText = ""
        Text1(15).ToolTipText = ""
        Text1(16).ToolTipText = ""
        Text1(33).ToolTipText = ""
        Text1(34).ToolTipText = ""

        If Text1(29).Text <> "" Then
            Text1(13).Text = Mid(Text1(29).Text, 1, 4)
            Text1(14).Text = Mid(Text1(29).Text, 5, 4)
            Text1(15).Text = Mid(Text1(29).Text, 9, 4)
            Text1(16).Text = Mid(Text1(29).Text, 13, 4)
            Text1(33).Text = Mid(Text1(29).Text, 17, 4)
            Text1(34).Text = Mid(Text1(29).Text, 21, 4)
            
            Dim CCC As String
            CCC = Text1(13).Text & " " & Text1(14).Text & " " & Text1(15).Text & " " & Mid(Text1(16).Text, 1, 2) & " " & Mid(Text1(16).Text, 3, 2) & Text1(33).Text & Text1(34).Text
            
            Text1(13).ToolTipText = CCC
            Text1(14).ToolTipText = CCC
            Text1(15).ToolTipText = CCC
            Text1(16).ToolTipText = CCC
            Text1(33).ToolTipText = CCC
            Text1(34).ToolTipText = CCC
            
        End If

    If Text1(26).Text <> "" And Len(Text1(0).Text) = 3 Then
        txtRegularizacion.Text = Text1(26).Text
    End If
    



    End If
End Sub

Private Sub PonerFrameGranEmpresa()
Dim B As Boolean
    
    B = False
    If vParam.GranEmpresa Then
        'y Si len 3 y cta 8 y 9
        If Len(Text1(0).Text) = 3 Then
            '8 y 9
            If Val(Mid(Text1(0), 1, 1)) >= 8 Then
                B = True
                'cuentaba en cuentas 7 y 8 a 3 digitos quiere decir DONDE regularizara
                txtRegularizacion.Text = Text1(16).Text
            End If
        End If
    End If
    Me.FrGranEmpresa.Visible = B
End Sub

Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(26).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    imgppal(0).Tag = vFecha
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(27).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(27).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmFPag_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(25).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
        Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmIVA_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(30).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(3).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmPais_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(12).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub

Private Sub Image1_Click(Index As Integer)

    Select Case Index
        Case 0 'formas de pago
            Set frmFPag = New frmBasico2
            AyudaFPago frmFPag
            Set frmFPag = Nothing

        Case 1 ' bancos
            Set frmBan = New frmBasico2
            AyudaBanco frmBan
            Set frmBan = Nothing

        Case 2 ' pais
            Set frmPais = New frmBasico2
            AyudaPais frmPais
            Set frmPais = Nothing
            
        Case 3 ' iva
            Set frmIVA = New frmBasico2
            AyudaTiposIva frmIVA
            Set frmIVA = Nothing
            
        Case 4 ' observaciones
            Indice = 10
            
            Set frmZ = New frmZoom
            frmZ.pValor = Text1(Indice).Text
            frmZ.pModo = Modo
            frmZ.Caption = "Observaciones Cuentas"
            frmZ.Show vbModal
            Set frmZ = Nothing
        
        Case 5 ' observaciones de tesoreria
            Indice = 22
            
            Set frmZ = New frmZoom
            frmZ.pValor = Text1(Indice).Text
            frmZ.pModo = Modo
            frmZ.Show vbModal
            Set frmZ = Nothing
        
        Case 6 ' cuenta habitual
            Set frmCtas = New frmColCtas
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.Show vbModal
            Set frmCtas = Nothing
        
            PonleFoco Text1(27)
        
        
    End Select

End Sub

Private Sub imgppal_Click(Index As Integer)
Dim Ix As Integer
    imgppal(0).Tag = ""
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    Select Case Index
    Case 0
        Ix = 18
    Case 1
        Ix = 20
    Case 3
        Ix = 28
    Case 4
        Ix = 32
    Case Else
        Ix = 23
    End Select
    
    If Text1(Ix).Text <> "" Then frmC.Fecha = CDate(Text1(Ix).Text)
    frmC.Show vbModal
    
    If imgppal(0).Tag <> "" Then Text1(Ix).Text = Format(imgppal(0).Tag, "dd/mm/yyyy")
        
    
End Sub

Private Sub imgWeb_Click(Index As Integer)
    LanzaVisorMimeDocumento Me.hWnd, Text1(9)
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    
    If vModo = 3 Then
        Text1(kCampo).BackColor = vbWhite
        Text1(Index).BackColor = vbLightBlue
        Else
            If Index <> 10 And Index <> 22 Then PonFoco Text1(Index)
    End If
    kCampo = Index
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 10 And Index <> 22 And Index <> 24 Then
        If KeyAscii = teclaBuscar Then
            Select Case Index
                Case 25: KEYImage KeyAscii, 0
                Case 26: KEYImage KeyAscii, 1
                Case 18: KEYFecha KeyAscii, 0
                Case 20: KEYFecha KeyAscii, 1
                Case 23: KEYFecha KeyAscii, 2
                Case 28: KEYFecha KeyAscii, 3
                Case 12: KEYImage KeyAscii, 2
                Case 30: KEYImage KeyAscii, 3
                Case 27: KEYImage KeyAscii, 6
            End Select
        Else
            KEYpress KeyAscii
        End If
    Else
        If (Index = 10 And Text1(10).Text = "") Or (Index = 22 And Text1(22).Text = "") Or (Index = 24 And Text1(24).Text = "") Then KEYpress KeyAscii
    End If
End Sub

Private Sub KEYImage(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    Image1_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgppal_Click (Indice)
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
    Dim I As Integer
    Dim Sql2 As String, Sql3 As String
    Dim mTag As CTag
    Dim Im As Currency
    
    If vModo = 3 Then
        If Text1(13).Text <> "" Then Text1(13).Text = UCase(Text1(13).Text)
        If Text1(14).Text <> "" Then Text1(14).Text = Format(Text1(14).Text, "0000")
        If Text1(15).Text <> "" Then Text1(15).Text = Format(Text1(15).Text, "0000")
        If Text1(16).Text <> "" Then Text1(16).Text = Format(Text1(16).Text, "0000")
        If Text1(33).Text <> "" Then Text1(16).Text = Format(Text1(33).Text, "0000")
        If Text1(34).Text <> "" Then Text1(16).Text = Format(Text1(34).Text, "0000")
    End If
    
    
    If vModo = 3 Or vModo = 0 Then Exit Sub 'Busqueda avanzada o ver solo
    
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbLightBlue Then
        Text1(Index).BackColor = vbWhite '&H80000018
    End If
    If Text1(Index).Text = "" Then
        If Index = 0 Then
            Frame1.Visible = True
            chkUltimo.Value = 0
        ElseIf Index = 25 Then
            Text2(0).Text = ""
        ElseIf Index = 26 Then
            Text2(1).Text = ""
        ElseIf Index = 12 Then
            Text2(2).Text = ""
        ElseIf Index = 30 Then
            Text2(3).Text = ""
        ElseIf Index = 27 Then
            Text2(27).Text = ""
        End If
        Exit Sub
    End If
    If Index <> 10 And Index <> 24 And Index <> 8 Then Text1(Index).Text = UCase(Text1(Index).Text)
    'Si queremos hacer algo ..
    Select Case Index
        Case 0
            PierdeFocoCodigoCuenta
        Case 1
            If vModo = 1 Then
                If Text1(2).Text = "" Then Text1(2).Text = Text1(1).Text
                If Text1(12).Text = "" Then Text1(12).Text = "ESPA�A"
            End If
        '....
        Case 13 To 16, 33, 34
            If vModo = 2 Then
                
                If Index <> 13 Then
                    If Not EsNumerico(Text1(Index).Text) Then
                        PonFoco Text1(Index)
                        Exit Sub
                    Else
                        Text1(Index).Text = Format(Text1(Index).Text, "0000")
                    End If
                
                    If Text1(14).Text <> "" And Text1(15).Text <> "" And Text1(16).Text <> "" And Text1(33).Text <> "" And Text1(34).Text <> "" Then
                        ' comprobamos si es correcto
                        Sql = Format(Text1(14).Text, "0000") & Format(Text1(15).Text, "0000") & Format(Text1(16).Text, "0000") & Format(Text1(33).Text, "0000") & Format(Text1(34).Text, "0000")
                        
                    End If
                Else
                
                    If Mid(Text1(Index).Text, 1, 2) = "ES" Then
                        If Not IBAN_Correcto(Me.Text1(Index).Text) Then Text1(Index).Text = ""
                
                    End If
                End If
                
                If Text1(13).Text <> "" And Text1(14).Text <> "" And Text1(15).Text <> "" And Text1(16).Text <> "" And Text1(33).Text <> "" And Text1(34).Text <> "" Then
                    
                    Sql = Format(Text1(14).Text, "0000") & Format(Text1(15).Text, "0000") & Format(Text1(16).Text, "0000") & Format(Text1(33).Text, "0000") & Format(Text1(34).Text, "0000")
            
                    Sql2 = CStr(Mid(Text1(13).Text, 1, 2))
                    If DevuelveIBAN2(CStr(Sql2), Sql, Sql) Then
                        If Mid(Text1(13).Text, 3, 2) <> Sql Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & Sql2 & Sql & "]", vbExclamation
                        End If
                    End If
                
                End If
                
                Text1(29).Text = Text1(13).Text & Format(ComprobarCero(Text1(14).Text), "0000") & Format(ComprobarCero(Text1(15).Text), "0000") & Format(ComprobarCero(Text1(16).Text), "0000") & Format(ComprobarCero(Text1(33).Text), "0000") & Format(Text1(34).Text, "0000")
            
            End If
                
        
        Case 18, 20, 23, 28, 32
            If Not EsFechaOK(Text1(Index)) Then
                MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
            End If
        
        Case 19, 21
            If Not CadenaCurrency(Text1(Index).Text, Im) Then
                MsgBox "Importe incorrecto: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
            Else
                Text1(Index).Text = Format(Im, FormatoImporte)
            End If
        
        Case 25
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(0).Text = PonerNombreDeCod(Text1(Index), "formapago", "nomforpa", "codforpa", "N")
                If Text2(0).Text = "" Then
                    MsgBox "No existe la Forma de Pago. Reintroduzca.", vbExclamation
                    PonFoco Text1(Index)
                End If
            Else
                Text2(0).Text = ""
            End If
        
        Case 26
            Sql = Text1(26).Text
            If CuentaCorrectaUltimoNivel(Sql, Sql2) Then
                Sql = DevuelveDesdeBD("codmacta", "bancos", "codmacta", Sql, "T")
                If Sql = "" Then
                    MsgBox "La cuenta NO pertenece a ning�na cta. bancaria", vbExclamation
                    Sql2 = ""
                    
                Else
                    'CORRECTO
                End If
            Else
                Sql = ""
                MsgBox Sql2, vbExclamation
                Sql2 = ""
            End If
            Text1(26).Text = Sql
            Text2(1).Text = Sql2
            If Sql = "" Then PonleFoco Text1(26)
            
        Case 29
            'IBAN no se ve
        
        Case 12 ' pais
            If Text1(Index).Text <> "" Then
                Text2(2).Text = PonerNombreDeCod(Text1(Index), "paises", "nompais", "codpais", "T")
                If Text2(2) = "" Then
                    MsgBox "No existe el Pa�s. Reintroduzca.", vbExclamation
                    PonFoco Text1(Index)
                End If
            Else
                Text2(2).Text = ""
            End If
            
        Case 30 ' tipo de iva
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(3).Text = PonerNombreDeCod(Text1(Index), "tiposiva", "nombriva", "codigiva", "N")
                If Text2(3) = "" Then
                    MsgBox "No existe el Tipo de Iva. Reintroduzca.", vbExclamation
                    PonFoco Text1(Index)
                End If
            Else
                Text2(3).Text = ""
            End If
            
        Case 27
            If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then
                Exit Sub
            End If
            
            Sql = Text1(27).Text
            If CuentaCorrectaUltimoNivel(Sql, Sql2) Then
                If EstaLaCuentaBloqueada(Text1(27).Text, Now) Then
                    MsgBox "Cuenta de contrapartida bloqueada, elim�nela o modif�quela.", vbExclamation
'                    SQL2 = ""
'                    SQL = ""
                    PonFoco Text1(27)
                Else
                    'CORRECTO
                End If
            Else
                Sql = ""
                MsgBox Sql2, vbExclamation
                Sql2 = ""
            End If
            Text1(27).Text = Sql
            Text2(27).Text = Sql2
            If Sql = "" Then PonleFoco Text1(27)
            
    End Select
    '---
End Sub

Private Function DatosOkLin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim B As Boolean
Dim cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLin = False

    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    If B And ModoLineas = 1 Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cConta, "departamentos", "dpto", "codmacta", txtaux3(0).Text, "T", , "dpto", txtaux3(1).Text, "N")
        If Sql <> "" Then
            MsgBox "El c�digo de departamento ya existe. Reintroduzca.", vbExclamation
            B = False
            PonFoco txtaux3(1)
        End If
    End If
    
    DatosOkLin = B

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation

End Function

Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Nivel As Integer
Dim RC As Byte
Dim RC2 As String
    
    
    DatosOK = False
    
    Text1(1).Text = UCase(Text1(1).Text)
    Text1(2).Text = UCase(Text1(2).Text)
    
       
    'Asignamos el campo apunte directo
    If chkUltimo.Value = 0 Then
        Text1(11).Text = "N"
    Else
        Text1(11).Text = "S"
    End If
    
    B = True
    If Len(Text1(0).Text) = vEmpresa.DigitosUltimoNivel Then
        'Digitos ultimo nivel
        If chkUltimo.Value = 0 Then
            MsgBox "La longitud de la cuenta es de ultimo nivel y no esta marcado", vbExclamation
            B = False
        End If
    Else
        'No tiene longitud de ultimo nivel
        If chkUltimo.Value = 1 Then
            MsgBox "No  es cuenta de ultimo nivel pero esta marcado", vbExclamation
            B = False
        End If
        
    End If
    If Not B Then Exit Function
    
    
    
    If Len(Text1(0).Text) < vEmpresa.DigitosUltimoNivel Then
        Check1.Value = 0
        Check3.Value = 0
        '--------------------------------
        'Si es nivel 3 entonces guardamos la oferta
        If Len(Text1(0).Text) = 3 Then
            Check1.Value = Check2.Value
            'Es gran empresa y digitos 8 9
            If Me.FrGranEmpresa.Visible Then
            
                If Mid(txtRegularizacion.Text, 1, 1) <> "1" Then
                    MsgBox "La regularizacion ser� contra las cuentas del grupo 1", vbExclamation
                    Exit Function
                End If
            
                'Compruebo que la cuenta existe
                Sql = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", txtRegularizacion.Text, "T")
                If Sql = "" Then
                    MsgBox "La cuenta " & txtRegularizacion.Text & " NO existe", vbExclamation
                    PonFoco txtRegularizacion
                    Exit Function
                End If
                Text1(16).Text = txtRegularizacion.Text
            End If
        End If
        
        'Si ha puesto observaciones las guardo
        Text1(10).Text = Text1(24).Text
    Else
        'Si estamos modificando o a�adiendo, el pais(text1(12)  cogera el valor que tenga el combo
'        Text1(12).Text = cboPais.Text
    End If
    
    
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    
    If Not IsNumeric(Text1(0).Text) Then
        MsgBox "Campo cuenta debe ser num�rico", vbExclamation
        Exit Function
    End If
    
    
    'Comprobamos de que nivel es la cuenta
    Nivel = NivelCuenta(Text1(0).Text)
    If Nivel < 1 Then
        MsgBox "El n�mero de d�gitos no pertenece a ning�n nivel contable", vbExclamation
        Exit Function
    End If
    
    
    If Text1(27).Text <> "" Then
        If EstaLaCuentaBloqueada(Text1(27).Text, Now) Then
            MsgBox "Cuenta de contrapartida bloqueada, elim�nela o modif�quela.", vbExclamation
            DatosOK = False
            PonFoco Text1(27)
            Exit Function
        End If
    End If
    
    
    
    'NIF
    If Text1(7).Text <> "" Then
        'Ha escrito el NIF
        If Text1(12).Text = "ES" Then
            If Not Comprobar_NIF(Text1(7).Text) Then
                If MsgBox("NIF incorrecto. �Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Function
            End If
        End If
        'Comprobacion NIFs
        'Comprobaremos si el NIF existe en cualquier otra contabilidad
        'comprobando que tenga permisos para ello
        ComprobarNifTodasContas
    End If
    
    
    
    
    
    
    If Nivel > 1 Then
    
    
        B = ExistenSubcuentas(Text1(0).Text, Nivel - 1)
        If Not B Then
            RC = MsgBox("No existen subcuentas inferiores para la cuenta : " & Text1(0).Text & vbCrLf & "Desea crealas ?", vbQuestion + vbYesNoCancel)
            If RC = vbYes Then
                'Hay que crear subcuentas
                B = CreaSubcuentas(Text1(0).Text, Nivel - 1, Text1(1).Text)
                If Not B Then Exit Function
            Else
                Exit Function
            End If
        End If
        
        
        
        
        
        
    End If
    
    
    'Compruebo cuenta bancaria
    
    If Text1(11).Text = "S" Then
       If Text1(13).Text <> "" Then
             Sql = Text1(14).Text & Text1(15).Text & Text1(16).Text & Text1(33).Text & Text1(34).Text
             If Len(Sql) <> 20 Then
                 MsgBox "Longitud cuenta bancaria incorrecta", vbExclamation
                 Exit Function
             End If
    
            'Compruebo EL IBAN
            'Meto el CC
            RC2 = Sql
            Sql = ""
            If Me.Text1(13).Text <> "" Then Sql = Mid(Text1(13).Text, 1, 2)

            If DevuelveIBAN2(Sql, RC2, RC2) Then
                If Me.Text1(13).Text = "" Then
                    If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(13).Text = RC2
                Else
                    If Mid(Text1(13).Text, 3, 2) <> RC2 Then
                        RC2 = "Calculado : " & Sql & RC2
                        RC2 = "Introducido: " & Me.Text1(13).Text & vbCrLf & RC2 & vbCrLf
                        RC2 = "Error en codigo IBAN" & vbCrLf & RC2 & "Continuar?"
                        If MsgBox(RC2, vbQuestion + vbYesNo) = vbNo Then Exit Function
                    End If
                End If
            End If
        Else
            Text1(29).Text = ""
        
        End If
    End If
    
    If txtRegularizacion.Text <> "" And Len(Text1(0).Text) = 3 Then
        Text1(26).Text = txtRegularizacion.Text
    End If
    
    
    
    DatosOK = True
End Function




Private Sub PierdeFocoCodigoCuenta()
Dim B As Boolean
If vModo = 3 Then Exit Sub  'B�squeda


If vModo = 1 Then Text1(0).Text = Trim(Text1(0).Text)

'Si no compruebo que es un campo numerico
If Not IsNumeric(Text1(0).Text) Then
    MsgBox "El c�digo de cuenta es un campo num�rico", vbExclamation
    Exit Sub
End If

'Vemos si a puesto el punto para rellenar
Text1(0).Text = RellenaCodigoCuenta(Text1(0).Text)

If Len(Text1(0).Text) > vEmpresa.DigitosUltimoNivel Then
    MsgBox "El n�mero m�ximo de d�gitos para las cuentas es de " & vEmpresa.DigitosUltimoNivel & _
        vbCrLf & "La cuenta que ha puesto tiene " & Len(Text1(0).Text), vbExclamation
    Exit Sub
End If

'Comprobamos que ya existe la cuenta, solo en nueva
If vModo = 1 Then
    Sql = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(0).Text, "T")
    If Sql <> "" Then
        MsgBox "La cuenta: " & Text1(0).Text & " ya esta asignada." & vbCrLf & "      .-" & Sql, vbExclamation
        Text1(0).SetFocus
        Exit Sub
    End If
End If
'Ponemos , si es de ultimo nivel habilitados los campos

B = EsCuentaUltimoNivel(Text1(0).Text)
Frame1.Visible = B
Frame1.Enabled = True
chkUltimo.Value = Abs(CInt(B))
If Not B Then
    'Si no es ultimo nivel
    Check2.Enabled = Len(Text1(0).Text) = 3
    PonerFrameGranEmpresa
Else
    'Ultimo nivel
    If vModo = 1 Then
        'A�adir cuenta
        Sql = DevuelveDesdeBD("model347", "cuentas", "codmacta", Mid(Text1(0).Text, 1, 3), "T")
        If Sql = "1" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    End If
End If

End Sub



Private Sub EnablarText(Si As Boolean)
Dim T As TextBox
    For Each T In Text1
        T.Locked = Not Si
    Next
    Image1(0).Enabled = Si
    Image1(1).Enabled = Si
    Check1.Enabled = Si
    Check3.Enabled = Si
    Me.Check2.Enabled = Si
    Me.txtRegularizacion.Enabled = Si
    Me.chkUltimo.Enabled = Si
    'Solo administradores puden bloquear cuenta
    Text1(23).Enabled = vUsu.Nivel <= 1
    imgppal(2).Enabled = vUsu.Nivel <= 1
    
End Sub

Private Sub PonerDatosDeOtraCuenta(QueEmpresa_ As String)
Dim C As String
    C = Text1(0).Text
    Text1(0).Visible = False
    CodCta = Sql
    PonerCampos QueEmpresa_
    lblIndicador.Caption = "Insertar"
    If QueEmpresa_ = "" Then
        Text1(0).Text = C
    Else
        If C <> "" Then Text1(0).Text = C
    End If
    Text1(0).Visible = True
    CodCta = ""
End Sub



Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea 2
        Case 2
            BotonModificarLinea 2
        Case 3
            BotonEliminarLinea 2
        Case Else
    End Select

End Sub

Private Sub ToolbarMail_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim dirMail As String
   
    Select Case Button.Index
        Case 1
            Screen.MousePointer = vbHourglass
            
            dirMail = Text1(8).Text
            
            If LanzaMailGnral(dirMail) Then espera 2
            Screen.MousePointer = vbDefault
    End Select
End Sub

Private Sub txtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtRegularizacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtRegularizacion_LostFocus()
    If vModo = 3 Or vModo = 0 Then Exit Sub 'Busqueda avanzada o ver solo
    
    
    If txtRegularizacion.Text = "" Then Exit Sub
    
    'Si no compruebo que es un campo numerico
    If Not IsNumeric(txtRegularizacion.Text) Then
        MsgBox "El c�digo de cuenta es un campo num�rico", vbExclamation
        txtRegularizacion.Text = ""
        PonFoco txtRegularizacion
        Exit Sub
    End If
    
    'Vemos si a puesto el punto para rellenar
    txtRegularizacion.Text = RellenaCodigoCuenta(txtRegularizacion.Text)
    
    
    
    'Solo son validad cuentas del grupo 1
    If Mid(txtRegularizacion.Text, 1, 1) <> "1" Then
        MsgBox "La regularizacion ser� contra las cuentas del grupo 1", vbExclamation
        txtRegularizacion.Text = ""
        PonFoco txtRegularizacion
        Exit Sub
    End If
    
    
    
    If Len(Text1(0).Text) > vEmpresa.DigitosUltimoNivel Then
        MsgBox "El n�mero m�ximo de d�gitos para las cuentas es de " & vEmpresa.DigitosUltimoNivel & _
            vbCrLf & "La cuenta que ha puesto tiene " & Len(Text1(0).Text), vbExclamation
        txtRegularizacion.Text = ""
        PonFoco txtRegularizacion
        Exit Sub
    End If
    
    
    
    
    
End Sub





Private Sub ComprobarNifTodasContas()
    Set miRsAux = New ADODB.Recordset
    DoEvents
    cargaempresas
    lblIndicador.Caption = "Modificar"
    Set miRsAux = Nothing
End Sub


Private Sub cargaempresas()
Dim Mensa As String
Dim Prohibidas As Boolean
Dim C As String
On Error GoTo Ecargaempresas

    
    
    Sql = "Select count(*) from Usuarios.usuarioempresasariconta WHERE codusu = " & (vUsu.Codigo Mod 1000)
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Prohibidas = False
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then Prohibidas = True
    End If
    miRsAux.Close

    
    Sql = "Select * from Usuarios.Empresasariconta where conta like 'ariconta%' order by codempre"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Sql = ""
    While Not miRsAux.EOF
        Sql = Sql & miRsAux!codempre & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Mensa = ""
    Do
        kCampo = InStr(1, Sql, "|")
        If kCampo > 0 Then
                C = Mid(Sql, 1, kCampo - 1)
                Sql = Mid(Sql, kCampo + 1)
                
                NumRegElim = Val(C)
                C = "conta" & C
                lblIndicador.Caption = "Comprobando NIF: " & C
                lblIndicador.Refresh
                C = "Select codmacta,nommacta FROM " & C & ".cuentas where apudirec='S'"
                If NumRegElim = vEmpresa.codempre Then
                    'Es esta empresa.
                    'Si es modificar a�adire el codmacta <> de esta cuenta
                    If vModo = 2 Then C = C & " AND codmacta <> '" & Text1(0).Text & "'"
                End If
                C = C & " AND nifdatos ='" & Text1(7).Text & "'"
                miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                C = "Conta: " & NumRegElim & vbCrLf
                kCampo = 0
                While Not miRsAux.EOF
                    kCampo = 1
                    C = C & "    " & miRsAux!codmacta & " - " & miRsAux!Nommacta & vbCrLf
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                If kCampo > 0 Then
                    Mensa = Mensa & C & vbCrLf
                Else
                    kCampo = 1
                End If
         End If
    Loop Until kCampo = 0
    
    
    If Mensa <> "" Then
        If Prohibidas Then
            Mensa = "YA existe el NIF en la contabilidad"
        Else
            Mensa = "El NIF aparece en la contabilidad." & vbCrLf & vbCrLf & Mensa
        End If
        MsgBox Mensa, vbExclamation
    End If
Ecargaempresas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos empresas"
   
End Sub



Private Function HayMasDeUnaEmpresa() As Boolean

    HayMasDeUnaEmpresa = False
    Sql = " not codempre in (select codempre from usuarios.usuarioempresasariconta where codusu=" & vUsu.Codigo Mod 1000 & ") and 1"
    Sql = DevuelveDesdeBD("count(*)", "usuarios.empresasariconta", Sql, "1", "N")
    If Sql <> "" Then
        If Val(Sql) > 1 Then HayMasDeUnaEmpresa = True
    End If

End Function

Private Sub CargaGrid(Index As Integer, Enlaza As Boolean)
Dim B As Boolean
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, Enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez


    Select Case Index
        Case 2 'pozos
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtAux3(1)|T|C�digo|1000|;" '
            tots = tots & "S|txtAux3(2)|T|Descripci�n|8200|;"

            arregla tots, DataGridAux(Index), Me

            DataGridAux(Index).Columns(2).Alignment = dbgLeft


    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic

    ' **** si n'hi han ll�nies en grids i camps fora d'estos ****
    If Not AdoAux(Index).Recordset.EOF Then
    
    Else
        LimpiarCamposFrame Index
    End If
    ' **********************************************************

ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub



Private Function MontaSQLCarga(Index As Integer, Enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informaci� proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enla�a en el data1
'           -> Si no el carreguem sense enlla�ar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
       Case 2 ' pozos
            tabla = "departamentos"
            Sql = "SELECT departamentos.codmacta, departamentos.dpto, departamentos.descripcion "
            Sql = Sql & " FROM " & tabla
            If Enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE codmacta = '-1'"
            End If
            Sql = Sql & " ORDER BY " & tabla & ".dpto "
            
            
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function

Private Sub LimpiarCamposFrame(Index As Integer)
Dim I As Integer
    On Error Resume Next

    Select Case Index
        Case 2 'departamentos
            For I = 0 To txtaux3.Count - 1
                txtaux3(I).Text = ""
            Next I
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " codmacta='" & Trim(Text1(0).Text) & "'"
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim Eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Ll�nia
    
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 2 'departamentos
            Sql = "�Seguro que desea eliminar el registro?"
            Sql = Sql & vbCrLf & "Departamento: " & AdoAux(Index).Recordset!Dpto
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM departamentos"
                Sql = Sql & vWhere & " AND dpto= " & DBLet(AdoAux(Index).Recordset!Dpto, "N")
                
            End If
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
'        TerminaBloquear
        Conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
        End If
'        ' ***************************************
'        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
        Modo = 4
    End If
    
    ModoLineas = 0
'    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
    Select Case Index
        Case 2 'departamentos
            Sql = "select count(*) from cobros where codmacta = '" & Trim(AdoAux(2).Recordset!codmacta) & "' and departamento =" & AdoAux(2).Recordset!Dpto
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "Este departamento se encuentra en el mantenimiento de cobros. Revise. ", vbInformation   '& vbCrLf & "� Desea eliminarlo de todas formas ?" & vbCrLf & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                Exit Function
            End If
    End Select
    ' ****************************************************
    
    SepuedeBorrar = True
End Function




Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim I As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Ll�nia
    
    PonerModo 5, Index

    ' *** posar el nom del les distintes taules de ll�nies ***
    Select Case Index
        Case 2: vTabla = "departamentos"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 2
            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 250
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            For I = 0 To txtaux3.Count - 1
                txtaux3(I).Text = ""
            Next I
            
            txtaux3(0).Text = Text1(0).Text 'cuenta
            txtaux3(1).Text = NumF 'departamento
            PonFoco txtaux3(1)
         
    End Select
End Sub


Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim J As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar ll�nia
       
    PonerModo 5, Index
  
    Select Case Index
        Case 0, 1, 2 ' *** pose els index de ll�nies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, I
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
        
    End Select
    
    Select Case Index
        Case 2 'departamentos
            For I = 0 To 2
                txtaux3(I).Text = DataGridAux(Index).Columns(I).Text
            Next I
        
            CargarValoresAnteriores Me, 2, "FrameAux2"
        
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 2 ' departamentos
            PonFoco txtaux3(2)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim I As Integer, NumReg As Byte
Dim B As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    For I = 0 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I
       
    '---------------------------------------------
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.Visible = B
    cmdAceptar.Visible = B
       

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 2, False
    End If
    
    B = (Modo = 4) Or (Modo = 2)
    DataGridAux(2).Enabled = B
    
    'departamentos
    B = (Modo = 5 Or Modo = 1)
    For I = 1 To 2
        txtaux3(I).Enabled = B
    Next I
    B = (Modo = 5 Or Modo = 1) And ModoLineas = 1
    txtaux3(1).Enabled = B
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************
       
    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Ll�nies
    Select Case Index
        Case 2 ' departamentos
            For jj = 1 To 2
                txtaux3(jj).Visible = B
                txtaux3(jj).Top = alto
            Next jj
    End Select
End Sub


'**************************************************************************
'**************************************************************************
'**************************************************************************

Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        
        Me.ToolbarAux.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 4 Or Modo = 2) And vParam.NroAriges = 0
        Me.ToolbarAux.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 4 Or Modo = 2) And vParam.NroAriges = 0
        Me.ToolbarAux.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 4 Or Modo = 2) And vParam.NroAriges = 0
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub

Private Sub TxtAux3_GotFocus(Index As Integer)
    If Not txtaux3(Index).MultiLine Then ConseguirFoco txtaux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 And KeyAscii = 13 Then
        cmdAceptar.SetFocus
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim CADENA As String
    
    If Not PerderFocoGnral(txtaux3(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
    Select Case Index
        Case 1 ' departamento
            PonerFormatoEntero txtaux3(Index)
            
    End Select
    
    ' ******************************************************************************
End Sub
