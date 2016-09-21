VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInmov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "frmInmov.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   7080
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   7575
      Begin VB.Frame FrameTesor 
         Height          =   2175
         Left            =   240
         TabIndex        =   142
         Top             =   4080
         Width           =   7215
         Begin VB.TextBox Text8 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   3000
            TabIndex        =   149
            Top             =   1680
            Width           =   3975
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   54
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   1
            Left            =   240
            TabIndex        =   51
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   3000
            TabIndex        =   146
            Top             =   480
            Width           =   3975
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Index           =   0
            Left            =   1800
            TabIndex        =   52
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtDescta 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   3000
            TabIndex        =   143
            Top             =   1080
            Width           =   3975
         End
         Begin VB.TextBox txtcta 
            Height          =   285
            Index           =   3
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   53
            Tag             =   "Cuentas pérdidas y ganancias|T|S|0||parametros|ctaperga|||"
            Top             =   1080
            Width           =   1125
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Agente"
            Height          =   195
            Index           =   6
            Left            =   1800
            TabIndex        =   148
            Top             =   1440
            Width           =   510
         End
         Begin VB.Image imgTesoreria 
            Height          =   240
            Index           =   2
            Left            =   2400
            Picture         =   "frmInmov.frx":000C
            Top             =   1440
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   6
            Left            =   1200
            Picture         =   "frmInmov.frx":0A0E
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgTesoreria 
            Height          =   240
            Index           =   1
            Left            =   2760
            Picture         =   "frmInmov.frx":0A99
            Top             =   840
            Width           =   240
         End
         Begin VB.Image imgTesoreria 
            Height          =   240
            Index           =   0
            Left            =   2760
            Picture         =   "frmInmov.frx":149B
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Forma pago"
            Height          =   195
            Index           =   5
            Left            =   1800
            TabIndex        =   145
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Cta cobro"
            Height          =   195
            Index           =   4
            Left            =   1800
            TabIndex        =   144
            Top             =   840
            Width           =   690
         End
         Begin VB.Label Label13 
            Caption         =   "Fecha Pago"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   147
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   70
         Text            =   "Text7"
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   45
         Text            =   "Text4"
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5040
         TabIndex        =   55
         Top             =   6480
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   6240
         TabIndex        =   56
         Top             =   6480
         Width           =   1095
      End
      Begin VB.Frame FrameVenta 
         Height          =   2055
         Left            =   240
         TabIndex        =   61
         Top             =   2040
         Width           =   7215
         Begin VB.Frame Frame8 
            Caption         =   "Frame8"
            Height          =   615
            Left            =   240
            TabIndex        =   157
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
            Begin VB.TextBox txtDescta 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               Left            =   1440
               TabIndex        =   158
               Top             =   360
               Width           =   3975
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Base Fac."
               Height          =   195
               Index           =   3
               Left            =   600
               TabIndex        =   159
               Top             =   120
               Width           =   360
            End
            Begin VB.Image imgCta 
               Height          =   240
               Index           =   2
               Left            =   1080
               Picture         =   "frmInmov.frx":1E9D
               Top             =   120
               Width           =   240
            End
         End
         Begin VB.TextBox txtCodCCost 
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   50
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txtNomCcost 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   3000
            TabIndex        =   151
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox txtDescta 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   3000
            TabIndex        =   62
            Top             =   480
            Width           =   3975
         End
         Begin VB.TextBox txtcta 
            Height          =   285
            Index           =   1
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   49
            Tag             =   "Cuentas pérdidas y ganancias|T|S|0||parametros|ctaperga|||"
            Top             =   480
            Width           =   1125
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   240
            TabIndex        =   48
            Text            =   "Text5"
            Top             =   480
            Width           =   1215
         End
         Begin VB.Image imgCCost 
            Height          =   240
            Index           =   2
            Left            =   2760
            Picture         =   "frmInmov.frx":289F
            Top             =   1200
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Centro coste"
            Height          =   195
            Index           =   7
            Left            =   1800
            TabIndex        =   150
            Top             =   1200
            Width           =   900
         End
         Begin VB.Label Label13 
            Caption         =   "Importe"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   64
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "Cliente"
            Height          =   195
            Index           =   1
            Left            =   1800
            TabIndex        =   63
            Top             =   240
            Width           =   600
         End
         Begin VB.Image imgCta 
            Height          =   240
            Index           =   1
            Left            =   2400
            Picture         =   "frmInmov.frx":32A1
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.TextBox txtDescta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   3240
         TabIndex        =   59
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox txtcta 
         Height          =   285
         Index           =   0
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   47
         Tag             =   "Cuentas pérdidas y ganancias|T|S|0||parametros|ctaperga|||"
         Top             =   1680
         Width           =   1125
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   46
         Text            =   "Text4"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Baja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   57
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   44
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Image imgElto 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmInmov.frx":3CA3
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Elemento"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   69
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label14 
         Caption         =   "Cuenta perdidas / beneficios"
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   60
         Top             =   1440
         Width           =   2040
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   0
         Left            =   4200
         Picture         =   "frmInmov.frx":46A5
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmInmov.frx":50A7
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   58
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Venta / Baja de inmovilizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   2985
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3960
         TabIndex        =   34
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdSimula 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   33
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox txtfecha 
         Height          =   285
         Left            =   2640
         TabIndex        =   32
         Text            =   "Text4"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtConce 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtConce 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1200
         Width           =   855
      End
      Begin VB.Image imgConceInmo 
         Height          =   240
         Index           =   1
         Left            =   750
         Picture         =   "frmInmov.frx":5132
         Top             =   1700
         Width           =   240
      End
      Begin VB.Image imgConceInmo 
         Height          =   240
         Index           =   0
         Left            =   760
         Picture         =   "frmInmov.frx":5B34
         Top             =   1250
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Simulación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha amortizacion"
         Height          =   195
         Left            =   720
         TabIndex        =   31
         Top             =   3000
         Width           =   1395
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   2280
         Picture         =   "frmInmov.frx":6536
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   28
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Concepto Inmovilizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   2580
      End
   End
   Begin VB.Frame Frame0 
      Height          =   5760
      Left            =   0
      TabIndex        =   10
      Top             =   375
      Width           =   5415
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1590
         TabIndex        =   7
         Top             =   5040
         Width           =   375
      End
      Begin VB.TextBox txtIVA 
         Height          =   285
         Index           =   0
         Left            =   240
         MaxLength       =   8
         TabIndex        =   6
         Tag             =   "Cuentas pérdidas y ganancias|T|S|0||parametros|ctaperga|||"
         Top             =   4560
         Width           =   765
      End
      Begin VB.TextBox txtIVA 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1080
         TabIndex        =   65
         Top             =   4560
         Width           =   3975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   5
         Tag             =   "Concepto haber|N|S|0||||||"
         Text            =   "Text2"
         Top             =   3420
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   1
         Tag             =   "Ultima fecha amortizacion|F|S|||||dd/mm/yyyy||"
         Text            =   "Text2"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4200
         TabIndex        =   9
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2760
         TabIndex        =   16
         Text            =   "Text3"
         Top             =   3420
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   4
         Tag             =   "Concepto debe|N|S|0||||||"
         Text            =   "Text2"
         Top             =   2940
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2760
         TabIndex        =   15
         Text            =   "Text3"
         Top             =   2940
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   3
         Tag             =   "Nº diario|N|S|0||||||"
         Text            =   "Text2"
         Top             =   2460
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   2460
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   2040
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmInmov.frx":65C1
         Left            =   1920
         List            =   "frmInmov.frx":65D1
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Papel preimpreso"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   68
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Venta inmovilizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   240
         TabIndex        =   67
         Top             =   3960
         Width           =   2100
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00004000&
         X1              =   2160
         X2              =   5040
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Image imgiva 
         Height          =   240
         Left            =   600
         Picture         =   "frmInmov.frx":65FC
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label14 
         Caption         =   "IVA"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   66
         Top             =   4320
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmInmov.frx":6FFE
         Top             =   3405
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmInmov.frx":7A00
         Top             =   2925
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   960
         Picture         =   "frmInmov.frx":8402
         Top             =   2445
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   1680
         Picture         =   "frmInmov.frx":8E04
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Contabilizacion auto."
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   2160
         X2              =   5160
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         X1              =   2160
         X2              =   5160
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label5 
         Caption         =   "Datos contables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   1740
      End
      Begin VB.Label Label4 
         Caption         =   "Datos generales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label3 
         Caption         =   "Concepto haber"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   3435
         Width           =   1140
      End
      Begin VB.Label Label3 
         Caption         =   "Concepto debe"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   2955
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Diario"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   2475
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Ultima fecha amorti."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de amortización"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   660
         Width           =   1470
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrDeshacer 
      Height          =   2655
      Left            =   0
      TabIndex        =   152
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton cmdDeshaz 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   154
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdDeshaz 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   153
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label13 
         Height          =   1215
         Index           =   6
         Left            =   240
         TabIndex        =   156
         Top             =   720
         Width           =   5175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Deshacer última amortización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   23
         Left            =   120
         TabIndex        =   155
         Top             =   240
         Width           =   5205
      End
   End
   Begin VB.Frame Frame5 
      Height          =   5940
      Left            =   0
      TabIndex        =   88
      Top             =   0
      Width           =   5235
      Begin VB.TextBox Text7 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   140
         Text            =   "Text7"
         Top             =   3300
         Width           =   3135
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   2
         Left            =   1020
         TabIndex        =   92
         Text            =   "Text4"
         Top             =   3300
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   1
         Left            =   1020
         TabIndex        =   91
         Text            =   "Text4"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   137
         Text            =   "Text7"
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Frame Frame7 
         Height          =   1095
         Index           =   1
         Left            =   180
         TabIndex        =   135
         Top             =   4140
         Width           =   4815
         Begin VB.CheckBox chkEstadisticas 
            Caption         =   "Totalmente amortizado"
            Height          =   255
            Index           =   11
            Left            =   2580
            TabIndex        =   96
            Top             =   720
            Width           =   1995
         End
         Begin VB.CheckBox chkEstadisticas 
            Caption         =   "Baja"
            Height          =   255
            Index           =   10
            Left            =   2580
            TabIndex        =   94
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chkEstadisticas 
            Caption         =   "Vendido"
            Height          =   255
            Index           =   9
            Left            =   660
            TabIndex        =   95
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chkEstadisticas 
            Caption         =   "Activo"
            Height          =   255
            Index           =   8
            Left            =   660
            TabIndex        =   93
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label Label6 
            Caption         =   "Incluir elementos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   136
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Salto de página por elemento"
         Height          =   255
         Left            =   180
         TabIndex        =   97
         Top             =   5520
         Width           =   2475
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4020
         TabIndex        =   99
         Top             =   5460
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2820
         TabIndex        =   98
         Top             =   5460
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1860
         TabIndex        =   101
         Text            =   "Text1"
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox txtConce 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1020
         TabIndex        =   90
         Text            =   "Text1"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1860
         TabIndex        =   100
         Text            =   "Text1"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtConce 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1020
         TabIndex        =   89
         Text            =   "Text1"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgElto 
         Height          =   240
         Index           =   2
         Left            =   780
         Picture         =   "frmInmov.frx":8E8F
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   141
         Top             =   3360
         Width           =   480
      End
      Begin VB.Label Label6 
         Caption         =   "Elemento inmovilizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   9
         Left            =   180
         TabIndex        =   139
         Top             =   2520
         Width           =   2580
      End
      Begin VB.Label Label13 
         Caption         =   "Desde"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   138
         Top             =   2940
         Width           =   480
      End
      Begin VB.Image imgElto 
         Height          =   240
         Index           =   1
         Left            =   780
         Picture         =   "frmInmov.frx":9891
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgConceInmo 
         Height          =   240
         Index           =   5
         Left            =   780
         Picture         =   "frmInmov.frx":A293
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgConceInmo 
         Height          =   240
         Index           =   4
         Left            =   780
         Picture         =   "frmInmov.frx":AC95
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Ficha de elementos de inmovilizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   105
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label7 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   104
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Desde"
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   103
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Concepto Inmovilizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   4
         Left            =   180
         TabIndex        =   102
         Top             =   780
         Width           =   2580
      End
   End
   Begin VB.Frame Frame4 
      Height          =   4560
      Left            =   0
      TabIndex        =   71
      Top             =   0
      Width           =   5415
      Begin VB.Frame Frame50 
         Height          =   1095
         Left            =   360
         TabIndex        =   82
         Top             =   2400
         Width           =   4575
         Begin VB.CheckBox chkEstadisticas 
            Caption         =   "Totalmente amortizado"
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   87
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chkEstadisticas 
            Caption         =   "Baja"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   86
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chkEstadisticas 
            Caption         =   "Vendido"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   85
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chkEstadisticas 
            Caption         =   "Activo"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   84
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Incluir elementos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   83
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdListado1 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   81
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdListado1 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   80
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox txtConce 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   74
         Text            =   "Text1"
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txtConce 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   73
         Text            =   "Text1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Image imgConceInmo 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmInmov.frx":B697
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgConceInmo 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmInmov.frx":C099
         Top             =   1400
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Concepto Inmovilizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   79
         Top             =   1080
         Width           =   2580
      End
      Begin VB.Label Label7 
         Caption         =   "Desde"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   78
         Top             =   1380
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   77
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Estadisticas inmovilizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   76
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   0
      TabIndex        =   36
      Top             =   360
      Width           =   5295
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   375
         Left            =   360
         TabIndex        =   41
         Top             =   3720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdCalcula 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1560
         TabIndex        =   40
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtFecAmo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text4"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Cálculo y contabilización amortización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   720
         Width           =   4695
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   3000
         Picture         =   "frmInmov.frx":CA9B
         Top             =   2160
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha amortizacion"
         Height          =   195
         Left            =   1560
         TabIndex        =   38
         Top             =   2160
         Width           =   1395
      End
   End
   Begin VB.Frame Frame6 
      Height          =   6120
      Left            =   0
      TabIndex        =   106
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtNomCcost 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   134
         Text            =   "Text8"
         Top             =   3720
         Width           =   2535
      End
      Begin VB.TextBox txtCodCCost 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   113
         Text            =   "Text8"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox txtNomCcost 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   132
         Text            =   "Text8"
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox txtCodCCost 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   112
         Text            =   "Text8"
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txtfec 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   111
         Text            =   "Text8"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtfec 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   110
         Text            =   "Text8"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2040
         TabIndex        =   122
         Text            =   "Text1"
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox txtConce 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1080
         TabIndex        =   109
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2040
         TabIndex        =   121
         Text            =   "Text1"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtConce 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1080
         TabIndex        =   108
         Text            =   "Text1"
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton cmdListado2 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   114
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton cmdListado2 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   115
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Frame Frame7 
         Height          =   1095
         Index           =   0
         Left            =   240
         TabIndex        =   107
         Top             =   4320
         Width           =   4815
         Begin VB.CheckBox chkEstadisticas 
            Caption         =   "Totalmente amortizado"
            Height          =   255
            Index           =   7
            Left            =   2280
            TabIndex        =   119
            Top             =   720
            Width           =   2295
         End
         Begin VB.CheckBox chkEstadisticas 
            Caption         =   "Vendido"
            Height          =   255
            Index           =   5
            Left            =   2280
            TabIndex        =   118
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chkEstadisticas 
            Caption         =   "Baja"
            Height          =   255
            Index           =   6
            Left            =   840
            TabIndex        =   117
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chkEstadisticas 
            Caption         =   "Activo"
            Height          =   255
            Index           =   4
            Left            =   840
            TabIndex        =   116
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Incluir elementos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   120
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmInmov.frx":CB26
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Desde"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   133
         Top             =   3735
         Width           =   495
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmInmov.frx":D528
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Centro de coste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   7
         Left            =   240
         TabIndex        =   131
         Top             =   3120
         Width           =   2580
      End
      Begin VB.Label Label7 
         Caption         =   "Desde"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   130
         Top             =   3375
         Width           =   495
      End
      Begin VB.Image imgConceInmo 
         Height          =   240
         Index           =   7
         Left            =   840
         Picture         =   "frmInmov.frx":DF2A
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgConceInmo 
         Height          =   240
         Index           =   6
         Left            =   840
         Picture         =   "frmInmov.frx":E92C
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   5
         Left            =   3000
         Picture         =   "frmInmov.frx":F32E
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   9
         Left            =   2520
         TabIndex        =   129
         Top             =   2520
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   4
         Left            =   960
         Picture         =   "frmInmov.frx":F3B9
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Desde"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   128
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   6
         Left            =   240
         TabIndex        =   127
         Top             =   2280
         Width           =   2580
      End
      Begin VB.Label Label9 
         Caption         =   "Estadisticas inmovilizado entre fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   126
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label7 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   125
         Top             =   1620
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Desde"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   124
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Concepto Inmovilizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   5
         Left            =   240
         TabIndex        =   123
         Top             =   840
         Width           =   2580
      End
   End
End
Attribute VB_Name = "frmInmov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Public opcion As Byte
    '0.- Parametros
    '1.- Simular
    '2.- Cálculo amort.
    '3.- Venta/Baja inmovilizado
    '---------------------------
    'los siguiente utilizan el mismo frame, con opciones
    '4.- Listado estadisticas
    '5.- Ficha elementos
    '6.- Entre fechas


    '10.- Deshacer ultima amortizacion

Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmC As frmConceptos
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmCt As frmColCtas
Attribute frmCt.VB_VarHelpID = -1
Private WithEvents frmE As frmInmoElto
Attribute frmE.VB_VarHelpID = -1
Private WithEvents frmI As frmIVA
Attribute frmI.VB_VarHelpID = -1
Private WithEvents frmCI As frmInmoConceptos
Attribute frmCI.VB_VarHelpID = -1
Private WithEvents frmCC As frmCCoste
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmBa As frmCuentasBancarias
Attribute frmBa.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim Rs As Recordset
Dim Cad As String
Dim i As Byte
Dim B As Boolean
Dim Importe As Currency
'
'Desde parametros
Dim Contabiliza As Boolean
Dim UltAmor As Date
Dim DivMes As Integer
Dim ParametrosContabiliza As String
Dim Mc As Contadores

'Tipo de IVA
Dim TipoIva As String
Dim AUX2 As String


'Contador para las lineas de apuntes
Dim Cont As Integer

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdCalcula_Click()
    If MsgBox("Seguro que desea realizar la amortización a fecha: " & txtFecAmo.Text & " ?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    If txtFecAmo.Text = "" Then
        MsgBox "Fecha incorrecta", vbExclamation
        Exit Sub
    End If
    If Me.Tag <> "" Then
        If CDate(Me.txtFecAmo.Text) < CDate(Me.Tag) Then
            MsgBox "Fecha no puede ser menor que la ultima fecha de amortizacion: " & Me.Tag, vbExclamation
            Exit Sub
        End If
    End If
    i = FechaCorrecta2(CDate(txtFecAmo.Text))
    If i > 1 Then
        If i = 2 Then
            MsgBox varTxtFec, vbExclamation
        Else
            If i = 2 Then
                MsgBox "Fecha de amortización pertence a un ejercicio cerrado.", vbExclamation
            Else
                MsgBox "Fecha amortización pertenece a un ejercicio todavía no abierto", vbExclamation
            End If
        End If
        Exit Sub
    End If
    'Leemos los parametros
    If ObtenerparametrosAmortizacion(DivMes, UltAmor, ParametrosContabiliza) = False Then Exit Sub
    Contabiliza = RecuperaValor(ParametrosContabiliza, 1) = "1"
    'Si contabilizamos hay k conseguir el numero de asiento
    Set Mc = New Contadores
    If Contabiliza Then
        B = (Mc.ConseguirContador("0", (i = 0), True) = 0)
    Else
        B = True
    End If
    
    If B Then
        Screen.MousePointer = vbHourglass
        
        'Grabamos el LOG
        Cad = "Fecha amortización: " & txtFecAmo.Text
        If Mc.Contador > 0 Then Cad = Cad & " Asiento asignado: " & Mc.Contador
        vLog.Insertar 13, vUsu, Cad

        
        
        PreparaBloquear
            Conn.BeginTrans
            Cad = "Select * from inmovele where   inmovele.fecventa is null and inmovele.valoradq > inmovele.amortacu and situacio=1"
            'Fecha adq
            Cad = Cad & " and fechaadq <='" & Format(CDate(txtFecAmo.Text), FormatoFecha) & "'"
            Cad = Cad & " for update "
            B = GeneraCalculoInmovilizado(Cad, 2)
            If B Then
                Conn.CommitTrans
            Else
                Conn.RollbackTrans
            End If
        TerminaBloquear
        pb1.Visible = False
        Screen.MousePointer = vbDefault
        If B Then
            'ha ido bien
            MsgBox "El cálculo se ha realizado con éxito. En la introducción de apuntes esta el asiento generado.", vbExclamation
            Set Mc = Nothing
            Unload Me
            Exit Sub
        Else
            If Contabiliza Then Mc.DevolverContador "0", (i = 0), Mc.Contador
        End If
    End If
    Set Mc = Nothing
End Sub

Private Sub cmdDeshaz_Click(Index As Integer)
    If Index = 1 Then
        'Hacemos deshacer
        Cad = "¿Seguro que desea deshacer la última amortizacion con fecha: " & Format(UltAmor, "dd/mm/yyyy")
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Screen.MousePointer = vbHourglass
        Set Rs = New ADODB.Recordset
        
        Me.Tag = Label13(6).Caption
        DeshacerUltimaAmortizacion
        'Ha habido error
        If Me.cmdDeshaz(1).Enabled Then
            Label13(6).Caption = Me.Tag
        Else
            Me.cmdDeshaz(0).Caption = "Salir"
        End If
        Set Rs = Nothing
        Screen.MousePointer = vbDefault
    Else
        Unload Me
    End If
    
End Sub

Private Sub cmdListado1_Click(Index As Integer)
Dim RC As String
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    If txtConce(3).Text <> "" And txtConce(2).Text <> "" Then
        If Val(txtConce(2).Text) > Val(txtConce(3).Text) Then
            MsgBox "Concepto desde mayor concepto hasta", vbExclamation
            Exit Sub
        End If
    End If
    Cad = ""
    For i = 0 To Me.chkEstadisticas.Count - 1
        If Me.chkEstadisticas(i).Value = 1 Then Cad = Cad & "1"
    Next i
    If Cad = "" Then
        MsgBox "Seleccione, al menos, un tipo de situacion  a inculir en el informe.", vbExclamation
        Exit Sub
    End If
    'Llegados aqui generaremos el sql
    'La primera parte es comun para todos los informes
    ParametrosContabiliza = "SELECT codconam,nomconam,codinmov,nominmov,tipoamor,coeficie,"
    ParametrosContabiliza = ParametrosContabiliza & "codprove,fechaadq,valoradq,amortacu,fecventa,impventa "
    ParametrosContabiliza = ParametrosContabiliza & " FROM inmovcon,inmovele WHERE inmovcon.codconam=inmovele.conconam"
    RC = ""
    
    If Len(Cad) = 4 Then
        Cad = ""  'HA seleccionado todos las situaciones
        
    Else
        'NO Estan todos seleccionados, pq si no la longitud seria 4
        Cad = ""
        For i = 0 To Me.chkEstadisticas.Count - 1
            If Me.chkEstadisticas(i).Value Then
              If Cad <> "" Then Cad = Cad & " OR "
              Select Case i
                Case 0
                  RC = RC & "Activo - "
                Case 1
                  RC = RC & "Vendido - "
                Case 2
                    RC = RC & "Baja - "
                Case 3
                    RC = RC & "Totalmente amortizado"
              End Select
              Cad = Cad & " situacio = " & i + 1
            End If
        Next i
        
    End If
    
    
    AUX2 = ""
    TipoIva = ""
    'Desde hasta concepto
    If txtConce(2).Text <> "" Then
        AUX2 = AUX2 & " codconam >= " & txtConce(2).Text
        'Texto para cabecera informe
        TipoIva = "Desde concepto : " & txtConce(2).Text & " - " & Text1(2).Text
    End If
    
    If txtConce(3).Text <> "" Then
        If AUX2 <> "" Then AUX2 = AUX2 & " AND "
        AUX2 = AUX2 & " codconam <= " & txtConce(3).Text
        'Texto para cabecera informe
        If TipoIva <> "" Then
            TipoIva = TipoIva & "  h"
        Else
            TipoIva = "H"
        End If
        TipoIva = TipoIva & "asta concepto : " & txtConce(3).Text & " - " & Text1(3).Text
    End If
    
    If Cad <> "" Then Cad = " AND (" & Cad & ")"
    If AUX2 <> "" Then
        Cad = Cad & " AND "
        Cad = Cad & AUX2
    End If
    
    If TipoIva <> "" Then  'Los textos en la cabecera del informe
        If RC <> "" Then RC = RC & """ + Chr(13) + """
        RC = RC & TipoIva
    End If
    
    RC = "CampoSeleccion = """ & RC & """|"
    Cad = ParametrosContabiliza & Cad
    If ListadoEstadisticas(Cad) Then
        With frmImprimir
            .OtrosParametros = RC
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .opcion = 26
            .Show vbModal
        End With
    End If

End Sub

Private Sub cmdListado2_Click(Index As Integer)
Dim RC As String


    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    '------------------------------------------------
    'Listado entre fechas
    '
    '
    '-----------------------------------------------
    'Comprobar valores
    If txtConce(6).Text <> "" And txtConce(7).Text <> "" Then
        If Val(txtConce(6).Text) > Val(txtConce(7).Text) Then
            MsgBox "Concepto desde mayor concepto hasta", vbExclamation
            Exit Sub
        End If
    End If
    
    If txtfec(0).Text <> "" And txtfec(1).Text <> "" Then
        If CDate(txtfec(0).Text) > CDate(txtfec(1).Text) Then
            MsgBox "Fecha desde mayor fecha hasta", vbExclamation
            Exit Sub
        End If
    End If
    
    If txtfec(0).Text = "" Or txtfec(1).Text = "" Then
        MsgBox "Las fechas son obligadas", vbExclamation
        Exit Sub
    End If
    
    'Opciones seleccionadas
    Cad = ""
    For i = 4 To 7
        If Me.chkEstadisticas(i).Value = 1 Then Cad = Cad & "1"
    Next i
    If Cad = "" Then
        MsgBox "Seleccione, al menos, un tipo de situacion  a inculir en el informe.", vbExclamation
        Exit Sub
    End If
    
    'MODIFICACION ENERO 2005
    '---------------------------
    'Al no poner LEFT JOIN, los de centro de cste NULL, no salian
    'Llegados aqui generaremos el sql
    
    'La primera parte es comun para todos los informes
    ParametrosContabiliza = "SELECT codconam,nomconam,codinmov,nominmov,tipoamor,coeficie,"
    ParametrosContabiliza = ParametrosContabiliza & "codprove,fechaadq,valoradq,amortacu,fecventa,impventa,nomccost,ccoste.codccost "
    'ANtes
    'ParametrosContabiliza = ParametrosContabiliza & " FROM inmovcon,inmovele,ccoste WHERE inmovcon.codconam=inmovele.conconam and inmovele.codccost=ccoste.codccost"
    'AHora
    ParametrosContabiliza = ParametrosContabiliza & " FROM (inmovele INNER JOIN inmovcon ON inmovele.conconam = inmovcon.codconam) LEFT JOIN ccoste ON inmovele.codccost = ccoste.codccost"
    ParametrosContabiliza = ParametrosContabiliza & " WHERE 1=1 "  'ASi no cambio nada
    
    
    
    
    RC = ""
    If Len(Cad) = 4 Then
        Cad = ""  'HA seleccionado todos las situaciones
    Else
        'NO Estan todos seleccionados, pq si no la longitud seria 4
        Cad = ""
        For i = 4 To 7
            If Me.chkEstadisticas(i).Value Then
              If Cad <> "" Then Cad = Cad & " OR "
              Select Case i
                Case 4
                  RC = RC & "Activo - "
                Case 5
                  RC = RC & "Vendido - "
                Case 6
                    RC = RC & "Baja - "
                Case 7
                    RC = RC & "Totalmente amortizado"
              End Select
              Cad = Cad & " situacio = " & i - 3
            End If
        Next i
        If Cad <> "" Then Cad = "(" & Cad & ")"
    End If
  
    'Fecha inicio
    If RC <> "" Then RC = RC & "  "
    RC = RC & "Fechas desde " & txtfec(0).Text
    'Fecha fin
    RC = RC & " hasta " & txtfec(1).Text
    'Primer parametro en el informe
    AUX2 = "CampoSeleccion= """ & RC & """|"
    RC = ""
    
    'CENTROS DE COSTE
    If txtCodCCost(0).Text <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        RC = "Desde C. coste: " & txtCodCCost(0).Text & " - " & txtNomCcost(0).Text
        Cad = Cad & "ccoste.codccost >= '" & txtCodCCost(0).Text & "'"
    End If
    If txtCodCCost(1).Text <> "" Then
        If RC <> "" Then
            RC = RC & "   h"
        Else
            RC = "H"
        End If
        If Cad <> "" Then Cad = Cad & " AND "
        RC = RC & "asta C. coste: " & txtCodCCost(1).Text & " - " & txtNomCcost(1).Text
        Cad = Cad & "ccoste.codccost <= '" & txtCodCCost(1).Text & "'"
    End If
    'Segundo parametros
    RC = "Centro= """ & RC & """|"
    AUX2 = AUX2 & RC
    
    
    '
    'CONCEPTO INMOVILIZADO
    '-------------------------------------------------------------------------------
    RC = ""
    TipoIva = ""
    If txtConce(6).Text <> "" Then
        RC = "Desde conc. inmo. " & txtConce(6).Text & " - " & Text1(6).Text
        TipoIva = "codconam >=" & txtConce(6).Text
    End If
    If txtConce(7).Text <> "" Then
        If RC <> "" Then
            RC = RC & "     h"
        Else
            RC = "H"
        End If
        If TipoIva <> "" Then TipoIva = TipoIva & " AND "
        RC = RC & "asta conc. inmo. " & txtConce(7).Text & " - " & Text1(7).Text
        TipoIva = TipoIva & "codconam <=" & txtConce(7).Text
    End If
    If TipoIva <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & TipoIva
    End If
    
    'TErcer parametro
    RC = "Conceptos= """ & RC & """|"
    AUX2 = AUX2 & RC
    'Los volcamos sobre rc
    RC = AUX2
    
    If Cad <> "" Then ParametrosContabiliza = ParametrosContabiliza & " AND " & Cad
    
    'Dejamos abierto a cambiar el codinmov
    ParametrosContabiliza = ParametrosContabiliza & " AND codinmov = "

    If GeneraDatosListado Then
        With frmImprimir
            .OtrosParametros = RC
            .NumeroParametros = 3
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .opcion = 28
            .Show vbModal
        End With
    End If
End Sub





Private Sub cmdSimula_Click()
Dim TEXTO As String
    If Me.txtFecha.Text = "" Then
        MsgBox "Inserte la fecha de la simulación.", vbExclamation
        Exit Sub
    End If
    If Me.Tag <> "" Then
        If CDate(Me.txtFecha.Text) <= CDate(Me.Tag) Then
            MsgBox "Fecha no puede ser menor que la ultima fecha de amortizacion: " & Me.Tag, vbExclamation
            Exit Sub
        End If
    End If
    Cad = ""
    TEXTO = ""
    If txtConce(0).Text <> "" Then
        Cad = "inmovele.conconam >=" & txtConce(0).Text
        TEXTO = TEXTO & "Desde concepto " & txtConce(0).Text & " - " & Text1(0).Text
    End If
    If txtConce(1).Text <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & "inmovele.conconam <=" & txtConce(1).Text
        If TEXTO = "" Then
            TEXTO = "H"
        Else
            TEXTO = TEXTO & "     h"
        End If
        TEXTO = TEXTO & "asta " & txtConce(1).Text & " - " & Text1(1).Text
    End If
    If TEXTO <> "" Then
        TEXTO = """ + chr(13) + """ & TEXTO
    End If
    TEXTO = "Fecha simulación: " & txtFecha.Text & TEXTO
    TEXTO = "CampoSeleccion= """ & TEXTO & """|"
    If HazSimulacion(Cad, txtFecha.Text, 0) Then
        With frmImprimir
                .OtrosParametros = TEXTO
                .NumeroParametros = 1
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                .opcion = 25
                .Show vbModal
        End With
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Command1_Click()
  DatosOK
  If Cad <> "" Then Exit Sub
    
  If InsertarModificar Then
    Me.Tag = "1"
    Habilitar False
    Me.Toolbar1.Buttons(1).Enabled = True
  Else
    If Me.Tag = "1" Then
        CargarDatos
    Else
        Limpiar Me
    End If
End If
End Sub



Private Sub Command2_Click()
If Me.Tag = "1" Then
    CargarDatos
Else
    Limpiar Me
End If
Habilitar False
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub



'BAJA / VENTA
Private Sub Command5_Click()
Dim Adelante As Boolean
Dim ContaLinASi As Long
Dim F As Date
'Comprobamos k esta el elemento
    If Text6(0).Text = "" Then
        MsgBox "El elemento no puede estar vacio", vbExclamation
        Exit Sub
    End If
    
    'Comprobamos k la fecha esta puesta
    If Text4(0).Text = "" Then
        MsgBox "Ponga la fecha de baja/venta", vbExclamation
        Text4(0).SetFocus
        Exit Sub
    End If
    If txtCta(0).Text = "" Then
        MsgBox "Introduzca la cuenta de pérdidas/beneficios.", vbExclamation
        txtCta(0).SetFocus
        Exit Sub
    End If
    'Si esta bloqueada
    If EstaLaCuentaBloqueada(txtCta(0).Text, CDate(Text4(0).Text)) Then
        MsgBox "Cuenta bloqueada: " & txtCta(0).Text, vbExclamation
        Exit Sub
    End If

    
    'Si es venta tenemos k comprobar tb el importe y la cta de cliente
    If Option1(0).Value Then
        'Es venta
        'Comprobamos importe
        If Text5.Text = "" Then
            MsgBox "Introduzca el importe de la venta.", vbExclamation
            Text5.SetFocus
            Exit Sub
        End If
            
        If txtCta(1).Text = "" Then
            MsgBox "Introduzca la cuenta de venta.", vbExclamation
            txtCta(1).SetFocus
            Exit Sub
        End If
        '
        If EstaLaCuentaBloqueada(txtCta(1).Text, CDate(Text4(0).Text)) Then
            MsgBox "Cuenta bloqueada: " & txtCta(1).Text, vbExclamation
            Exit Sub
        End If
        
        'Si la cuenta necesita CC
        If vParam.autocoste Then
            If HayKHabilitarCentroCoste(txtCta(0).Text) Then
                If Me.txtCodCCost(2).Text = "" Then
                    MsgBox "Debe poner el Centro de coste para la cuenta base", vbExclamation
                    Exit Sub
                End If
            End If
        End If
        'Si tiene tesoreria y es una venta. Entonce introducimos el vencimiento
        If vEmpresa.TieneTesoreria Then
            Cad = ""
            If Text4(1).Text = "" Then Cad = "Falta fecha pago"
            If Text8(0).Text = "" Then Cad = "Falta forma pago"
            If txtCta(3).Text = "" Then Cad = "Falta cta prevista de pago"
            If Text8(2).Text = "" Then Cad = "Falta el agente"
            If Cad <> "" Then
                Cad = "Campos requeridos." & Cad
                MsgBox Cad, vbExclamation
                Exit Sub
            End If
            
            
            
            If EstaLaCuentaBloqueada(txtCta(3).Text, CDate(Text4(0).Text)) Then
                MsgBox "Cuenta bloqueada: " & txtCta(3).Text, vbExclamation
                Exit Sub
            End If
        End If
    End If


    i = FechaCorrecta2(CDate(Text4(0).Text))
    If i > 1 Then
        If i = 2 Then
            MsgBox varTxtFec, vbExclamation
        Else
            If i = 3 Then
                MsgBox "Fecha  pertence a un ejercicio cerrado.", vbExclamation
            Else
                MsgBox "Fecha  pertenece a un ejercicio todavia no abierto", vbExclamation
            End If
        End If
        Exit Sub
    End If



    If ObtenerparametrosAmortizacion(DivMes, UltAmor, ParametrosContabiliza) = False Then
        MsgBox "Error obteniendo datos parametros amortización", vbExclamation
        Exit Sub
    End If
    
    
    If CDate(Text4(0).Text) <= UltAmor Then
        MsgBox "La fecha es menor que la ultima fecha de amortizacion", vbExclamation
        Exit Sub
    End If
        
        
    'Tenemos que comprobar si la fecha es mayor que la proxima fecha amortizacion
    F = CDate(SugerirFechaNuevo)
    Debug.Print F
    If CDate(Text4(0).Text) > F Then
        MsgBox "La fecha venta/baja es mayor que la próxima fecha de amortizacion(" & Format(F, "dd/mm/yyyy") & ")", vbExclamation
        Exit Sub
    End If
    
        
    'Si es venta tenemos k comprobar tb el importe y la cta de cliente
    If Option1(0).Value And vEmpresa.TieneTesoreria Then
        If CDate(Text4(1).Text) <= vParam.fechaini Then
            AUX2 = "La fecha de cobro es menor que la fecha de ejercicio" & vbCrLf & "¿Continuar?"
            If MsgBox(AUX2, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    End If
     
    If Not ComprobaDatosVentaBajaElemento Then Exit Sub
        
        
    
    AUX2 = ""


'Llegados aqui todo bien, con lo cual hacemos ya lo siguiente
'------------------------------------------------------------
    If ObtenerparametrosAmortizacion(DivMes, UltAmor, ParametrosContabiliza) = False Then Exit Sub
    Contabiliza = RecuperaValor(ParametrosContabiliza, 1) = "1"
    'Si contabilizamos hay k conseguir el numero de asiento
    If Contabiliza Then
        Set Mc = New Contadores
        B = (Mc.ConseguirContador("0", CDate(Text4(0).Text) <= vParam.fechafin, True) = 0)
    Else
        B = True
    End If
    
    If B Then
        Screen.MousePointer = vbHourglass
        PreparaBloquear
        Conn.BeginTrans
        Adelante = False
        If Option1(0).Value Then
            i = 0
        Else
            i = 1
        End If
  
  
  
        'Intentamos cargar los datos
        If CargarDatosInmov Then
            'Veremos si ya esta totalmente amortizado o no.
            'Si  lo esta entonces generaremos la cabecera del apunte desde aqui, si no, al realizzar la amortizacion la crea
            If HayQueAmortizar Then
               'Cad = "Select * from sinmov where codinmov=" & Text6.Text & " for update "
                'cont=1  -> Lo inicaliza en el modulo
               B = GeneraCalculoInmovilizado(Cad, CByte(i))
               
               'Volvemos a cargar los datos despues de la amortizacion
               If B Then
                    Rs.Close   'Cierro el RS. para volverlo abrir con los datos actualizados de amortiz
                    B = CargarDatosInmov
               End If
               
            Else
                Cont = 1 'Contador para las lineas de asiento
                B = GeneracabeceraApunte(CByte(i))
            End If
            'Contador de asiento
            ContaLinASi = Cont
            If B Then
            
                'Modificacion del 26 de Abril. Si hay venta se vende, pero
                'la cancelacion del elemento se produce siempre
            
                If Option1(0).Value Then
                    'VENTA ---------------------------------------
                    CadenaDesdeOtroForm = ""  'para guardar datos y despues pasarlos a la factura impresa
                    Adelante = VentaElemento
                    
                    'Aqui tb habra que recargar el elemento
                    
                Else
                    Adelante = True
                End If
                
                If Adelante Then
                    'BAJA
                    Cont = ContaLinASi
                    Adelante = CancelarCuentaElemento
                End If
            End If
            Set Rs = Nothing
        End If
        If Adelante Then
            Conn.CommitTrans
        Else
            Conn.RollbackTrans
            B = False
        End If
        TerminaBloquear
        pb1.Visible = False
        Screen.MousePointer = vbDefault
        If B Then
            If Option1(0).Value Then
                EmiteFacturaVentaInmmovilizado
                Cad = RecuperaValor(ParametrosContabiliza, 5)
                Cad = "Preimpreso= " & Cad & "|"
                With frmImprimir
                    .OtrosParametros = Cad
                    .NumeroParametros = 1
                    Cad = "{ado.codusu}=" & vUsu.Codigo
                    .FormulaSeleccion = Cad
                    .SoloImprimir = False
                    'Opcion dependera del combo
                    .opcion = 48
                    .Show vbModal
                End With
            Else
                'ha ido bien
                MsgBox "Venta / Baja realizada.", vbInformation
            End If
            Limpiar Me
            Unload Me
        Else
            If Contabiliza Then Mc.DevolverContador "0", Option1(0).Value, Mc.Contador
        End If
    End If
    Set Mc = Nothing
End Sub




Private Function ComprobaDatosVentaBajaElemento() As Boolean

    ComprobaDatosVentaBajaElemento = False
    If ObtenerparametrosAmortizacion(DivMes, UltAmor, ParametrosContabiliza) = False Then Exit Function

    
    If Not HazSimulacion("codinmov =" & Text6(0).Text, CDate(Text4(0).Text), 1) Then Exit Function
    
    'Ahora, en ztmpsimula tengo los datos del elmento
    Set Rs = New ADODB.Recordset
    
    
    Cad = "Select valoradq,amortacu,totalamor from Usuarios.zsimulainm where codusu = " & vUsu.Codigo
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = "      "
    Importe = 0
    AUX2 = ""
    B = False
    
    
    
    
    If Not Rs.EOF Then
        B = True
        AUX2 = "Importe adq : " & Cad & Format(Rs!valoradq, FormatoImporte) & vbCrLf
        AUX2 = AUX2 & "Amort. acum. : " & Cad & Format(Rs!amortacu, FormatoImporte) & vbCrLf
        Importe = Rs!valoradq - Rs!amortacu
        AUX2 = AUX2 & "Pendiente:     " & Cad & Format(Importe, FormatoImporte) & vbCrLf & vbCrLf
        
        
        AUX2 = AUX2 & "Amort. periodo : " & Cad & Format(Rs!totalamor, FormatoImporte) & vbCrLf
        Importe = Importe - Rs!totalamor
        'Si es venta.
        If Option1(0).Value Then
            AUX2 = AUX2 & "Importe venta : " & Cad & Format(CCur(Text5.Text), FormatoImporte) & vbCrLf
            Importe = Importe - CCur(Text5.Text)
        End If

        
    Else
        AUX2 = "- Totalmente amortizado" & vbCrLf & vbCrLf
        
        'Si es venta, todo sera ganancias
        If Option1(0).Value Then
            B = True
            Importe = -1 * CCur(Text5.Text)
        End If
    End If
    Rs.Close
    
    
    If B Then
        TipoIva = String(35, "*") & vbCrLf
        If Importe > 0 Then
            'Significa que a la baja o a la venta, falta por amortizar
            'Con lo cual vamos a una cuenta de perdidas
            AUX2 = AUX2 & "Pérdidas inm.: " & Cad & Format(CCur(Importe), FormatoImporte) & vbCrLf & vbCrLf
            If Mid(txtCta(0).Text, 1, 1) <> "6" Then AUX2 = AUX2 & TipoIva & "Deberia poner una cuenta de PERDIDAS" & vbCrLf & TipoIva
        Else
            AUX2 = AUX2 & "Ganancias inm.: " & Cad & Format(CCur(Abs(Importe)), FormatoImporte) & vbCrLf & vbCrLf
            If Mid(txtCta(0).Text, 1, 1) <> "7" Then AUX2 = AUX2 & TipoIva & "Deberia poner una cuenta de GANANCIAS" & vbCrLf & TipoIva
        End If
    End If
    
    
    
    TipoIva = ""
    'En importe tengo lo que me faltaria amortizar. Con lo cual. Lo que venda, o de de baja, ira a perdidas
    'o ganancias del grupo 6 o del 7

    If Option1(0).Value Then
        Cad = "venta"
    Else
        Cad = "baja"
    End If
    Cad = "Va a realizar la " & Cad & " del "
    
    
    
    AUX2 = Cad & "elemento:" & vbCrLf & vbCrLf & Text6(0).Text & " - " & Text7(0).Text & vbCrLf & vbCrLf & AUX2
    
    

    AUX2 = AUX2 & vbCrLf & vbCrLf & "¿Continuar?"
    If MsgBox(AUX2, vbQuestion + vbYesNo) = vbNo Then Exit Function
    ComprobaDatosVentaBajaElemento = True
End Function



Private Sub Command6_Click(Index As Integer)
Dim RC As String
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    
    'Opciones seleccionadas
    Cad = ""
    For i = 8 To 11
        If Me.chkEstadisticas(i).Value = 1 Then Cad = Cad & "1"
    Next i
    If Cad = "" Then
        MsgBox "Seleccione, al menos, un tipo de situacion  a inculir en el informe.", vbExclamation
        Exit Sub
    End If
    RC = ""
    If Len(Cad) = 4 Then
        Cad = ""  'HA seleccionado todos las situaciones
    Else
        'NO Estan todos seleccionados, pq si no la longitud seria 4
        Cad = ""
        TipoIva = ""
        For i = 8 To 11
            If Me.chkEstadisticas(i).Value Then
              If Cad <> "" Then Cad = Cad & " OR "
              If RC <> "" Then TipoIva = " - "
              Select Case i
                Case 8
                  RC = RC & TipoIva & "Activo"
                Case 9
                  RC = RC & TipoIva & "Vendido"
                Case 10
                    RC = RC & TipoIva & "Baja"
                Case 11
                    RC = RC & TipoIva & "Totalmente amortizado"
              End Select
              Cad = Cad & " situacio = " & i - 7
            End If
        Next i
        If Cad <> "" Then Cad = "(" & Cad & ")"
    End If
  
    'Elemento
    If Text6(1).Text <> "" And Text6(2).Text <> "" Then
        If Val(Text6(1).Text) > Val(Text6(2).Text) Then
            MsgBox "Elemento desde mayor elemento hasta", vbExclamation
            Exit Sub
        End If
    End If
    
    TipoIva = ""
    If Text6(1).Text <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & "inmovele.codinmov >= " & Text6(1).Text
        TipoIva = "Elto. desde: " & Text6(1).Text & "- " & Text7(1).Text
    End If
    If Text6(2).Text <> "" Then
        If Cad <> "" Then
            Cad = Cad & " AND "
            TipoIva = TipoIva & "    "
        End If
        Cad = Cad & " inmovele.codinmov <= " & Text6(2).Text
        TipoIva = TipoIva & "Elto hasta: " & Text6(2).Text & "- " & Text7(2).Text
    End If
    AUX2 = "Eltos= """ & TipoIva & """"
    
    
    
    
    
    
    'Concepto
    If txtConce(4).Text <> "" And txtConce(5).Text <> "" Then
        If Val(txtConce(4).Text) > Val(txtConce(5).Text) Then
            MsgBox "Concepto desde mayor concepto hasta", vbExclamation
            Exit Sub
        End If
    End If
    
    TipoIva = ""
    If txtConce(4).Text <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & "codconam >= " & txtConce(4).Text
        TipoIva = "Concepto desde: " & Text1(4).Text
    End If
    If txtConce(5).Text <> "" Then
        If Cad <> "" Then
            Cad = Cad & " AND "
            TipoIva = TipoIva & "    "
        End If
        Cad = Cad & " codconam <= " & txtConce(5).Text
        TipoIva = TipoIva & "Concepto hasta: " & Text1(5).Text
    End If
    ParametrosContabiliza = "SELECT inmovcon.codconam, inmovcon.nomconam, inmovele.codinmov, inmovele.nominmov, inmovele.fechaadq,"
    ParametrosContabiliza = ParametrosContabiliza & "inmovele.valoradq, inmovele_his.fechainm, inmovele_his.imporinm, inmovele_his.porcinm"
    ParametrosContabiliza = ParametrosContabiliza & " FROM inmovele,inmovcon ,inmovele_his WHERE    inmovele.conconam = inmovcon.codconam"
    ParametrosContabiliza = ParametrosContabiliza & " and inmovele.codinmov = inmovele_his.codinmov"
    
    If Cad <> "" Then Cad = " AND " & Cad
    Cad = ParametrosContabiliza & Cad
    Cad = Cad & " ORDER BY codconam,codinmov,fechainm"
    If RC <> "" Then RC = "      " & RC
    TipoIva = Trim(TipoIva & RC)
    TipoIva = "CampoSeleccion= """ & TipoIva & """|"
    'El salto por grupo
    TipoIva = TipoIva & "Salto= " & Abs(Check3.Value) & "|"
    TipoIva = TipoIva & AUX2 & "|"
    If ListadoFichaInmo(Cad) Then
        With frmImprimir
            .OtrosParametros = TipoIva
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .opcion = 27
            .Show vbModal
        End With
    End If
        
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_Activate()
If PrimeraVez Then
    PrimeraVez = False

    Select Case opcion
    Case 0
        If Not CargarDatos Then
            Me.Tag = "0"
            Toolbar1.Buttons(1).Enabled = False
            Habilitar True
            Combo1.SetFocus
        Else
            Me.Tag = "1"
        End If
        Command2.Cancel = True
    Case 1
        Me.Command3.Cancel = True
    Case 3
        Me.Command4.Cancel = True
    Case 4
        'ANTES
'        For I = 0 To 3
'            Me.chkEstadisticas(I).Value = 1
'        Next I
        Me.chkEstadisticas(0).Value = 1
        'Me.cmdListado2(1).Cancel = True
    Case 5
        Command6(1).Cancel = True
'        For I = 8 To 11
'            Me.chkEstadisticas(I).Value = 1
'        Next I
        Me.chkEstadisticas(8).Value = 1
    Case 6
'        For I = 4 To 7
'            Me.chkEstadisticas(I).Value = 1
'        Next I
        Me.chkEstadisticas(4).Value = 1
        
    Case 10
        'Deshacer ultima amortizacion
        CargarDatosAmortizacion
    End Select
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    Me.Icon = frmPpal.Icon

    Set miTag = New CTag
    Limpiar Me
    pb1.Visible = False
    PrimeraVez = True
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 4
        .Buttons(2).Image = 15
    End With
    Toolbar1.Buttons(1).Visible = False ''El boton de modificar
    
    Frame0.Visible = False
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    Frame5.Visible = False
    Frame6.Visible = False
    FrDeshacer.Visible = False
    Select Case opcion
    Case 0
        Toolbar1.Buttons(1).Visible = True ''El boton de modificar
        Frame0.Visible = True
        Frame0.Enabled = False
        Me.Width = Frame0.Width + 150
        Me.Height = Frame0.Height + 800
        Me.Command1.Visible = False
        Me.Command2.Visible = False
        Caption = "Parámetros inmovilizado"
    Case 1
        txtFecha.Text = SugerirFechaNuevo
        
        Frame1.Visible = True
        Me.Width = Frame1.Width + 150
        Me.Height = Frame1.Height + 500
        'Caption = "Simulación amortización"
    Case 2
        txtFecAmo.Text = SugerirFechaNuevo
        txtFecAmo.Enabled = vUsu.Nivel < 2
        Frame2.Visible = True
        Me.Width = Frame2.Width + 150
        Me.Height = Frame2.Height + 1000
        Caption = "Amortización"
    Case 3
        Frame3.Visible = True
        Me.Width = Frame3.Width + 150
        If vEmpresa.TieneTesoreria Then
            Frame3.Height = 7080
            
        Else
            Frame3.Height = 4800
        End If
        Me.Command4.Top = Frame3.Height - 480
        Me.Command5.Top = Command4.Top
        Me.FrameTesor.Visible = vEmpresa.TieneTesoreria
        Me.Height = Frame3.Height + 500
        
        
        txtCodCCost(2).Visible = vParam.autocoste
        Label14(7).Visible = vParam.autocoste
        imgCCost(2).Visible = vParam.autocoste
        txtNomCcost(2).Visible = vParam.autocoste
        'Caption = "Venta / Baja"
    Case 4
        Frame4.Visible = True
        Me.Width = Frame4.Width + 150
        Me.Height = Frame4.Height + 400
        'Caption = "Listado estadisticas"
        cmdListado1(1).Cancel = True
    Case 5
        Frame5.Visible = True
        Me.Width = Frame5.Width + 150
        Me.Height = Frame5.Height + 500
        'Caption = "Ficha inmovilizado"
    Case 6
        Frame6.Visible = True
        Me.Width = Frame6.Width + 150
        Me.Height = Frame6.Height + 500
        txtfec(0).Text = Format(vParam.fechaini, "dd/mm/yyyy")
        txtfec(1).Text = Format(vParam.fechafin, "dd/mm/yyyy")
        cmdListado2(1).Cancel = True
        'Caption = "Estdis. inmovilizado"
        
    Case 10
        FrDeshacer.Visible = True
        Me.Width = FrDeshacer.Width + 150
        Me.Height = FrDeshacer.Height + 500
        Caption = "Deshacer"
    End Select


        
    '0.- Parametros
    '1.- Simular
    '2.- Cálculo amort.
    '3.- Venta/Baja inmovilizado
    '--- los siguiente utilizan el mismo frame, con opciones
    '4.- Listado estadisticas
    '5.- Ficha elementos
    i = 0
    If opcion = 0 Or opcion = 2 Then i = 1
    Toolbar1.Visible = (i = 1)
    If i = 0 Then Caption = "Informes"

End Sub

Private Function SugerirFechaNuevo() As String
Dim RC As String
    RC = "tipoamor"
    Cad = DevuelveDesdeBD("ultfecha", "paramamort", "codigo", "1", "N", RC)

    If Cad <> "" Then
        Me.Tag = Cad   'Ultima actualizacion
        Select Case Val(RC)
        Case 2
            'Semestral
            i = 6
            'Siempre es la ultima fecha de mes
        Case 3
            'Trimestral
            i = 3
        Case 4
            'Mensual
            i = 1
        Case Else
            'Anual
            i = 12
        End Select
        RC = PonFecha
    Else
        Cad = "01/01/1991"
        RC = Format(Now, "dd/mm/yyyy")
    End If
    'If Simulacion Then
    '     txtFecha.Text = Format(RC, "dd/mm/yyyy")
    'Else
    '     txtFecAmo.Text = Format(RC, "dd/mm/yyyy")
    '     'Dejamos cambiar la fecha, si , y solo si, es administrador
    '     txtFecAmo.Enabled = vUsu.Nivel < 2
        
    'End If
    SugerirFechaNuevo = Format(RC, "dd/mm/yyyy")
    
End Function



Private Function PonFecha() As Date
Dim d As Date
'Dada la fecha en Cad y los meses k tengo k sumar
'Pongo la fecha
d = DateAdd("m", i, CDate(Cad))
Select Case Month(d)
Case 2
    If ((Year(d) - 2000) Mod 4) = 0 Then
        i = 29
    Else
        i = 28
    End If
Case 1, 3, 5, 7, 8, 10, 12
    '31
        i = 31
Case Else
    '30
        i = 30
End Select
Cad = i & "/" & Month(d) & "/" & Year(d)
PonFecha = CDate(Cad)
End Function


Private Function CargarDatos() As Boolean
On Error GoTo ECargarDatos
    CargarDatos = False
    Set Rs = New ADODB.Recordset
    Cad = "Select * from paramamort where codigo=1"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        CargarDatos = True
        '------------------  Ponemos los datos
        Combo1.ListIndex = Rs!tipoamor - 1
        Check1.Value = Rs!intcont
        Text2(0).Text = Format(Rs!ultfecha, "dd/mm/yyyy")
        Text2(1).Text = DBLet(Rs!NumDiari)
        Text2_LostFocus 1
        Text2(2).Text = DBLet(Rs!condebes)
        Text2_LostFocus 2
        Text2(3).Text = DBLet(Rs!conhaber)
        Text2_LostFocus 3
        txtIVA(0).Text = DBLet(Rs!codiva)
        txtIVA_LostFocus 0
        Check2.Value = Rs!Preimpreso
    End If
    Rs.Close
ECargarDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando parametros"
    Set Rs = Nothing
End Function


Private Sub Form_Unload(Cancel As Integer)
    Set miTag = Nothing
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    If i = 0 Then
        Text8(0).Text = RecuperaValor(CadenaDevuelta, 1)
        Text8(1).Text = RecuperaValor(CadenaDevuelta, 2)
    Else
        Text8(2).Text = RecuperaValor(CadenaDevuelta, 1)
        Text8(3).Text = RecuperaValor(CadenaDevuelta, 2)
    End If
End Sub

Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    If i = 1 Then
        'Cuenta bancaria
        txtCta(3).Text = RecuperaValor(CadenaSeleccion, 1)
        txtDescta(3).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    Text2(i + 2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text3(i + 1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    txtCodCCost(i).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNomCcost(i).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCI_DatoSeleccionado(CadenaSeleccion As String)
    i = Val(imgConceInmo(0).Tag)
    txtConce(i).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(i).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCt_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(i).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescta(i).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 1)
    Text3(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmE_DatoSeleccionado(CadenaSeleccion As String)
    Cad = RecuperaValor(CadenaSeleccion, 3)
    If Cad = "" Or Cad = "1" Or Cad = "2" Then
        MsgBox "El elemento esta dado de baja o vendido", vbExclamation
        Exit Sub
    End If
    i = CInt(Me.imgElto(0).Tag)
    Text6(i).Text = RecuperaValor(CadenaSeleccion, 1)
    Text7(i).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Cad = Format(vFecha, "dd/mm/yyyy")
    Select Case i
    Case 0
        Text2(0).Text = Cad
    Case 1
        txtFecha.Text = Cad
    Case 2
        txtFecAmo.Text = Cad
    Case 3
        Text4(0).Text = Cad
    Case 4, 5
        txtfec(i - 4).Text = Cad
    Case 6
        Text4(1).Text = Cad
    End Select
End Sub

Private Sub frmI_DatoSeleccionado(CadenaSeleccion As String)
    txtIVA(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtIVA(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmF = New frmCal
    frmF.Fecha = Now
    i = Index
    Select Case Index
    Case 0
        If Text2(0).Text <> "" Then
            If IsDate(Text2(0).Text) Then frmF.Fecha = CDate(Text2(0).Text)
        End If
    Case 1
        If txtFecha.Text <> "" Then
            If IsDate(txtFecha.Text) Then frmF.Fecha = CDate(txtFecha.Text)
        End If
    Case 2
        If txtFecAmo.Text <> "" Then frmF.Fecha = CDate(txtFecAmo.Text)
    Case 3
        If Text4(0).Text <> "" Then frmF.Fecha = CDate(Text4(0).Text)
        
    Case 4, 5
        'Fechas inofrmes entre fechas
        If txtfec(Index - 4).Text <> "" Then frmF.Fecha = CDate(txtfec(Index - 4).Text)
    Case 6
        'Fec vencimiento
        If Text4(1).Text <> "" Then frmF.Fecha = CDate(Text4(1).Text)
    End Select
    frmF.Show vbModal
    Set frmF = Nothing
End Sub

Private Sub Image2_Click()
    Set frmD = New frmTiposDiario
    frmD.DatosADevolverBusqueda = "0|1|"
    frmD.Show vbModal
    Set frmD = Nothing
End Sub

Private Sub Image3_Click(Index As Integer)
    i = Index
    Set frmC = New frmConceptos
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub imgCCost_Click(Index As Integer)
    i = Index
    Set frmCC = New frmCCoste
    frmCC.DatosADevolverBusqueda = "0|1|"
    frmCC.Show vbModal
    Set frmCC = Nothing
End Sub

Private Sub imgConceInmo_Click(Index As Integer)
    imgConceInmo(0).Tag = Index
    Set frmCI = New frmInmoConceptos
    frmCI.DatosADevolverBusqueda = "0|1|"
    frmCI.Show vbModal
    Set frmCI = Nothing
End Sub

Private Sub imgcta_Click(Index As Integer)
    i = Index
    Set frmCt = New frmColCtas
    frmCt.DatosADevolverBusqueda = "0|1|"
    frmCt.Show vbModal
    Set frmCt = Nothing
End Sub

Private Sub imgElto_Click(Index As Integer)
    Set frmE = New frmInmoElto
    imgElto(0).Tag = Index
    frmE.DatosADevolverBusqueda = "0|1|"
    frmE.Show vbModal
    Set frmE = Nothing
End Sub

Private Sub imgiva_Click()
    Set frmI = New frmIVA
    frmI.DatosADevolverBusqueda = "0|1|"
    frmI.Show vbModal
    Set frmI = Nothing
End Sub

Private Sub imgTesoreria_Click(Index As Integer)
    i = Index
    Select Case Index
    Case 0
        'FORMA PAGO
        Set frmB = New frmBuscaGrid
        frmB.vCampos = "Codigo|codforpa|N|10·" & "Descripcion|nomforpa|T|60·"
        frmB.vTabla = "formapago"
        frmB.vDevuelve = "0|1|"
        frmB.vSelElem = 0
        frmB.vTitulo = "Forma pago"
        frmB.vSQL = ""
        frmB.Show vbModal
    Case 1
        'Cuenta prevista pago
        Set frmBa = New frmCuentasBancarias
        frmBa.DatosADevolverBusqueda = "0|1|"
        frmBa.Show vbModal
        Set frmBa = Nothing
    Case 2
        'Agente
        Set frmB = New frmBuscaGrid
        frmB.vCampos = "Codigo|codigo|N|10·" & "Nombre|nombre|T|60·"
        frmB.vTabla = "agentes"
        frmB.vDevuelve = "0|1|"
        frmB.vSelElem = 0
        frmB.vTitulo = "Agentes"
        frmB.vSQL = ""
        frmB.Show vbModal
    End Select
End Sub



Private Sub Option1_Click(Index As Integer)
    Me.FrameVenta.Visible = (Option1(0).Value)
    Me.FrameTesor.Visible = vEmpresa.TieneTesoreria And Me.FrameVenta.Visible
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    With Text2(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text2_LostFocus(Index As Integer)
    Text2(Index).Text = Trim(Text2(Index).Text)
    If Text2(Index).Text = "" Then
        If Index > 0 Then Text3(Index - 1).Text = ""
        Exit Sub
    End If
    If Index = 0 Then
        'Fecha
        If Not EsFechaOK(Text2(0)) Then
            MsgBox "Fecha incorrecta", vbExclamation
            Text2(0).Text = ""
            Text2(0).SetFocus
            Exit Sub
        End If
        Text2(0).Text = Format(Text2(0).Text)
    Else
        If Not IsNumeric(Text2(Index).Text) Then
            MsgBox "El campo tiene que ser numérico", vbExclamation
            Text2(Index).Text = ""
            Text2(Index).SetFocus
            Exit Sub
        End If
        Select Case Index
        Case 1
             Cad = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text2(1).Text, "N")
             If Cad = "" Then
                    MsgBox "Diario no encontrado: " & Text2(1).Text, vbExclamation
                    Text2(1).Text = ""
                    Text2(1).SetFocus
            End If
            Text3(0).Text = Cad
        Case 2, 3
                Cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text2(Index).Text, "N")
                If Cad = "" Then
                    MsgBox "Concepto NO encontrado: " & Text2(Index).Text, vbExclamation
                    Text2(Index).Text = ""
                End If
                Text3(Index - 1).Text = Cad
                
        End Select
    End If
End Sub

Private Sub Text4_LostFocus(Index As Integer)
    If Text4(Index).Text <> "" Then
        If Not EsFechaOK(Text4(Index)) Then
            MsgBox "Fecha incorrecta: " & Text4(Index).Text, vbExclamation
            Text4(Index).Text = ""
        Else
            Text4(Index).Text = Format(Text4(Index).Text, "dd/mm/yyyy")
        End If
    End If
End Sub

Private Sub Text5_GotFocus()
    With Text5
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text5_LostFocus()
    Text5.Text = Trim(Text5.Text)
    If Text5.Text = "" Then Exit Sub
    If Not IsNumeric(Text5.Text) Then
        MsgBox "Importe debe ser numérico: " & Text5.Text, vbExclamation
        Text5.SetFocus
    End If
    DivMes = InStr(1, Text5.Text, ",")
    If DivMes > 0 Then
        'Esta formateado
        Importe = ImporteFormateado(Text5.Text)
    Else
        Cad = TransformaPuntosComas(Text5.Text)
        Importe = CCur(Cad)
    End If
    Text5.Text = Format(Importe, FormatoImporte)
    
End Sub








Private Sub Text6_LostFocus(Index As Integer)

    With Text6(Index)
        .Text = Trim(.Text)
        If .Text = "" Then
            Text7(Index).Text = ""
            Exit Sub
        End If
        If Not IsNumeric(.Text) Then
            MsgBox "Elemento de inmovilizado debe ser numérico: " & .Text, vbExclamation
            .Text = ""
            .SetFocus
            Exit Sub
        End If
        ParametrosContabiliza = "situacio"
        Cad = DevuelveDesdeBD("nominmov", "inmovele", "codinmov", .Text, "N", ParametrosContabiliza)
        If Cad = "" Then
            MsgBox "elemento de inmovlizado NO encontrado: " & .Text, vbExclamation
        Else
            'Esta comprobacion solo es para la venta/baja
            If Index = 0 Then
                If ParametrosContabiliza = "2" Or ParametrosContabiliza = "3" Then
                    MsgBox "El elemento : " & Cad & " ya ha sido vendido o dado de baja", vbExclamation
                    Cad = ""
                End If
            End If
        End If
        Text7(Index).Text = Cad
        If Cad = "" Then
            .Text = ""
            .SetFocus
        End If
    End With
End Sub



Private Sub Text8_GotFocus(Index As Integer)
    PonFoco Text8(Index)
End Sub


Private Sub Text8_LostFocus(Index As Integer)
    Text8(Index).Text = Trim(Text8(Index).Text)
    
    If Index = 0 Then
        If Text8(0).Text = "" Then
            Text8(1).Text = ""
            Exit Sub
        End If
        If Not IsNumeric(Text8(0).Text) Then
            Cad = ""
            i = 1
        Else
            Cad = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", Text8(0).Text, "N")
            i = 2
        End If
        If Cad = "" Then
            Cad = "Error en forma pago."
            If i = 1 Then
                Cad = Cad & " Campo debe ser numérico"
            Else
                Cad = Cad & " No existe forma pago:" & Text8(0).Text
            End If
            MsgBox Cad, vbExclamation
            Text8(0).Text = ""
            Text8(1).Text = ""
        Else
            Text8(1).Text = Cad
        End If
    Else
        If Index = 2 Then
            If Text8(2).Text = "" Then
                Text8(3).Text = ""
                Exit Sub
            End If
        
            If Not IsNumeric(Text8(2).Text) Then
                Cad = ""
                i = 1
            Else
                Cad = DevuelveDesdeBD("nombre", "agentes", "codigo", Text8(2).Text, "N")
                i = 2
            End If
            If Cad = "" Then
                Cad = "Error en el agente."
                If i = 1 Then
                    Cad = Cad & " Campo debe ser numérico"
                Else
                    Cad = Cad & " No existe agente:" & Text8(2).Text
                End If
                MsgBox Cad, vbExclamation
                Text8(2).Text = ""
                Text8(3).Text = ""
            Else
                Text8(3).Text = Cad
            End If
        
        
        End If
        
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    Habilitar True
Case 2
    Unload Me
End Select
End Sub



Private Function InsertarModificar() As Boolean
On Error GoTo EInsertarModificar
    InsertarModificar = False

    If Me.Tag = "1" Then
        'Modificar
        Cad = "UPDATE paramamort SET tipoamor= " & Combo1.ListIndex + 1
        Cad = Cad & ", intcont= " & Check1.Value
        Cad = Cad & ", ultfecha= '" & Format(Text2(0).Text, FormatoFecha)
        Cad = Cad & "', condebes= " & ParaBD(Text2(2))
        Cad = Cad & ", conhaber= " & ParaBD(Text2(3))
        Cad = Cad & ", numdiari=  " & ParaBD(Text2(1))
        
        Cad = Cad & ", codiva = " & ParaBD(txtIVA(0))
        Cad = Cad & ", preimpreso =" & Check2.Value
        Cad = Cad & " WHERE codigo=1"
        
    Else
        'INSERTAR
        Cad = "INSERT INTO paramamort (codigo, tipoamor, intcont, ultfecha, condebes, conhaber, numdiari,codiva,preimpreso) VALUES (1,"
        Cad = Cad & Combo1.ListIndex + 1 & "," & Me.Check1.Value & ",'" & Format(Text2(0).Text, FormatoFecha)
        Cad = Cad & "'," & ParaBD(Text2(2)) & "," & ParaBD(Text2(3)) & "," & ParaBD(Text2(1))
        Cad = Cad & ",'" & ParaBD(txtIVA(0)) & "'," & Check2.Value & ")"
    End If
    Conn.Execute Cad
    InsertarModificar = True
    Exit Function
EInsertarModificar:
    MuestraError Err.Number, "Insertar-Modificar"
End Function


Private Function ParaBD(ByRef T As TextBox) As String
If T.Text = "" Then
    ParaBD = "NULL"
Else
    ParaBD = T.Text
End If
End Function




Private Sub DatosOK()
Dim i As Integer
Cad = "MAL"
If Combo1.ListIndex < 0 Then
    Cad = "Selecciona un tipo de amortización"
    MsgBox Cad, vbExclamation
    Exit Sub
End If
For i = 0 To 3
    miTag.Cargar Text2(i)
    If Not miTag.Comprobar(Text2(i)) Then Exit Sub
Next i

If Check1.Value = 1 Then
    'Contabiliza autmaticamente luego obligamos a poner el resto de campos contabilizacion
    For i = 1 To 3
        If Text2(i).Text = "" Then
            Cad = "Campos contabilizacion requeridos"
            MsgBox Cad, vbExclamation
            Exit Sub
        End If
    Next i
End If

Cad = ""
End Sub

Private Sub Habilitar(Veradero As Boolean)
    Frame0.Enabled = Veradero
    Command1.Visible = Veradero
    Command2.Visible = Veradero
End Sub




Private Sub txtCodCCost_GotFocus(Index As Integer)
    With txtConce(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCodCCost_LostFocus(Index As Integer)
    txtCodCCost(Index).Text = Trim(txtCodCCost(Index).Text)
    If txtCodCCost(Index).Text = "" Then
        txtNomCcost(Index).Text = ""
        Exit Sub
    End If
    Cad = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtCodCCost(Index).Text, "T")
    If Cad = "" Then
        MsgBox "C. coste NO encontrado: " & txtCodCCost(Index).Text, vbExclamation
        txtCodCCost(Index).Text = ""
        txtCodCCost(Index).SetFocus
    End If
    txtNomCcost(Index).Text = Cad
End Sub

Private Sub txtConce_GotFocus(Index As Integer)
    With txtConce(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtConce_LostFocus(Index As Integer)
    txtConce(Index).Text = Trim(txtConce(Index).Text)
    If txtConce(Index).Text = "" Then
        Text1(Index).Text = ""
        Exit Sub
    End If
    
    i = 0
    Cad = ""
    If Not IsNumeric(txtConce(Index).Text) Then
        MsgBox "Concepto debe ser numérico: " & txtConce(Index).Text, vbExclamation
        txtConce(Index).Text = ""
        i = 1
    End If
    If i = 0 Then
        Cad = DevuelveDesdeBD("nomconam", "inmovcon", "codconam", txtConce(Index).Text, "N")
        If Cad = "" Then
            MsgBox "Concepto no encontrado: " & txtConce(0).Text, vbExclamation
            txtConce(Index).Text = ""
        Else
            i = 1
        End If
    End If
    Text1(Index).Text = Cad
    If i = 0 Then txtConce(Index).SetFocus


End Sub


Private Sub txtCta_GotFocus(Index As Integer)
    With txtCta(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
With txtCta(Index)
    .Text = Trim(.Text)
    If .Text = "" Then
        txtDescta(Index).Text = ""
        Exit Sub
    End If
    ParametrosContabiliza = .Text
    If CuentaCorrectaUltimoNivel(ParametrosContabiliza, Cad) Then
        .Text = ParametrosContabiliza
        txtDescta(Index).Text = Cad
        If Index = 3 Then
            Cad = DevuelveDesdeBD("codmacta", "bancos", "codmacta", ParametrosContabiliza, "T")
            If Cad = "" Then
                MsgBox "Cuenta no asociada a ningun banco", vbExclamation
                .Text = ""
                txtDescta(Index).Text = ""
            End If
        End If
    Else
        MsgBox Cad, vbExclamation
        .Text = ""
        txtDescta(Index).Text = ""
        .SetFocus
    End If

End With
End Sub

Private Sub txtfec_GotFocus(Index As Integer)
    With txtfec(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtfec_LostFocus(Index As Integer)
    txtfec(Index).Text = Trim(txtfec(Index).Text)
    If txtfec(Index).Text <> "" Then
        If Not EsFechaOK(txtfec(Index)) Then
            MsgBox "Fecha incorrecta: " & txtfec(Index).Text, vbExclamation
            txtfec(Index).Text = ""
            txtfec(Index).SetFocus
        End If
    End If
End Sub

Private Sub txtFecAmo_GotFocus()
With txtFecAmo
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtFecAmo_KeyPress(KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        KEYFecAmo KeyAscii
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub txtFecAmo_LostFocus()
With txtFecAmo
    .Text = Trim(.Text)
    If .Text = "" Then Exit Sub
    If Not EsFechaOK(txtFecAmo) Then
        MsgBox "Fecha incorrecta: " & .Text, vbExclamation
        .Text = ""
        .SetFocus
    End If
End With
End Sub


Private Sub txtfecha_LostFocus()
txtFecha.Text = Trim(txtFecha.Text)
If txtFecha.Text = "" Then Exit Sub
If Not EsFechaOK(txtFecha) Then
    MsgBox "Fecha incorrecta: " & txtFecha.Text, vbExclamation
    txtFecha.Text = ""
    txtFecha.SetFocus
End If
    
End Sub

Private Sub txtIVA_LostFocus(Index As Integer)
If Index = 1 Then Exit Sub
txtIVA(0).Text = Trim(txtIVA(0).Text)
If txtIVA(0).Text = "" Then
    txtIVA(1).Text = ""
    Exit Sub
End If

i = 0
Cad = ""
If Not IsNumeric(txtIVA(0).Text) Then
    MsgBox "Tipo IVA debe ser numérico: " & txtIVA(0).Text, vbExclamation
    txtIVA(0).Text = ""
    i = 1
End If
If i = 0 Then
    Cad = DevuelveDesdeBD("nombriva", "tiposiva", "codigiva", txtIVA(0).Text, "N")
    If Cad = "" Then
        MsgBox "IVA no encontrado: " & txtIVA(0).Text, vbExclamation
        txtIVA(0).Text = ""
    End If
End If
txtIVA(1).Text = Cad
End Sub

'++
Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYCta KeyAscii, 0
            Case 1: KEYCta KeyAscii, 1
            Case 2: KEYCta KeyAscii, 2
            Case 3: KEYTesoreria KeyAscii, 1
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtConce_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYConcepto KeyAscii, 0
            Case 1: KEYConcepto KeyAscii, 1
            Case 2: KEYConcepto KeyAscii, 2
            Case 3: KEYConcepto KeyAscii, 3
            Case 4: KEYConcepto KeyAscii, 4
            Case 6: KEYConcepto KeyAscii, 6
            Case 7: KEYConcepto KeyAscii, 7
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtCodCCost_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYCCoste KeyAscii, 0
            Case 1: KEYCCoste KeyAscii, 1
            Case 2: KEYCCoste KeyAscii, 2
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub txtiva_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        KEYIva KeyAscii, Index
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        KEYFecha KeyAscii, 1
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtfec_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYFec KeyAscii, 4
            Case 1: KEYFec KeyAscii, 5
            Case 2: KEYFec KeyAscii, 2
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYFec KeyAscii, 0
            Case 1: KEYImage2 KeyAscii
            Case 2: KEYImage3 KeyAscii, 0
            Case 3: KEYImage3 KeyAscii, 1
            
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYFec KeyAscii, 3
            Case 1: KEYFec KeyAscii, 6
            
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYElto KeyAscii, 0
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYTesoreria KeyAscii, 0
            Case 2: KEYTesoreria KeyAscii, 2
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYTesoreria(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgTesoreria_Click (indice)
End Sub

Private Sub KEYElto(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgElto_Click (indice)
End Sub

Private Sub KEYCta(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgcta_Click (indice)
End Sub

Private Sub KEYIva(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgiva_Click
End Sub

Private Sub KEYConcepto(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgConceInmo_Click (indice)
End Sub

Private Sub KEYCCoste(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgCCost_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    Image1_Click (indice)
End Sub

Private Sub KEYFec(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    Image1_Click (indice)
End Sub

Private Sub KEYFecAmo(KeyAscii As Integer)
    KeyAscii = 0
    Image1_Click (2)
End Sub

Private Sub KEYImage2(KeyAscii As Integer)
    KeyAscii = 0
    Image2_Click
End Sub

Private Sub KEYImage3(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    Image3_Click (indice)
End Sub

'++


'TIPO:
'       0.- Venta
'       1.- Baja
'       2.- Calculo de amortizacion
Private Function GeneraCalculoInmovilizado(ByRef SeleccionInmovilizado As String, Tipo As Byte) As Boolean
Dim Codinmov As Long
Dim B As Boolean
On Error GoTo EGen

    GeneraCalculoInmovilizado = False
    If Tipo = 2 Then
        'Para el calculo del amortizado
        Set Rs = New ADODB.Recordset
        Rs.Open SeleccionInmovilizado, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        If Rs.EOF Then
            MsgBox "Ningun registro", vbExclamation
            Rs.Close
            Exit Function
        End If
    End If
    'Vemos cuantos hay
    Cont = 0
    While Not Rs.EOF
        Cont = Cont + 1
        Rs.MoveNext
    Wend
    Rs.MoveFirst
    If Cont > 3 Then pb1.Visible = True
    pb1.Max = Cont + 1
    pb1.Value = 0
    
    
    
    'Vemos si contabilizamos
    'Insertamos cabecera del asiento
    If Contabiliza Then GeneracabeceraApunte (Tipo)
    Cont = 1
    While Not Rs.EOF
        Codinmov = Rs!Codinmov
       
        'La fecha depende si estamos calculando normal o estamos vendiendo
        If opcion = 3 Then
            Cad = Text4(0).Text
        Else
            Cad = Me.txtFecAmo.Text
        End If
      
        B = CalculaAmortizacion(Codinmov, CDate(Cad), DivMes, UltAmor, ParametrosContabiliza, Mc.Contador, Cont, Tipo < 2)
        If Not B Then
            Rs.Close
            Exit Function
        End If
        
        'Siguiente
        pb1.Value = pb1.Value + 1
        Cont = Cont + 1
        Rs.MoveNext
    Wend
    'Actualizamos la fecha de ultima amortizacion en paraemtros
    If opcion <> 3 Then
        Cad = "UPDATE paramamort SET ultfecha= '" & Format(Cad, FormatoFecha)
        Cad = Cad & "' WHERE codigo=1"
        Conn.Execute Cad
        Rs.Close
    Else
        'Estamos dando de baja o vendiendo un inmovilizado. Solo hay uno y hay k situarlo
        'en el primero
        Rs.Requery
        Rs.MoveFirst
    End If
    GeneraCalculoInmovilizado = True
    Exit Function
EGen:
    MuestraError Err.Number
End Function


'Para cancelar elto
Private Sub PonerCadenaLinea()
    Cad = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce"
    Cad = Cad & ", ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) VALUES ("
    Cad = Cad & RecuperaValor(ParametrosContabiliza, 4) & ",'"
    Cad = Cad & Format(Text4(0).Text, FormatoFecha)
    Cad = Cad & "'," & Mc.Contador & ","
End Sub


Private Function CargarDatosInmov() As Boolean
On Error GoTo ECar
    CargarDatosInmov = False
    Cad = "Select * from inmovele where codinmov =" & Text6(0).Text & " for update"
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Rs.EOF Then
        MsgBox "Error leyendo datos inmovilizado: " & Text6(0).Text, vbExclamation
        Rs.Close
    Else
        CargarDatosInmov = True
    End If
    Exit Function
ECar:
    MuestraError Err.Number, "Cargar datos inmovilizado"
End Function


Private Function CancelarCuentaElemento() As Boolean
Dim AUx As String

    On Error GoTo ECancelarCuentaElemento
    CancelarCuentaElemento = False
    If Not CargarDatosInmov Then Exit Function
    
    
    If Rs!repartos = 1 Then
     
        MsgBox "Error: REPARTOS incorrectos", vbExclamation
        Exit Function
    Else
        '---------------------------
        'NO tiene reparto de gastos
        '---------------------------
        PonerCadenaLinea
        Cad = Cad & Cont & ",'" & Rs!codmact3 & "','" & Format(Rs!Codinmov, "000000") & "',"
        Cad = Cad & RecuperaValor(ParametrosContabiliza, 2)   'Concepto DEBE
        Cad = Cad & ",'" & DevNombreSQL(Rs!nominmov)
        AUx = TransformaComasPuntos(CStr(Rs!amortacu))
        Cad = Cad & "'," & AUx & ",NULL" & ","     'AUX tiene el importe del inmovilizado
        AUx = "NULL"
        If Not IsNull(Rs!codccost) Then
            If HayKHabilitarCentroCoste(Rs!codmact3) Then AUx = "'" & Rs!codccost & "'"
        End If
        Cad = Cad & AUx
        Cad = Cad & ",'" & Rs!codmact1 & "','CONTAI',0)"
        Conn.Execute Cad
        Cont = Cont + 1
        
        
        '------------------------------------------------------------------------
        'Cancelacion del elemento
        PonerCadenaLinea
        

        

        Importe = Rs!valoradq - Rs!amortacu

        'La diferencia, si la hubiere se va a las perd/ganan de inmobilizado
        
        If Importe > 0 Then
            PonerCadenaLinea
        
            Cad = Cad & Cont & ",'" & txtCta(0).Text & "','" & Format(Rs!Codinmov, "000000") & "',"
            Cad = Cad & RecuperaValor(ParametrosContabiliza, 2)   'Concepto DEBE
            Cad = Cad & ",'" & DevNombreSQL(Rs!nominmov)
            AUx = TransformaComasPuntos(CStr(Importe))
            Cad = Cad & "'," & AUx & ",NULL" & ","
            
            AUx = "NULL"
            If Not IsNull(Rs!codccost) Then
                If HayKHabilitarCentroCoste(txtCta(0).Text) Then AUx = "'" & Rs!codccost & "'"
            End If
            
            Cad = Cad & AUx
            Cad = Cad & ",'" & Rs!codmact3 & "','CONTAI',0)"
            Conn.Execute Cad
            Cont = Cont + 1
        End If





        
        PonerCadenaLinea
        Cad = Cad & Cont & ",'" & Rs!codmact1 & "','" & Format(Rs!Codinmov, "000000") & "',"
        Cad = Cad & RecuperaValor(ParametrosContabiliza, 2)   'Concepto DEBE
        Cad = Cad & ",'" & DevNombreSQL(Rs!nominmov)
        AUx = TransformaComasPuntos(CStr(Rs!valoradq))
        Cad = Cad & "',NULL," & AUx & ","    'AUX tiene el importe del inmovilizado
        
        AUx = "NULL"
        If Not IsNull(Rs!codccost) Then
            If HayKHabilitarCentroCoste(Rs!codmact1) Then AUx = "'" & Rs!codccost & "'"
        End If
        
        Cad = Cad & AUx
        If Importe = 0 Then
            'SI QUE CANCELO en la ctapartida la cuenta
            Cad = Cad & ",'" & Rs!codmact3 & "'"
        Else
            Cad = Cad & ",NULL"
        End If
        Cad = Cad & ",'CONTAI',0)"
        Conn.Execute Cad
        Cont = Cont + 1
        
        
        
        
    End If
    

    'Si es venta tengo que restarle el valor de la venta

'    Importe = RS!valoradq - RS!amortacu - Importe
'
'    If Importe <> 0 Then
'       'No esta amortizado
'        PonerCadenaLinea
'        Cad = Cad & Cont & ",'" & txtCta(0).Text & "','" & Format(RS!Codinmov, "000000") & "',"
'        Cad = Cad & RecuperaValor(ParametrosContabiliza, 2)   'Concepto DEBE
'        Cad = Cad & ",'" & DevNombreSQL(RS!nominmov) & "',"
'        Aux = TransformaComasPuntos(CStr(Abs(Importe)))
'        'Si importe negativo va al debe, si no al haber
'        If Importe > 0 Then
'            Cad = Cad & Aux & ",NULL"
'        Else
'            Cad = Cad & "NULL," & Aux
'        End If
'
'        Aux = "NULL"
'        If Not IsNull(RS!codccost) Then
'            If HayKHabilitarCentroCoste(txtCta(0).Text) Then Aux = "'" & RS!codccost & "'"
'        End If
'
'        Cad = Cad & "," & Aux & ",NULL,'CONTAI',0)"
'        Conn.Execute Cad
'    End If
    Rs.Close
    
    
    
    
'   Es para baja
    If Option1(1).Value Then
        Cad = "UPDATE inmovele SET fecventa = '" & Format(Text4(0).Text, FormatoFecha)
        Cad = Cad & "', situacio =3  "
        Cad = Cad & " Where Codinmov = " & Text6(0).Text
        Conn.Execute Cad
    End If
    CancelarCuentaElemento = True
ECancelarCuentaElemento:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cancelar cuenta elto."
    Set Rs = Nothing
End Function


Private Function GeneracabeceraApunte(vTipo As Byte) As Boolean
Dim Fecha As Date
On Error GoTo EGeneracabeceraApunte
        GeneracabeceraApunte = False
        Cad = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES ("
        Cad = Cad & RecuperaValor(ParametrosContabiliza, 4) & ",'"
        If opcion = 3 Then
            Fecha = CDate(Text4(0).Text)
        Else
            Fecha = CDate(txtFecAmo.Text)
        End If
        Cad = Cad & Format(Fecha, FormatoFecha)
        Cad = Cad & "'," & Mc.Contador
        Cad = Cad & ",0,null,'"
        'Segun sea VENTA, BAJA, o calculo de inmovilizado pondremos una cosa u otra
        Select Case vTipo
        Case 0, 1
            'VENTA
            If vTipo = 0 Then
                Cad = Cad & "Venta de "
            Else
                Cad = Cad & "Baja de "
            End If
            Cad = Cad & DevNombreSQL(Rs!nominmov)
        Case Else
            Cad = Cad & "Amortización: " & Fecha
        End Select
        Cad = Cad & "')"
        Conn.Execute Cad
        GeneracabeceraApunte = True
        Exit Function
EGeneracabeceraApunte:
     MuestraError Err.Number, "Genera cabecera Apunte"
     Set Rs = Nothing
End Function


Private Function VentaElemento() As Boolean
Dim AUx As String
Dim RI As Recordset
Dim ImporteTotal As Currency
Dim mZ As Contadores
Dim TotalFactura As Currency

On Error GoTo EVentaElemento
    VentaElemento = False
        
        TipoIva = DevuelveDesdeBD("codiva", "paramamort", "codigo", "1", "N")
        If TipoIva = "" Then
            MsgBox "Error en el tipo de iva.", vbExclamation
            Exit Function
        End If
        Set RI = New ADODB.Recordset
        RI.Open "Select * from tiposiva WHERE codigiva=" & TipoIva, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If RI.EOF Then
            MsgBox "Error leyendo valores del IVA: " & TipoIva, vbExclamation
            RI.Close
            Exit Function
        End If
        
        'Conseguimos el contador para la factura
        Set mZ = New Contadores
        Cont = FechaCorrecta2(Text4(0).Text)
        If mZ.ConseguirContador("Z", (Cont = 0), True) = 1 Then
            RI.Close
            Rs.Close
            MsgBox "Error contador facuras inmovilizado", vbExclamation
            Set mZ = Nothing
            Exit Function
        End If
        
        
    
        
        
        'CABECEREA DE FACTURA ---
        'Genereamos la cabecera de factura
        Cad = "INSERT INTO cabfact (numserie, codfaccl, fecfaccl, codmacta, anofaccl, confaccl,"
        Cad = Cad & " pi1faccl, ba1faccl, ti1faccl, totfaccl, tp1faccl,fecliqcl) VALUES ("
        'Ejemplo:  'A', 111111112, '2022-02-02', '1', 2002, 'VENTA elto 1',
     
        AUx = "'" & mZ.TipoContador & "'," & mZ.Contador & ",'" & Format(Text4(0).Text, FormatoFecha) & "','"
        AUx = AUx & txtCta(1).Text & "'," & Year(Text4(0).Text) & ",'" & DevNombreSQL(Rs!nominmov) & "',"
        'Pocentaje iva, imponible tal  ytal
        Importe = RI!porceiva
        
        'Para la facutra impresa   Fecha,Numfac,desc, %IVA, total IVA
        CadenaDesdeOtroForm = Text4(0).Text & "|"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & mZ.TipoContador & Format(mZ.Contador, "0000000000") & "|"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & DevNombreSQL(Rs!nominmov) & "|"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Format(Importe, FormatoImporte) & "|"
        
        '-----------
        ImporteTotal = ImporteFormateado(Text5.Text)
        AUx = AUx & TransformaComasPuntos(CStr(Importe)) & "," & TransformaComasPuntos(ImporteSinFormato(Text5.Text)) & ","
        'Total iva
        Importe = (Importe * ImporteTotal) / 100
        Importe = Round(Importe, 2)
        AUx = AUx & TransformaComasPuntos(CStr(Importe)) & ","
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Format(Importe, FormatoImporte) & "|" 'TOTALIVA
        
        'Total factura
        Importe = Importe + ImporteTotal
        AUx = AUx & TransformaComasPuntos(CStr(Importe)) & "," & RI!codigiva
        TotalFactura = Importe   'Para la tesoreria
        
        
        'Fecha liquidacion
        AUx = AUx & ",'" & Format(Format(Text4(0).Text, FormatoFecha), FormatoFecha) & "')"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Format(Importe, FormatoImporte) & "|" 'TOTAL FAC
        
        Conn.Execute Cad & AUx
        
        ImporteTotal = ImporteFormateado(Text5.Text)
        Cont = 1
        
        '-------------------------------------------------------
        '-------------------------------------------------------
        'lINEAS . comun
        Cad = "INSERT INTO linfact (numserie, codfaccl, anofaccl, numlinea, codtbase, impbascl,"
        Cad = Cad & "codccost) VALUES ('" & mZ.TipoContador & "'," & mZ.Contador & "," & Year(Text4(0).Text) & ","
        
        
        'Modificacion de 26 Abril 2004
        '--------------------------------
        'Estos apuntes van en la cancelacion del elemento
        'Amortizacion acumulada DEBE
        
        'Generara 2 lineas de factura
        Importe = Rs!valoradq - Rs!amortacu    'Ahora importe tiene: pendiente de anmortizar
        If Importe > 0 Then
            'Cancelamos la amortizacion del elmento
        
        
            'Elemento
            AUx = TransformaComasPuntos(CStr(Importe))
            AUx = Cont & ",'" & Rs!codmact3 & "'," & AUx & ","
            TipoIva = "NULL"
            If vParam.autocoste Then
                If HayKHabilitarCentroCoste(Rs!codmact3) Then TipoIva = "'" & Me.txtCodCCost(2).Text & "'"
    
            End If
            Cont = Cont + 1
            Conn.Execute Cad & AUx & TipoIva & ")"
            
            
            'Ahora como he generado este apunte...
            'EL elto queda totalmente amortizado
            AUx = "UPDATE inmovele set amortacu=valoradq Where codinmov =" & Text6(0).Text
            Conn.Execute AUx
            
        Else
            Importe = 0 'Por si acaso estuviera mal. El importe acumulado NO puede ser mayor que el valoradq
        End If
            
        'Ganancias / perdidas de la venta
        Importe = ImporteTotal - Importe
        
        'Elemento
        AUx = TransformaComasPuntos(CStr(Importe))
        AUx = Cont & ",'" & txtCta(0).Text & "'," & AUx & ","
        TipoIva = "NULL"
        If vParam.autocoste Then
            If HayKHabilitarCentroCoste(txtCta(0).Text) Then TipoIva = "'" & Me.txtCodCCost(2).Text & "'"

        End If
        Cont = Cont + 1
        Conn.Execute Cad & AUx & TipoIva & ")"
        
        
        
        
        
        
        
        
        Rs.Close
        RI.Close
        
        
        'Si teiene tesoreria genero el cobro
        If vEmpresa.TieneTesoreria Then
            'Generamos el cobro
            Cad = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci,"
            Cad = Cad & "impvenci, ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, "
            Cad = Cad & "impcobro, emitdocum, recedocu, contdocu, text33csb, text41csb,"
            Cad = Cad & "text42csb, text43csb, text51csb, text52csb, text53csb, text61csb,"
            Cad = Cad & "text62csb, text63csb, text71csb, text72csb, text73csb, text81csb, text82csb,text83csb,"
            Cad = Cad & "ultimareclamacion, agente, departamento, codrem, anyorem, siturem, gastos, Devuelto, "
            Cad = Cad & "situacionjuri, noremesar, obs, transfer) VALUES ('"
        
        
            'Los datos del cobro                                                                                codmacta
            Cad = Cad & mZ.TipoContador & "'," & mZ.Contador & ",'" & Format(Text4(0).Text, FormatoFecha) & "',1,'" & txtCta(1).Text
            Cad = Cad & "'," & Text8(0).Text & ",'" & Format(Text4(1).Text, FormatoFecha) & "'," & TransformaComasPuntos(CStr(TotalFactura))
            Cad = Cad & ",'" & txtCta(3).Text & "',"
            
            
            AUx = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", Text8(0).Text, "N")
            i = Val(AUx)
            AUx = "NULL,NULL,NULL,NULL,"
            AUX2 = "NULL"
            If i = 4 Then 'RECIBO BANCARIO
                Rs.Open "Select entidad,oficina,cc,cuentaba from cuentas where codmacta = '" & txtCta(1).Text & "'", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not Rs.EOF Then
                    Cont = 2
                    For i = 0 To 3
                        If Not IsNull(Rs.Fields(CInt(i))) Then Cont = 3
                    Next i
                    If Cont = 3 Then
                        'AL MENOS HAY UN DATO
                        AUx = ""
                        AUX2 = ""
                        For i = 0 To 3
                            If IsNull(Rs.Fields(CInt(i))) Then
                                AUx = AUx & "NULL"
                                AUX2 = AUX2 & " - "
                            Else
                                AUx = AUx & Rs.Fields(CInt(i))
                                AUX2 = AUX2 & Rs.Fields(CInt(i))
                            End If
                            AUx = AUx & ","
                            AUX2 = AUX2 & " "
                        Next i
                        AUX2 = "'" & AUX2 & "'"
                    End If
                End If
                Rs.Close
            End If
            Cad = Cad & AUx
            
            'ctabna2, fecultco,impcobro, emitdocum, recedocu, contdocu,
            Cad = Cad & "NULL,NULL,NULL,0,0,0,"
            'Textos
            'Solo el 1 y el dos
            Cad = Cad & "'Factura Nº: " & mZ.TipoContador & " / " & mZ.Contador & "','Vto a fecha: " & Text4(1).Text & "',"
                'el resto de textos a NULL
            Cad = Cad & "NULL, NULL, NULL, NULL, NULL, NULL,"
            Cad = Cad & "NULL, NULL, NULL, NULL, NULL, NULL, NULL,NULL,"
            'ultimareclamacion, agente, departamento, codrem, anyorem, siturem, gastos, Devuelto,
            Cad = Cad & "NULL," & Text8(2).Text & ",NULL,NULL,NULL,NULL,NULL,NULL,"
            'situacionjuri, noremesar, obs, transfer
            Cad = Cad & "0,0,NULL,NULL)"
            Conn.Execute Cad
            
            
            
            
            
            'Tambien , para la factura meteremos en la tabla tesoreria comun
            'los datos del vto
             Cad = "DELETE from ztesoreriacomun where codusu =" & vUsu.Codigo
             EjecutaSQL Cad
             Cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2"
             Cad = Cad & " ,importe1,  fecha1) values (" & vUsu.Codigo & ",1,'"
             Cad = Cad & Text8(1).Text & "'," & AUX2 & "," & TransformaComasPuntos(CStr(Importe)) & ",'"
             Cad = Cad & Format(Text4(1).Text, FormatoFecha) & "')"
             EjecutaSQL Cad
        End If
        
        

        
        
        'Ahora hay k poner el elemento a vendido, con el importe de venta y la fecha de venta
        Cad = "UPDATE inmovele SET fecventa = '" & Format(Text4(0).Text, FormatoFecha)
        Cad = Cad & "', situacio =2 , impventa="
        Cad = Cad & TransformaComasPuntos(ImporteSinFormato(Text5.Text)) & " WHERE codinmov =" & Text6(0).Text
        Conn.Execute Cad
        
        VentaElemento = True
EVentaElemento:
    If Err.Number <> 0 Then MuestraError Err.Number, "Venta Elemento"
    Set Rs = Nothing
    Set RI = Nothing
End Function



'Devuelve TRUE si esta activo
Private Function HayQueAmortizar() As Boolean
HayQueAmortizar = False
Cad = DevuelveDesdeBD("situacio", "inmovele", "codinmov", Text6(0).Text, "N")
If Cad <> "" Then
    If Cad = "1" Then HayQueAmortizar = True
End If
End Function



'Datos listado ENRE FECHAS
Private Function GeneraDatosListado() As Boolean
Dim RT As ADODB.Recordset
Dim Veces As Integer

    On Error GoTo EGeneraDatosListado
    GeneraDatosListado = False
    
    Cad = "DELETE from tmpSimula WHERE codusu = " & vUsu.Codigo
    Conn.Execute Cad
    
    Cad = "DELETE FROM Usuarios.zentrefechas where codusu=" & vUsu.Codigo
    Conn.Execute Cad
    
    
    
    'Cargamos los eltos k podrian ser mosr a ser mostrados por tener movimientos entre las fechas indicada
    Cad = "INSERT INTO tmpsimula "
    Cad = Cad & " SELECT  " & vUsu.Codigo & ",codinmov,0,sum(imporinm) FROM  inmovele_his WHERE "
    'las fechas
    Cad = Cad & " fechainm >= '" & Format(txtfec(0).Text, FormatoFecha)
    Cad = Cad & "' AND fechainm <= '" & Format(txtfec(1).Text, FormatoFecha)
    Cad = Cad & "' GROUP BY codinmov"
    Conn.Execute Cad
    
    
    'Nuevo Marzo 2010
    'Para los elementos que por algun motivo no tienen valor en inmovele_his pero si que deberian salir
    
    DivMes = 1 'Contador
    Cad = "INSERT INTO Usuarios.zentrefechas (codusu, codigo, codccost, nomccost, conconam, nomconam, "
    Cad = Cad & "codinmov, nominmov, fechaadq, valoradq, amortacu, fecventa, impventa,impperiodo) VALUES (" & vUsu.Codigo & ","
   
    Set RT = New ADODB.Recordset
    Set Rs = New ADODB.Recordset
    
    'LO VOY A REALIZAR DOS VECES.
    '   La primera: la que habia.
    '   Segunda. Eltos que no estan en inmovele_his(por algun motivo) pero tienen que salir
    For Veces = 1 To 2
    
    
        If Veces = 1 Then
            RT.Open "Select * from tmpSimula where codusu =" & vUsu.Codigo, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
        Else
            'La segunda vez es otra historia. Cogere aquellos eltos quecon fecha adquisicion .
            'no esten en inmovele_his pero si que si que puedan salir
            Cont = InStr(1, ParametrosContabiliza, "WHERE")
            AUX2 = Mid(ParametrosContabiliza, Cont)
            
            AUX2 = Replace(AUX2, "codconam", "conconam")
            
            AUX2 = Mid(AUX2, 1, Len(AUX2) - 11) 'LE quito el ultimo codinmov=
            AUX2 = AUX2 & "  fechaadq >= '" & Format(txtfec(0).Text, FormatoFecha)
            AUX2 = AUX2 & "' AND fechaadq <= '" & Format(txtfec(1).Text, FormatoFecha)
            AUX2 = AUX2 & "' AND not codinmov IN (select distinct(codinmov) from inmovele_his)"
            AUX2 = " SELECT " & vUsu.Codigo & ",codinmov,0,0 totalamor FROM inmovele " & AUX2
            RT.Open AUX2, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            AUX2 = ""
        End If
        
        While Not RT.EOF
        
            Rs.Open ParametrosContabiliza & RT.Fields(1), Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
            If Not Rs.EOF Then
                If IsNull(Rs!codccost) Then
                    'AUX2 = DivMes & ",'" & DBLet(RS!codccost, "T") & "','" & DevNombreSQL(DBLet(RS!nomccost, "T")) & "',"
                    AUX2 = DivMes & ",'','SIN CENTRO DE COSTE',"
                Else
                    AUX2 = DivMes & ",'" & Rs!codccost & "','" & DevNombreSQL(Rs!nomccost) & "',"
                End If
                AUX2 = AUX2 & Rs!codconam & ",'" & Rs!nomconam & "'," & Rs!Codinmov & ",'"
                AUX2 = AUX2 & DevNombreSQL(Rs!nominmov) & "','" & Format(Rs!fechaadq, "dd/mm/yyyy") & "',"
                TipoIva = TransformaComasPuntos(CStr(Rs!valoradq))
                AUX2 = AUX2 & TipoIva & ","
                TipoIva = TransformaComasPuntos(CStr(Rs!amortacu))
                AUX2 = AUX2 & TipoIva & ","
                If IsNull(Rs!fecventa) Then
                    AUX2 = AUX2 & "NULL"
                Else
                    AUX2 = AUX2 & "'" & Format(Rs!fecventa, FormatoFecha) & "'"
                End If
                AUX2 = AUX2 & ","
                If IsNull(Rs!impventa) Then
                    AUX2 = AUX2 & "NULL"
                Else
                    TipoIva = TransformaComasPuntos(CStr(Rs!impventa))
                    AUX2 = AUX2 & TipoIva
                End If
                AUX2 = AUX2 & ","
                
                'El importe del peirodo
                If IsNull(RT!totalamor) Then
                    AUX2 = AUX2 & "NULL"
                Else
                    TipoIva = TransformaComasPuntos(CStr(RT!totalamor))
                    AUX2 = AUX2 & TipoIva
                End If
                
                AUX2 = AUX2 & ")"
                Conn.Execute Cad & AUX2
                
                DivMes = DivMes + 1

            End If
            Rs.Close
            RT.MoveNext
        Wend
        RT.Close
    Next Veces '** LO hago dos veces. La de antes y la de busqueda de eltos  que no tienene en inmovele_his
        
    If DivMes = 1 Then
        MsgBox "Ningún dato con esos valores", vbExclamation
    Else
        GeneraDatosListado = True
    End If
EGeneraDatosListado:
        If Err.Number <> 0 Then MuestraError Err.Number
        Set Rs = Nothing
End Function


'///////////////////////////////////////////////////////////
'
'   Este procedimento utilizad dos tablas ya creadas que son
'   en USUARIOS  z347 y z347carta
Private Sub EmiteFacturaVentaInmmovilizado()

On Error GoTo EEmiteFacturaVentaInmmovilizado
    Cad = "DELETE FROM Usuarios.z347  WHERE codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    'Los datos del encabezado
    CargaEncabezadoCarta 1
    DivMes = 0
    
    
    Set Rs = New ADODB.Recordset
    Cad = "Select * from Cuentas where codmacta='" & txtCta(1).Text & "'"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = "INSERT INTO Usuarios.z347 (codusu, cliprov, nif, importe, razosoci, dirdatos, codposta, despobla) VALUES (" & vUsu.Codigo
    If Rs.EOF Then
        Cad = Cad & "," & DivMes & ",'nif'," & TransformaComasPuntos(ImporteSinFormato(Text5.Text)) & ",'" & DevNombreSQL(txtDescta(1).Text) & "','Direccion','codpos','Poblacion')"
    Else
        Cad = Cad & "," & DivMes & ",'"
        Cad = Cad & DBLet(Rs!nifdatos) & "'," & TransformaComasPuntos(ImporteSinFormato(Text5.Text)) & ",'" & DevNombreSQL(txtDescta(1).Text) & "','"
        Cad = Cad & DevNombreSQL(DBLet(Rs!dirdatos)) & "','" & DBLet(Rs!codposta) & "','"
        Cad = Cad & DevNombreSQL(DBLet(Rs!desPobla)) & "')"
    End If
    Conn.Execute Cad
    
    Exit Sub
EEmiteFacturaVentaInmmovilizado:
    MuestraError Err.Number, "Generando factura"
End Sub



'-----------------------------------------------------------------------------
'
'
'       Deshacer ultima amortizacion
'
Private Sub CargarDatosAmortizacion()

    'Obtengo la ultima fecha a partir de la amortizacion y ultima fecha amortizada
    UltAmor = "01/01/1901"
    Cad = DevuelveDesdeBD("ultfecha", "paramamort", "codigo", "1", "N")
    If Cad <> "" Then UltAmor = CDate(Cad)
    
    If ObtenerparametrosAmortizacion(DivMes, UltAmor, ParametrosContabiliza) Then
        AUX2 = Format(UltAmor, "dd/mm/yyyy")
        B = True
    Else
        B = False
        AUX2 = "### ERROR obten. fecha ###"
    End If
    Cad = "Fecha última amortización:"
    Cad = Cad & Space(20) & AUX2
    Label13(6).Caption = Cad
              
    'Habilitamos o no el boton de deshacer
    cmdDeshaz(1).Enabled = B
End Sub



Private Sub DeshacerUltimaAmortizacion()

    'Constara de varios pasos
    '-------------------------------------------------------------------------------
    'Algunas comprobaciones. Ejercicios contables, que nos se ha vendido ni dado de baja....
    If Not Datosok_Deshacer Then Exit Sub
    


    'Deshacemos en inmovele_his y en inmovele. En los inmovilizados propiamente dicho
    'Transaccionamos esta accion
    PreparaBloquear
    Conn.BeginTrans
    
    If EliminarAmortizacion Then
        B = True
        Conn.CommitTrans
        Me.cmdDeshaz(1).Enabled = False
        'Grabamos el LOG
        Cad = "Fecha ult amortizacion: " & UltAmor
        vLog.Insertar 14, vUsu, Cad

    Else
        B = False
        Conn.RollbackTrans
    End If
    TerminaBloquear
    'Si da error nos piramos
    If Not B Then MsgBox "Se han producido errores.", vbExclamation
        
   
    



End Sub



Private Function EliminarAmortizacion() As Boolean
Dim Valor As Currency
Dim F As Date
    On Error GoTo EEliminarAmortizacion

    EliminarAmortizacion = False
    
    Label13(6).Caption = "Comprobar datos"
    Me.Refresh
    DoEvents
    
    'Compreubo cuantos hay. Para que no haya errores
    Cad = "Select count(*) from inmovele_his where fechainm = '" & Format(UltAmor, FormatoFecha) & "'"
    Cont = 0
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then Cont = DBLet(Rs.Fields(0), "N")
    Rs.Close
    AUX2 = CStr(Cont)
    
    If Cont = 0 Then
        MsgBox "Error: NUmero de registos de hcoinmovilizado con fecha " & UltAmor & " es cero", vbExclamation
        Exit Function
    End If
    
    'Abro el rs para actualizar
    Cad = "select l.codinmov,imporinm,amortacu,valoradq,nominmov from inmovele_his l,inmovele where l.codinmov=inmovele.codinmov "
    Cad = Cad & " and fechainm = '" & Format(UltAmor, FormatoFecha) & "'"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    While Not Rs.EOF
        Label13(6).Caption = Rs!Codinmov & " " & Rs!nominmov
        Label13(6).Refresh
        
        'Para cada elemento le sumo lo que a amortizado
        Importe = DBLet(Rs!amortacu, "N")
        Importe = Importe - Rs!imporinm
        
        'Control auxiliar
        If Importe < 0 Then Importe = 0
        
        'Creo SQL update
        Cad = "UPDATE inmovele set amortacu=" & TransformaComasPuntos(CStr(Importe))
        Cad = Cad & ", situacio= 1"
        Cad = Cad & " WHERE codinmov=" & Rs!Codinmov
        
        'Muevo al siguiente
        Rs.MoveNext
        'Updateo
        Conn.Execute Cad
        'cont++
        Cont = Cont + 1
        
        
    Wend
    Rs.Close
    
    
    If Cont <> Val(AUX2) Then
        'ERROR. Iban a ser val(aux2)  registros y solo se han preocesado cont
        Cad = "Registros del count(*)= " & AUX2 & vbCrLf & "Registros procesados= " & Cont
        Cad = "Error. " & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Exit Function
    End If
    
    Label13(6).Caption = "Restaurando datos situacion anterior"
    Label13(6).Refresh
    
    
    'Borramos todos los datos de inmovele_his con esta fecha
    Cad = "DELETE from inmovele_his where fechainm = '" & Format(UltAmor, FormatoFecha) & "'"
    Conn.Execute Cad
    
    'ACtualizamos la fecha de ultamor
    '--------------------------------
    AUX2 = "tipoamor"
    Cad = DevuelveDesdeBD("intcont", "paramamort", "codigo", "1", "N", AUX2)
    Contabiliza = (Cad = 1)
    DivMes = Val(AUX2)
    Select Case DivMes
        Case 2
            'Semestral
            i = 6
            'Siempre es la ultima fecha de mes
        Case 3
            'Trimestral
            i = 3
        Case 4
            'Mensual
            i = 1
        Case Else
            'Anual
            i = 12
    End Select
    F = DateAdd("m", -i, UltAmor)
    i = DiasMes(CByte(Month(UltAmor)), Year(UltAmor))
    If i = Day(UltAmor) Then
        'Es ultimo dia mes
        'Leugo la fecha sera el ultimo dia de mes
        i = DiasMes(CByte(Month(F)), Year(F))
        F = CDate(i & "/" & Month(F) & "/" & Year(F))
        
    End If
    Cad = Format(F, FormatoFecha)
    Cad = "UPDATE paramamort set ultfecha='" & Cad & "'"
    Conn.Execute Cad
    
    

    If Not Contabiliza Then
        'Proceso finalizado con exito. No busco el asiento
        AUX2 = "Proceso finalizado con exito"
    Else
        'Si contabiliza tratamos de indicarle cual fue el asiento generado.
        'Busco el cabapu que cuadra con fechaent='uktamor' y en observaciones lleva amortizacion
        Cad = "cabapu where fechaent = '" & Format(UltAmor, FormatoFecha) & "' AND obsdiari like '%mortiza%'"
        Cont = 0
        'En introduccion
        AUX2 = "Select * from " & Cad
        Rs.Open AUX2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            'LO HE ENCONTRADO
            Cont = 1
            Cad = "Asiento: " & Rs!NumAsien & "      Diario: " & Rs!NumDiari & "      Fecha: " & Rs!FechaEnt & vbCrLf & "Observaciones: " & DBMemo(Rs!obsdiari)
        End If
        Rs.Close
        
        If Cont = 0 Then
            'Si no esta, pruebo en el hco
            AUX2 = "Select * from h" & Cad
            Rs.Open AUX2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                'LO HE ENCONTRADO
                Cont = 2
                Cad = "Asiento: " & Rs!NumAsien & "     Diario: " & Rs!NumDiari & "      Fecha: " & Rs!FechaEnt & vbCrLf & "Observaciones: " & DBMemo(Rs!obsdiari)
            End If
            Rs.Close
        End If
        
        'Si cont>0 entonces SI que lo ha encontrado
        
        If Cont > 0 Then
            If Cont = 1 Then
                AUX2 = "la introducción"
            Else
                AUX2 = "el histórico"
            End If
            Cad = "El asiento se encuentra en " & AUX2 & " de apuntes:" & vbCrLf & Cad
        Else
            Cad = "El asiento NO ha sido encontrado"
        End If
        AUX2 = "Proceso finalizado con exito." & vbCrLf & vbCrLf & vbCrLf & Cad
    End If
    
    MsgBox AUX2, vbInformation


    Label13(6).Caption = AUX2

    EliminarAmortizacion = True
    Exit Function
EEliminarAmortizacion:
    MuestraError Err.Number, Err.Description
End Function


Private Function Datosok_Deshacer() As Boolean
    On Error GoTo Edatosok_deshacer
    Datosok_Deshacer = False


    varFecOk = FechaCorrecta2(UltAmor)
    Cad = ""
    If varFecOk > 1 Then
        If varFecOk = 2 Then
            Cad = Mid(varTxtFec, 6)
        Else
            Cad = " fuera de ejercicios. "
        End If
    End If
    If Cad <> "" Then
        Cad = "Fecha última amortizacion " & LCase(Cad)
        MsgBox Cad, vbExclamation
        Exit Function
    End If

    Cad = "select distinct(inmovele_his.codinmov) from inmovele_his, inmovele where inmovele_his.codinmov=inmovele.codinmov and"
    Cad = Cad & " fechainm>='" & Format(UltAmor, FormatoFecha) & "'  and fecventa >='" & Format(UltAmor, FormatoFecha) & "'"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    While Not Rs.EOF
        Cont = Cont + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    If Cont > 0 Then
        Cad = "Hay " & Cont & " elemento(s) de inmovilizado que están en el hco inmovilizado  y han sido vendidos o dados de baja"
        MsgBox Cad, vbExclamation
        Exit Function
    End If
    
    
    Cad = "select distinct(inmovele_his.codinmov) from inmovele_his where  fechainm > '" & Format(UltAmor, FormatoFecha) & "'"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    While Not Rs.EOF
        Cont = Cont + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    If Cont > 0 Then
        Cad = "Hay " & Cont & " elemento(s) de inmovilizado que están en el hco inmovilizado."
        MsgBox Cad, vbExclamation
        Exit Function
    End If

    Datosok_Deshacer = True
    Exit Function
Edatosok_deshacer:
    MuestraError Err.Number, Err.Description
End Function
