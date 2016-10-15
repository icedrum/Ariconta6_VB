VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#17.2#0"; "Codejock.SkinFramework.v17.2.0.ocx"
Begin VB.Form frmPpalOLD 
   Caption         =   "Form1"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11640
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmPpal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9720
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListView3 
      Height          =   1695
      Left            =   7470
      TabIndex        =   7
      Top             =   6810
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2990
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageListPPal48 
      Left            =   2250
      Top             =   4860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   7858
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
   End
   Begin VB.Frame FrameSeparador 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      DragMode        =   1  'Automatic
      Height          =   3015
      Left            =   0
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      Top             =   0
      Width           =   45
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5535
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9763
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList4"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1695
      Left            =   3030
      TabIndex        =   6
      Top             =   6810
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2990
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "usuario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Mensaje"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   390
      Top             =   6510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM 
      Left            =   1200
      Top             =   7140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN 
      Left            =   1200
      Top             =   7740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun16 
      Left            =   2010
      Top             =   7740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN16 
      Left            =   1980
      Top             =   7140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM16 
      Left            =   1980
      Top             =   6510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   390
      Top             =   7110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":712C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":98DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListDocumentos 
      Left            =   390
      Top             =   7740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":10140
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":113C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":13B74
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":15CAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":15FC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":193BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1AFCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1BDA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1CD18
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1DC90
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImaListBotoneras 
      Left            =   1470
      Top             =   5550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1EC2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2548F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2BCF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":32553
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":38DB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3F617
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":45E79
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4C6DB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   720
      Top             =   5970
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImaListBotoneras32 
      Left            =   510
      Top             =   5370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4D0ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5394F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5A1B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":60A13
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":67275
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6DAD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":74339
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImaListBotoneras32_BN 
      Left            =   1230
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7AB9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":813FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":87C5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8E4C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":94D23
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":9B585
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A1DE7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImaListBotoneras_BN 
      Left            =   2160
      Top             =   5550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483626
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A8649
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":AEEAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B570D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":BBF6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":C27D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":C9033
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":CF895
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D60F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":DC959
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E31BB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListPpal16 
      Left            =   2280
      Top             =   4170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   1665
      Left            =   7410
      TabIndex        =   9
      Top             =   6840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2937
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
   End
   Begin MSComctlLib.ImageList imgIcoForms 
      Left            =   390
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E3BCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E45DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E467A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListviews 
      Left            =   1230
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E508C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EB8EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EE0A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F3CC2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListComun 
      Left            =   2070
      Top             =   8550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun2 
      Left            =   3210
      Top             =   8550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":FA524
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":FB5B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":FC648
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":FD6DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":FE76C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":FF7FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":100890
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":101922
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1029B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":103A46
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":104AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":105B6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":106BFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":107C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":108D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":109DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":10AE44
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":10BED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":10CF68
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":10DFFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":10F08C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":11011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1111B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":112242
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1132D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":114366
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   210
      Top             =   6030
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblMsgApli 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   0
      Left            =   7470
      TabIndex        =   8
      Top             =   6540
      Width           =   1305
   End
   Begin VB.Label lblMsgUsu 
      Caption         =   "Empresas disponibles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   5
      Top             =   6540
      Width           =   3135
   End
   Begin VB.Image ImageLogo 
      Height          =   720
      Left            =   7800
      Picture         =   "frmPpal.frx":1153F8
      Top             =   0
      Width           =   1890
   End
   Begin VB.Label Label33 
      BackColor       =   &H0070532E&
      Caption         =   " AriCONTA 6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9195
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Height          =   3105
      Left            =   -120
      Top             =   840
      Width           =   10455
   End
   Begin VB.Label Label22 
      BackColor       =   &H0070532E&
      Height          =   690
      Left            =   7440
      TabIndex        =   4
      Top             =   -120
      Width           =   3135
   End
   Begin VB.Menu mnPopUp 
      Caption         =   "mnPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnPopUp1 
         Caption         =   "Editar"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnPopUp1 
         Caption         =   "Eliminar"
         Index           =   1
      End
      Begin VB.Menu mnPopUp1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnPopUp1 
         Caption         =   "Organizar"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmPpalOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nomempre As String  'Vendra con los parametrros,


Public UnaVez As Boolean
Dim Base
Dim AnchoListview As Integer

Dim PrimeraVez As Boolean


Dim EstadoAnterior As Byte

'Ajustar iconos
Dim IconoSeleccionado As Boolean

Private frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1


Private Sub Form_Activate()

    Screen.MousePointer = vbHourglass
    If UnaVez Then
        UnaVez = False
        CargaMenu "ariconta", Me.TreeView1
        CargaMenu "introcon", Me.TreeView2
        
        MenuComoEstaba Me.TreeView1, "ariconta"
        MenuComoEstaba Me.TreeView2, "introcon"
        
        ListView2.SmallIcons = Me.ImageList2
        ListView1.Icons = Me.ImageListPPal48
        
        CargaShortCuts 0
        
        BuscaEmpresas
        NumeroEmpresaMemorizar True

        SituarItemList ListView2

    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
        
'    Me.Icon = frmEntrada.Icon
    PrimeraVez = True
    
'    Me.Icon = LoadPicture(App.Path & "\AriConta redondo.ico")
    
    
    ImageList1.ImageHeight = 48
    ImageList1.ImageWidth = 48
    GetIconsFromLibrary App.Path & "\icoconppal.dll", 1, 48


    ImageList2.ImageHeight = 16
    ImageList2.ImageWidth = 16
    GetIconsFromLibrary App.Path & "\icoconppal.dll", 1, 16

    ImageListPPal48.ImageHeight = 48
    ImageListPPal48.ImageWidth = 48
    GetIconsFromLibrary App.Path & "\icoconppal2.dll", 8, 48


    ImageListPpal16.ImageHeight = 16
    ImageListPpal16.ImageWidth = 16
    GetIconsFromLibrary App.Path & "\icoconppal2.dll", 9, 16

'    Me.Icon = Me.ImageListPpal16.ListImages(2).Picture


    ImgListComun.ImageHeight = 24
    ImgListComun.ImageWidth = 24
    GetIconsFromLibrary App.Path & "\iconosconta.dll", 2, 24 'antes icolistcon
    
    '++
    imgListComun_BN.ImageHeight = 24
    imgListComun_BN.ImageWidth = 24
    GetIconsFromLibrary App.Path & "\iconosconta_BN.dll", 3, 24
    
    imgListComun_OM.ImageHeight = 24
    imgListComun_OM.ImageWidth = 24
    GetIconsFromLibrary App.Path & "\iconosconta_OM.dll", 4, 24
    
    imgListComun16.ImageHeight = 16
    imgListComun16.ImageWidth = 16
    GetIconsFromLibrary App.Path & "\iconosconta.dll", 5, 16
    
    GetIconsFromLibrary App.Path & "\iconosconta_BN.dll", 6, 16
    GetIconsFromLibrary App.Path & "\iconosconta_OM.dll", 7, 16
    '++

    
    ' sirve para calcular despues el width
    Base = 1290
    Base = Base + 550 '550 es lo k mide de alto la imagen de ariadna
   
    ListView1.Picture = LoadPicture(App.Path & "\fondo.dat")
    
    
    PonerCaption
       
    PonerDatosFormulario

    EstablecerSkin CInt(vUsu.Skin)
    
    EstadoAnterior = WindowState

End Sub



Private Sub PonerCaption()
        Caption = "AriCONTA 6    V-" & App.Major & "." & App.Minor & "." & App.Revision & "    usuario: " & vUsu.Nombre & "      Ejercicio: " & vParam.fechaini & " - " & vParam.fechafin
        Label33.Caption = "   " & vEmpresa.nomempre
End Sub



Private Sub Form_Resize()
    Dim X, Y As Integer
Dim v ''

    On Error GoTo eResize

    If WindowState = 1 Then
        EstadoAnterior = 1
        Exit Sub         ' ha pulsado minimizar
    End If
    
    If EstadoAnterior = 1 Then
        EstadoAnterior = Me.WindowState
        Exit Sub
    End If
    
    X = Me.Width
    Y = Me.Height
    If X < 5990 Then Me.Width = 5990
    If Y < 4100 Then Me.Height = 4100
    ImageLogo.Left = Me.Width - ImageLogo.Width - 240
    Label33.Left = 30
    X = Me.Height - Base
    

    TreeView1.Height = X
    X = X \ 6
    ListView1.Height = X * 4
    
    ListView2.top = ListView1.top + ListView1.Height + 500
    ListView2.Height = Me.Height - ListView2.top - 850
    ListView3.Height = Me.Height - ListView2.top - 850
    TreeView2.top = ListView2.top
    TreeView2.Height = ListView2.Height
    
    
    Y = Me.Width - 200
    Y = ((30 / 100) * Y)
    
    TreeView1.Left = 30
    TreeView1.Width = Y - 30
    
    'Separador
    Me.FrameSeparador.Left = Y + 15
    Me.FrameSeparador.top = TreeView1.top
    Me.FrameSeparador.Height = Me.TreeView1.Height
    
    ListView1.Left = Y + 60
    Me.ListView2.Left = Y + 60
    
    AnchoListview = Me.Width - 200 - Y - 30
    ListView1.Width = AnchoListview
    v = Me.ImageLogo.Left
    
    Label33.Width = v + 20
    Label33.Left = -15
    Label22.Left = Label33.Width - 120
    Label22.Width = Me.Width - Label22.Left
    Label22.top = 0
    
    ImageLogo.top = Label33.top
    ImageLogo.Width = 1890
    ImageLogo.Left = Me.Width - ImageLogo.Width - 120
    ImageLogo.Height = Label22.Height
    
    X = AnchoListview \ 3
    ListView2.Width = 2 * X
    
    Me.TreeView2.Left = Me.ListView2.Left + Me.ListView2.Width + 30
    TreeView2.Width = X
    
    'Dos listiview
    lblMsgUsu(0).top = ListView2.top - 340 '240
    lblMsgApli(0).top = ListView2.top - 340 '240
    
    'Left
    lblMsgUsu(0).Left = ListView2.Left + 60
    lblMsgApli(0).Left = TreeView2.Left + 60
    lblMsgApli(0).Visible = True
    
    
    Shape1.Width = Me.Width - Shape1.Left - 50
    Shape1.Height = Me.Height - Shape1.top - 50
    
    EstadoAnterior = Me.WindowState
        
    Exit Sub
    
eResize:
'    Caption = Now
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    FijarUltimoSkin False
'    FreeLibrary m_hMod: UnloadApp: End
    ActualizarExpansionMenus vUsu.Id, Me.TreeView1, "ariconta"
    ActualizarExpansionMenus vUsu.Id, Me.TreeView2, "intracon"
    
    ActualizarItems vUsu.Id, Me.ListView1, "ariconta"
    
    NumeroEmpresaMemorizar False
    On Error Resume Next
    Set Conn = Nothing
End Sub

Private Sub Image1_Click()
    LanzaVisorMimeDocumento Me.hWnd, "http://www.ariadnasw.com"
End Sub

Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub

    AbrirFormularios CLng(Mid(ListView1.SelectedItem.Key, 3))
End Sub

Private Sub AbrirFormularios(Accion As Long)

    Select Case Accion
        Case 101 ' empresa
            frmempresa.Show vbModal
        Case 102 ' parametros contabilidad
            If Not (vEmpresa Is Nothing) Then
                frmparametros.Show vbModal
            End If
        Case 103 ' parametros tesoreria
        Case 104 ' contadores
            Screen.MousePointer = vbHourglass
            If vUsu.Nivel = 0 Then frmContadores.Show vbModal
        Case 105 ' usuarios
            frmMantenusu.Show vbModal
        Case 106 ' informes
            frmCrystal.Show vbModal
        Case 107 ' crear nueva empresa
            If vUsu.Nivel > 1 Then Exit Sub
            
            frmCentroControl.Opcion = 2
            frmCentroControl.Show vbModal
        Case 108 'Configurar Balances
            Screen.MousePointer = vbHourglass
            frmColBalan.Show vbModal
        Case 201 ' plan contable
            Screen.MousePointer = vbHourglass
            frmColCtas.ConfigurarBalances = 0
            frmColCtas.DatosADevolverBusqueda = ""
            frmColCtas.Show vbModal
        Case 202 ' tipos de diario
            Screen.MousePointer = vbHourglass
            frmTiposDiario.Show vbModal
        Case 203 ' conceptos
            Screen.MousePointer = vbHourglass
            frmConceptos.Show vbModal
        Case 204 ' tipos de iva
            Screen.MousePointer = vbHourglass
            frmIVA.Show vbModal
        Case 205 ' tipos de pago
            Screen.MousePointer = vbHourglass
            frmTipoPago.Show vbModal
        Case 206 ' formas de pago
            Screen.MousePointer = vbHourglass
            frmFormaPago.Show vbModal
        Case 207 ' bancos
            Screen.MousePointer = vbHourglass
            frmBanco.Show vbModal
        Case 208 ' bic
            Screen.MousePointer = vbHourglass
            frmBic.Show vbModal
        Case 209 ' agentes
            Screen.MousePointer = vbHourglass
            frmAgentes.Show vbModal
        Case 210 ' departamentos
        Case 211 ' asientos predefinidos
            Screen.MousePointer = vbHourglass
            frmAsiPre.Show vbModal
        Case 212 ' cartas de reclamacion
            Screen.MousePointer = vbHourglass
            frmCartas.Show vbModal
        
        Case 301 ' asientos
            Screen.MousePointer = vbHourglass
            frmAsientosHco.ASIENTO = ""
            frmAsientosHco.DesdeNorma43 = 0
            frmAsientosHco.Show vbModal
        Case 303 ' extractos
            Screen.MousePointer = vbHourglass
            frmConExtr.EjerciciosCerrados = False
            frmConExtr.Cuenta = ""
            frmConExtr.Show vbModal
        Case 304 ' punteo
            Screen.MousePointer = vbHourglass
            frmPuntear.EjerciciosCerrados = False
            frmPuntear.Show vbModal
        Case 305 ' reemision de diarios
'            AbrirListado 6, False
        Case 306 ' sumas y saldos
            frmInfBalSumSal.Show vbModal
            
        Case 307 ' cuenta de explotacion
            frmInfCtaExplo.Show vbModal
            
        Case 308 ' balance de situacion
            frmInfBalances.Opcion = 0
            frmInfBalances.Show vbModal
            
        Case 309 ' perdidas y ganancias
            frmInfBalances.Opcion = 1
            frmInfBalances.Show vbModal

        Case 310 ' totales por concepto
            frmInfTotCtaCon.Show vbModal
        Case 311 ' evolucion de saldos
            frmInfEvolSal.Show vbModal
        Case 312 ' ratios y graficas
            frmInfRatios.Show vbModal
        Case 314 ' puntero extracto bancario
            frmPunteoBanco.Show vbModal
        Case 401 ' emitidas
            Screen.MousePointer = vbHourglass
            frmFacturasCli.Show vbModal
        Case 402 ' libro emitidas
            frmFacturasCliListado.Show vbModal
        Case 403 ' relacion clientes por cuenta
            frmFacturasCliCtaVtas.Show vbModal
        Case 404 ' recibidas
            Screen.MousePointer = vbHourglass
            frmFacturasPro.Show vbModal
        Case 405 ' libro recibidas
            frmFacturasProListado.Show vbModal
        Case 406 ' relacion proveedores por cuenta
            frmFacturasProCtaGastos.Show vbModal
        Case 407 ' liquidacion iva
'            AbrirListado 12, False
        Case 408 ' certificado iva
            frmModelo303.OpcionListado = 0
            frmModelo303.Show vbModal
        Case 409 ' modelo 340
            frmModelo340.Show vbModal
        Case 410 ' modelo 347
            frmModelo347.Show vbModal
        Case 411 ' modelo 349
            frmModelo349.Show vbModal
        Case 412 ' liquidacion de iva
            frmHcoLiqIVA.Show vbModal
        
        Case 502 ' conceptos
            Screen.MousePointer = vbHourglass
            frmInmoConceptos.Show vbModal
        Case 503 ' elementos
            frmInmoElto.DatosADevolverBusqueda = ""
            frmInmoElto.Show vbModal
        Case 505 ' estadistica
            frmInmoInfEst.Show vbModal
        Case 507 ' historico inmovilizado
            Screen.MousePointer = vbHourglass
            frmInmoHco.Show vbModal
        Case 508 ' simulacion
            frmInmoSimu.Show vbModal
        Case 509 ' calculo y contabilizacion
            frmInmoGenerar.Opcion = 2
            frmInmoGenerar.Show vbModal
        Case 510 ' deshacer amortizacion
            frmInmoDeshacer.Show vbModal
        Case 511 ' venta-baja inmmovilizado
            frmInmoVenta.Opcion = 3
            frmInmoVenta.Show vbModal
        Case 601 ' cartera de cobros
            frmTESCobros.Show vbModal
        Case 602 ' informe de cobros pendientes
            frmTESCobrosPdtesList.Show vbModal
        
        Case 603 ' impresion de recibos
            frmTESImpRecibo.Show vbModal
        Case 604 ' realizar cobro
            With frmTESRealizarCobros
                .ImporteGastosTarjeta_ = 0 '--Importe
                '--.vSQL = SQL
                .Regresar = False
                .Cobros = True
                .ContabTransfer = False
                .SegundoParametro = ""
                'Los textos
'                .vTextos = Text1(2).Text & "|" & Me.txtCta(0).Text & " - " & Me.txtDescCta(0).Text & "|" & SubTipo & "|"
                
                'Marzo2013   Cobramos un solo cliente
                'Aparecera un boton para traer todos los cobros
                '.CodmactaUnica = "4300000001" 'Trim(txtCtaNormal(9).Text)
                .Show vbModal
            End With
        
        Case 606 ' compensaciones
            frmTESCompensaciones.Show vbModal
        Case 607 ' compensar cliente
            CadenaDesdeOtroForm = ""
            frmTESCompensaAboCli.Show vbModal
        Case 608 ' reclamaciones
            frmTESReclamaCli.Show vbModal
        Case 609 ' remesas
            frmTESRemesas.Tipo = 1 ' efectos
            frmTESRemesas.Show vbModal
        Case 610 ' Informe Impagados
            frmTESCobrosDevList.Show vbModal
        Case 611 ' Recepción Talón-Pagaré
            frmTESRecepcionDoc.Show vbModal
        Case 612 ' Remesas Talón-Pagaré
            frmTESRemesasTP.Tipo = 2 ' talon pagare
            frmTESRemesasTP.Show vbModal
            
        Case 613 ' Norma 57: Pago por ventanilla
            frmTESNorma57.Opcion = 42
            frmTESNorma57.Show vbModal
            
        Case 614 ' transferencia abonos
            frmTESTransferencias.TipoTrans = 0 ' de abonos
            frmTESTransferencias.Show vbModal
            
        Case 709 ' Abono remesa
        Case 710 ' Devoluciones
        Case 711 ' Eliminar riesgo
        
        Case 801 ' Cartera de Pagos
            frmTESPagos.Show vbModal
        Case 802 ' Informe Pagos pendientes
            frmTESPagosPdtesList.Show vbModal
        Case 803 ' Informe Pagos bancos
            frmTESPagosBancoList.Show vbModal
        Case 804 ' Realizar Pago
            frmTESRealizarPagos.Show vbModal
        Case 805 ' Transferencias
            frmTESTransferencias.TipoTrans = 1 ' de pagos
            frmTESTransferencias.Show vbModal
        Case 806 ' Pagos domiciliados
            frmTESTransferencias.TipoTrans = 2 ' pagos domiciliados
            frmTESTransferencias.Show vbModal
        
        Case 807 ' Gastos Fijos
            frmTESGastosFijos.Show vbModal
        
        Case 808 ' Memoria Pagos proveedores
        
        Case 809 ' Compensar proveedor
            CadenaDesdeOtroForm = ""
            frmTESCompensaAboPro.Show vbModal
        
        Case 810 ' Confirming
            frmTESTransferencias.TipoTrans = 3 ' confirming
            frmTESTransferencias.Show vbModal
        
        Case 901 ' Informe por NIF
            frmTESInfSituacionNIF.Show vbModal
            
        Case 902 ' Informe por cuenta
            frmTESInfSituacionCta.Show vbModal
        
        Case 903 ' Situación Tesoreria
            frmTESInfSituacion.Show vbModal
        
        
        ' Analitica
        Case 1001 ' Centros de Coste
            frmCCCentroCoste.Show vbModal
            
        Case 1002 ' Consulta de Saldos
            frmCCConExtr.Show vbModal
        
        Case 1003 ' Cuenta de Explotación
            frmCCCtaExplo.Show vbModal
        Case 1004 ' Centros de coste por cuenta
            AbrirListado 17, False
        Case 1005 ' Detalle de explotación
            frmCCDetalleExplota.Show vbModal
            
        ' Presupuestaria
        Case 1101 ' Presupuestos
            Screen.MousePointer = vbHourglass
            'frmColPresu.Show vbModal
            frmPresu.Show vbModal
        Case 1102 ' Listado de Presupuestos
'            AbrirListado 9, False
        Case 1103 ' Balance Presupuestario
            frmPresuBal.Show vbModal
            
        ' Consolidado
        Case 1201 ' Sumas y Saldos
            AbrirListado 24, False
        Case 1202 ' Balance de Situación
            AbrirListado 51, False
        Case 1203 ' Pérdidas y Ganancias
            AbrirListado 50, False
        Case 1204 ' Cuenta de Explotación
            AbrirListado 31, False
        Case 1205 ' Listado Facturas Clientes
            AbrirListado 53, False
        Case 1206 ' Listado Facturas Proveedores
            AbrirListado 52, False
        
        ' Cierre de Ejercicio
        Case 1301 ' Renumeración de asientos
            frmCierre.Opcion = 0
            frmCierre.Show vbModal
        Case 1302 ' Simulación de cierre
            frmCierre.Opcion = 4
            frmCierre.Show vbModal
        Case 1303 ' Cierre de Ejercicio
            frmCierre.Opcion = 1
            frmCierre.Show vbModal
        Case 1304 ' Deshacer cierre
            frmCierre.Opcion = 5
            frmCierre.Show vbModal
        Case 1305 ' Diario Oficial
'            AbrirListado 14, False
        Case 1306 ' Diario Oficial Resumen
'            AbrirListado 18, False
            frmInfDiarioOficial.Show vbModal
        Case 1307 ' Presentación cuentas anuales
            Telematica 0
        Case 1308 ' Presentación Telemática de Libros
            Telematica 1
        Case 1309 ' memoria de Plazos de Pago
            frmTESMemoriaPlazos.Show vbModal
        
        ' Utilidades
        Case 1401 ' Comprobar cuadre
            Screen.MousePointer = vbHourglass
            frmMensajes.Opcion = 2
            frmMensajes.Show vbModal
        Case 1403 ' Revisar caracteres especiales
'            Screen.MousePointer = vbHourglass
'            frmMensajes.opcion = 14
'            frmMensajes.Show vbModal
        
        Case 1404 ' Agrupacion cuentas
        Case 1405 'Buscar ...
        
        Case 1407 'Desbloquear asientos
            mnHerrAriadnaCC_Click (0)
        Case 1408 'Mover cuentas
            mnHerrAriadnaCC_Click (1)
        Case 1409 'Renumerar registros proveedor
            mnHerrAriadnaCC_Click (5)
        Case 1410 'Aumentar dígitos contables
            mnHerrAriadnaCC_Click (3)
        Case 1411 'cambio de iva
            mnHerrAriadnaCC_Click (4)
        Case 1412 'log de acciones
            Screen.MousePointer = vbHourglass
            Load frmLog
            DoEvents
            frmLog.Show vbModal
            Screen.MousePointer = vbDefault
        Case 1413 'usuarios activos
        
        Case Else
  
    End Select

End Sub


Private Sub AbrirFormulariosAyuda(Accion As Long)

    Select Case Accion
        Case 4
            'Zona ARIADNA
            LanzaVisorMimeDocumento Me.hWnd, vParam.WebSoporte '"http://www.ariadnasw.com/"
        Case 6
            'CAlendario del contribuyente
            LanzaVisorMimeDocumento Me.hWnd, "http://www.agenciatributaria.es/AEAT.internet/Bibl_virtual/folletos/calendario_contribuyente.shtml"
        Case 7
            'licencia de usuario final
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & "/Licenciadeuso.html" ' "http://www.ariadnasw.com/clientes/"
        Case 8 ' documentos
            frmVarios.Opcion = Accion - 2
            frmVarios.Show vbModal
        Case 9 ' ayuda
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & "/Ariconta-6.html"  ' "http://www.ariadnasw.com/clientes/"
        
        Case 10 ' arimailges.exe
            Dim Lanza As String
            Dim Aux As String
            
            
            Lanza = vParam.MailSoporte & "||"
            
            'Aqui pondremos lo del texto del BODY
            Lanza = Lanza & "|"
            'Envio o mostrar
            Lanza = Lanza & "0"   '0. Display   1.  send
            
            'Campos reservados para el futuro
            Lanza = Lanza & "||||"
            
            'El/los adjuntos
            Lanza = Lanza & "|"
            
            Aux = App.Path & "\ARIMAILGES.EXE" & " " & Lanza  '& vParamAplic.ExeEnvioMail & " " & Lanza
            Shell Aux, vbNormalFocus
        
        Case 12 ' Informacion de la base de datos
            If CargarInformacionBBDD Then
                Set frmMens = New frmMensajes
            
                frmMens.Opcion = 25
                frmMens.Show vbModal
            
                Set frmMens = Nothing
            End If
            
        Case 13
            ' Panel de control donde seleccionamos los iconos que vamos a mostrar
            frmMensajes.Opcion = 24
            frmMensajes.Show vbModal
            
            If Reorganizar Then
                Reorganizar = False
                CargaShortCuts 0
                'RecolocarItems vUsu.Codigo, Me.ListView1, "ariconta"
            End If
            
        Case 14 'Usuarios activos
            'mnUsuariosActivos_Click
            Set frmMens = New frmMensajes
            
            frmMens.Opcion = 26
            frmMens.Show vbModal
            
            Set frmMens = Nothing
        
        
        Case Else
'            frmCalendario.Show vbModal
    End Select
    
End Sub


Private Function CargarInformacionBBDD() As String
Dim SQL As String
Dim SQL2 As String
Dim CadValues As String
Dim NroRegistros As Long
Dim NroRegistrosSig As Long
Dim NroRegistrosTot As Long
Dim NroRegistrosTotSig As Long
Dim FecIniSig As Date
Dim FecFinSig As Date
Dim Porcen1 As Currency
Dim Porcen2 As Currency
Dim RS As ADODB.Recordset

    On Error GoTo eCargarInformacionBBDD
    
    CargarInformacionBBDD = False
    
    SQL = "delete from tmpinfbbdd where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    FecIniSig = DateAdd("yyyy", 1, vParam.fechaini)
    FecFinSig = DateAdd("yyyy", 1, vParam.fechafin)
    
    SQL2 = "insert into tmpinfbbdd (codusu,posicion,concepto,nactual,poractual,nsiguiente,porsiguiente) values "
    
    'asientos
    SQL = "select count(*) from hcabapu where fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    NroRegistros = DevuelveValor(SQL)
    SQL = "select count(*) from hcabapu where fechaent between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
    NroRegistrosSig = DevuelveValor(SQL)
    
    CadValues = "(" & vUsu.Codigo & ",1,'Asientos'," & DBSet(NroRegistros, "N") & ",0," & DBSet(NroRegistrosSig, "N") & ",0)"
    Conn.Execute SQL2 & CadValues
    
    'apuntes
    SQL = "select count(*) from hlinapu where fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    NroRegistros = DevuelveValor(SQL)
    SQL = "select count(*) from hlinapu where fechaent between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
    NroRegistrosSig = DevuelveValor(SQL)
    
    CadValues = "(" & vUsu.Codigo & ",2,'Apuntes'," & DBSet(NroRegistros, "N") & ",0," & DBSet(NroRegistrosSig, "N") & ",0)"
    Conn.Execute SQL2 & CadValues
    
    'facturas de venta
    SQL = "select count(*) from factcli where "
    SQL = SQL & " fecfactu between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    
    NroRegistrosTot = DevuelveValor(SQL)
    
    
    SQL = "select count(*) from factcli where "
    SQL = SQL & " fecfactu between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
    
    NroRegistrosTotSig = DevuelveValor(SQL)
    
    i = 3
    
    SQL = "select * from contadores where not tiporegi in ('0','1')"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
    
        SQL = "select count(*) from factcli where numserie = " & DBSet(RS!tiporegi, "T")
        SQL = SQL & " and fecfactu between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    
        NroRegistros = DevuelveValor(SQL)
        Porcen1 = 0
        If NroRegistrosTot <> 0 Then
            Porcen1 = Round(NroRegistros * 100 / NroRegistrosTot, 2)
        End If
        
        SQL = "select count(*) from factcli where numserie = " & DBSet(RS!tiporegi, "T")
        SQL = SQL & " and fecfactu between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
        
        NroRegistrosSig = DevuelveValor(SQL)
        Porcen2 = 0
        If NroRegistrosTotSig <> 0 Then
            Porcen2 = Round(NroRegistrosSig * 100 / NroRegistrosTotSig, 2)
        End If
    
        CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(RS!nomregis, "T") & "," & DBSet(NroRegistros, "N") & "," & DBSet(Porcen1, "N") & ","
        CadValues = CadValues & DBSet(NroRegistrosSig, "N") & "," & DBSet(Porcen2, "N") & ")"
        Conn.Execute SQL2 & CadValues
        
        i = i + 1
    
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    'facturas de proveedor
    i = i + 1
    
    SQL = "select count(*) from factpro where fecharec between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    NroRegistros = DevuelveValor(SQL)
    SQL = "select count(*) from factpro where fecharec between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
    NroRegistrosSig = DevuelveValor(SQL)
    
    CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & ",'Facturas Proveedores'," & DBSet(NroRegistros, "N") & ",0,"
    CadValues = CadValues & DBSet(NroRegistrosSig, "N") & ",0)"
    
    Conn.Execute SQL2 & CadValues
    CargarInformacionBBDD = True
    Exit Function


eCargarInformacionBBDD:
    MuestraError Err.Number, "Cargar Temporal de BBDD", Err.Description
End Function



Private Sub mnHerrAriadnaCC_Click(Index As Integer)
 
        If vUsu.Nivel > 1 Then
            MsgBox "No tiene permisos", vbExclamation
            Exit Sub
        End If
        'El index 3 , que es la barra, en frmCC es la opcion de NUEVA EMPRESA
        ' y no se llma desde aqui, con lo cual no hay problemo
        'Para el restro cojo el valor del helpidi
        
        frmCentroControl.Opcion = Index
        frmCentroControl.Show vbModal
    
End Sub

Private Sub Telematica(Caso As Integer)
        Me.Enabled = False
        frmTelematica.Opcion = Caso
        frmTelematica.Show
End Sub
    

'El usuarios si tiene maximizada unas cosas y minimiazadas otras se las guardaremos
Private Sub MenuComoEstaba(ByRef TW1 As TreeView, aplicacion As String)
Dim N As Node
Dim SQL As String

    For i = 1 To TW1.Nodes.Count
        If aplicacion = "introcon" Then
            TW1.Nodes(i).Expanded = True
        Else
            SQL = "select expandido from menus_usuarios where codusu = " & DBSet(vUsu.Id, "N") & " and  aplicacion = '" & aplicacion & "' and codigo in (select codigo from menus where descripcion = " & DBSet(Me.TreeView1.Nodes(i), "T") & " and aplicacion = '" & aplicacion & "')"
    
            If DevuelveValor(SQL) = 0 Then
                TW1.Nodes(i).Expanded = False
            Else
                TW1.Nodes(i).Expanded = True
            End If
        End If
    Next i

End Sub

Private Sub OcultarHijos(Padre As String)
Dim SQL As String

    SQL = "update menus_usuarios set ver = 0 where codusu = " & vUsu.Id & " and padre = " & DBSet(Padre, "N")

    Conn.Execute SQL
    
End Sub

Private Sub CargaShortCuts(Seleccionado As Long)
Dim Aux As String
Dim RS As ADODB.Recordset
Dim SQL As String
Dim CadAux As String
 
 
    'Para cada usuario, y a partir del menu del que disponga
    Set miRsAux = New ADODB.Recordset
    Aux = "Select menus.imagen, menus.codigo, menus.descripcion from menus_usuarios inner join menus on menus_usuarios.codigo = menus.codigo and menus_usuarios.aplicacion = menus.aplicacion "
    Aux = Aux & " WHERE codusu =" & vUsu.Id & " AND menus.aplicacion='ariconta' and menus_usuarios.ver = 1 and menus.imagen <> 0 and menus_usuarios.vericono = 1 "
    
    
    If Not vEmpresa.TieneTesoreria Then
        
        Aux = Aux & " and tipo = 0"
    
    End If
    
    If Not vEmpresa.TieneContabilidad Then
    
        Aux = Aux & " and tipo = 2"
    
    
    End If
    
    
    If Reorganizar Then Me.ListView1.Arrange = lvwAutoTop
    
    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    AnchoListview = 0
    ListView1.ListItems.Clear
    While Not miRsAux.EOF
        
        If Not BloqueaPuntoMenu(miRsAux!Codigo, "ariconta") Then
            AnchoListview = AnchoListview + 1
            
            Me.ListView1.ListItems.Add , CStr("LW" & Format(miRsAux!Codigo, "000000")), DBLet(miRsAux!Descripcion, "T"), CInt(miRsAux!imagen)
        End If
        miRsAux.MoveNext
    Wend

    miRsAux.Close
    Set miRsAux = Nothing

    If Not Reorganizar Then
        For i = 1 To ListView1.ListItems.Count
            Set RS = New ADODB.Recordset
            
            SQL = "select posx, posy from menus_usuarios where aplicacion = 'ariconta' and codusu = " & vUsu.Id & " and posx <> 0 and codigo in (select codigo from menus where aplicacion = 'ariconta' and descripcion = " & DBSet(ListView1.ListItems(i).Text, "T") & ")"
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not RS.EOF Then
                Me.ListView1.ListItems(i).Left = DBLet(RS.Fields(0).Value)
                Me.ListView1.ListItems(i).top = DBLet(RS.Fields(1).Value)
            End If
            Set RS = Nothing
        Next i
    Else
        Reorganizar = False
        Me.ListView1.Arrange = lvwNone
    End If

    Set Me.ListView1.SelectedItem = Nothing
    

End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    If ListView1.SelectedItem Is Nothing Then Exit Sub
'
'    AbrirFormularios CLng(Mid(ListView1.SelectedItem.Key, 3))

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ListView1_DblClick
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim SQL As String
Dim RS As ADODB.Recordset

     If Button = 2 Then
        'PopupMenu mnPopUp
    Else
        If Not ListView1.SelectedItem Is Nothing Then
            SQL = "select posx, posy from menus_usuarios where codusu = " & vUsu.Id & " and aplicacion = 'ariconta' and "
            SQL = SQL & " codigo in (select codigo from menus where aplicacion = 'ariconta' and descripcion =  " & DBSet(ListView1.SelectedItem, "T") & ")"
            
            Set RS = New ADODB.Recordset
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not RS.EOF Then
                XAnt = RS.Fields(0)
                YAnt = RS.Fields(1)
            End If
            Set RS = Nothing
            
        End If
    End If
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If Not ListView1.SelectedItem Is Nothing Then
            IconoSeleccionado = True
'            Caption = ListView1.SelectedItem.Left & "," & ListView1.SelectedItem.Top
        End If
    End If
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim RefrescarDatos As Boolean   'Porque a movido iconos fuera de los margenes
Dim SQL As String

    If IconoSeleccionado Then
        RefrescarDatos = False
        
        
        Dim HaDesplazado As Boolean
        HaDesplazado = False
        For J = 1 To ListView1.ListItems.Count
               ' Debug.Print "Item" & I & " : " & ListView1.ListItems(j).Left
                If ListView1.ListItems(J).Left < 0 Then
                    ListView1.ListItems(J).Left = 0
                    RefrescarDatos = True
                    HaDesplazado = True
                End If
        Next

        
        'If Not RefrescarDatos Then
            ActualizarItemCuadricula vUsu.Id, Me.ListView1, "ariconta", X, Y, RefrescarDatos
        'End If
        'If HaDesplazado Then Stop
        If RefrescarDatos Then
         '   ListView1.ListItems.Clear
         '   Debug.Print ListView1.Width
         '   ListView1.Refresh
            ListView1.PictureAlignment = lvwTopLeft
            ListView1.Arrange = lvwAutoTop
            ListView1.Refresh
            ListView1.Arrange = lvwNone
            ListView1.ListItems.Clear
            CargaShortCuts 0
        End If
        
    
    End If
    IconoSeleccionado = False

End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    
'    Caption = Data.GetData(1)
'
'    If ListView1.ListItems.Count > 8 Then
'        MsgBox "Numero maximo de accesos directos superado", vbExclamation
'        Exit Sub
'    End If
'
'
'    'Aqui tendrmos la configuracion perosnalizada
'    If TreeView1.SelectedItem = Data.GetData(1) Then
'        'OK. El nodo selecionado es el que estamos moviendo
'        If TreeView1.SelectedItem.Children > 0 Then
'            MsgBox "Solo ultimo nivel", vbExclamation
'            Exit Sub
'        End If
'    Else
'        MsgBox "Error en drag/drop", vbExclamation
'        Exit Sub
'    End If
'
'    '
'    LanzaPersonalizarEdicion Val(Mid(TreeView1.SelectedItem.Key, 3)), TreeView1.SelectedItem.Text
'
'
End Sub



Private Sub LanzaPersonalizarEdicion(Valor As Long, TextoInicio As String)
    'AHORA, de momento NO dejamos personalizar los ICONOS NI LOS textos
    'El el form al final hace un:
    '   REPLACE INTO usuarios.usuariosiconosppal(codusu,aplicacion,PuntoMenu,icono,TextoOrigen,TextoVisible) VALUES (1,'ariconta',7,1,'Parámetros','Parámetros')
    
    'frmMenusPersonalizaIconos.TextoMenu = TextoInicio
    'frmMenusPersonalizaIconos.idPuntoMenu = valor
    'frmMenusPersonalizaIconos.Show vbModal
    Msg$ = "REPLACE INTO usuarios.usuariosiconosppal(codusu,aplicacion,PuntoMenu,icono,TextoOrigen,TextoVisible) "
    Msg = Msg$ & " VALUES (" & vUsu.Codigo & ",'ariconta'," & Valor & "," & Valor
    Msg = Msg$ & "," & DBSet(TextoInicio, "T") & "," & DBSet(TextoInicio, "T") & ")"
    
    Conn.Execute Msg$
    espera 0.2
    CargaShortCuts Valor
    
End Sub



Private Sub ListView2_Click()
Dim cad As String

    CambiarEmpresa
    
    CargaMenu "ariconta", Me.TreeView1
    CargaMenu "introcon", Me.TreeView2
    
    MenuComoEstaba Me.TreeView1, "ariconta"
    MenuComoEstaba Me.TreeView2, "introcon"
'--
    CargaShortCuts 0

    ListView2.ListItems.Clear
    ListView2.SmallIcons = Me.ImageList2
    
    BuscaEmpresas
    NumeroEmpresaMemorizar True

    SituarItemList ListView2
    
End Sub





Private Sub mnPopUp1_Click(Index As Integer)
    If Index <= 1 Then
        If ListView1.SelectedItem Is Nothing Then Exit Sub
    End If
    
    
    Select Case Index
    Case 0
        'LanzaPersonalizarEdicion Val(Mid(ListView1.SelectedItem.Key, 3)), ""
    Case 1
        If MsgBox("Desea eliminar el acceso directo: " & Me.ListView1.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbYes Then
            
            Conn.Execute "DELETE from  usuarios.usuariosiconosppal WHERE codusu =" & vUsu.Codigo & " AND aplicacion='ariconta' AND PuntoMenu =" & Mid(ListView1.SelectedItem.Key, 3)
            CargaShortCuts 0
        End If
    Case 3
    
    End Select
        
    
End Sub

Private Sub TreeView1_DblClick()
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    If TreeView1.SelectedItem.Children > 0 Then Exit Sub
    
    AbrirFormularios CLng(Mid(TreeView1.SelectedItem.Key, 3))
    
End Sub


Private Sub TreeView2_DblClick()

    If TreeView2.SelectedItem Is Nothing Then Exit Sub
    If TreeView2.SelectedItem.Children > 0 Then Exit Sub
    
    AbrirFormulariosAyuda CLng(Mid(TreeView2.SelectedItem.Key, 3))

End Sub



Private Sub CambiarEmpresa()

    CadenaDesdeOtroForm = vUsu.Login & "|" & vEmpresa.codempre & "|"
        
    ActualizarExpansionMenus vUsu.Id, Me.TreeView1, "ariconta"
    
        
    Set vUsu = New Usuario
    vUsu.Leer RecuperaValor(CadenaDesdeOtroForm, 1)
    
    vUsu.CadenaConexion = ListView2.SelectedItem.ToolTipText
    
    vUsu.LeerFiltros "ariconta", 301 ' asientos
    vUsu.LeerFiltros "ariconta", 401 ' facturas de cliente
    
    AbrirConexion vUsu.CadenaConexion
    
    Set vEmpresa = New Cempresa
    Set vParam = New Cparametros
    Set vParamT = New CparametrosT
    'NO DEBERIAN DAR ERROR
    vEmpresa.Leer
    vParam.Leer
    If vEmpresa.TieneTesoreria Then vParamT.Leer
    
    PonerCaption

    NumeroEmpresaMemorizar False
    
    

End Sub

Private Sub TreeView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    If Not TreeView1.SelectedItem Is Nothing Then
        If TreeView1.SelectedItem.Children > 0 Then TreeView1.Drag vbCancel
    End If
End Sub



Private Sub mnUsuariosActivos_Click()
Dim SQL As String
Dim i As Integer
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad(False)
    If CadenaDesdeOtroForm <> "" Then
        i = 1
        Me.Tag = "Los siguientes PC's están conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
        Do
            SQL = RecuperaValor(CadenaDesdeOtroForm, i)
            If SQL <> "" Then Me.Tag = Me.Tag & "    - " & SQL & vbCrLf
            i = i + 1
        Loop Until SQL = ""
        MsgBox Me.Tag, vbExclamation
    Else
        MsgBox "Ningun usuario, además de usted, conectado a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf, vbInformation
    End If
    CadenaDesdeOtroForm = ""
End Sub


Private Sub AbrirListado(numero As Byte, Cerrado As Boolean)
'    Screen.MousePointer = vbHourglass
'    frmListado.EjerciciosCerrados = Cerrado
'    frmListado.Opcion = numero
'    frmListado.Show vbModal
End Sub


Private Sub PonerDatosFormulario()
Dim Config As Boolean

'    Config = (vParam Is Nothing) Or (vEmpresa Is Nothing)
'
'    If Not Config Then HabilitarSoloPrametros_o_Empresas True
'
'    'FijarConerrores
'    CadenaDesdeOtroForm = ""
'
'    'Poner datos visible del form
'    PonerDatosVisiblesForm
'    'Poner opciones de nivel de usuario
'    PonerOpcionesUsuario
'
'
'    If Not Config Then
'        Me.mnTraspasoEntreSecciones(0).Visible = vParam.TraspasCtasBanco > 0
'        Me.mnTraspasoEntreSecciones(1).Visible = mnTraspasoEntreSecciones(0).Visible
'    End If
'    'Habilitar
'    If Config Then HabilitarSoloPrametros_o_Empresas False
'    'Panel con el nombre de la empresa
'    If Not vEmpresa Is Nothing Then
'        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
'    Else
'        Me.StatusBar1.Panels(2).Text = "Falta configurar"
'    End If
'
'    'Primero los pongo a visible
'    mnDatosExternos347.Visible = True
'    mnbarra101.Visible = True
'
'
'
'
'    'Si tiene editor de menus
'    If TieneEditorDeMenus Then PoneMenusDelEditor
'
'     mnCheckVersion.Visible = False 'Siempre oculto
'
'
'    If Not Config Then
'        mnDatosExternos347.Visible = mnDatosExternos347.Visible And vParam.AgenciaViajes
'        mnbarra101.Visible = mnbarra101.Visible And vParam.AgenciaViajes
'    End If
'    '---------------------------------------------------
'    'Las asociaciones entre menu y botones  del TOOLBAR
'    With Me.Toolbar1
'        .Buttons(1).Visible = mnDatos.Visible And Me.mnPlanContable.Visible
'        '---
'        .Buttons(3).Visible = mnDiario.Visible And Me.mnIntroducirAsientos.Visible    'Diario
'        .Buttons(4).Visible = mnHcoApuntes.Visible And mnVerHistoricoApuntes.Visible    'Hco
'        .Buttons(5).Visible = mnDiario.Visible And mnConsultaExtractos.Visible   'Con extractos
'        .Buttons(6).Visible = mnHcoApuntes.Visible And mnCtaExplotacion.Visible   'CTA EXPLOTACION
'        '----
'        .Buttons(8).Visible = mnMenuIVA.Visible And mnClientes.Visible And Me.mnRegFacCli.Visible     'Fac CLI
'        .Buttons(9).Visible = mnMenuIVA.Visible And mnMenuProveedores.Visible And Me.mnRegFac.Visible    'Fac PRO
'        .Buttons(10).Visible = mnMenuIVA.Visible And Me.mnLiquidacion.Visible   'Liquidacion IVA
'        '----
'        .Buttons(12).Visible = mnHcoApuntes.Visible And mnBalanceMensual.Visible  'Balance
'        .Buttons(13).Visible = mnHcoApuntes.Visible And mnBalancesituacion.Visible
'        .Buttons(14).Visible = mnHcoApuntes.Visible And Me.mnPerdyGan.Visible  'Cuenta P y G
'        '----
'        .Buttons(16).Image = 8  'Usuarios
'        .Buttons(17).Image = 9  'Impresora
'        '----
'        .Buttons(19).Visible = TieneIntegracionesPendientes
'        .Buttons(19).Image = 11
'        'Antes
'        .Buttons(20).Visible = False
'        '.Buttons(20).Visible = BuscarIntegraciones(True)
'        .Buttons(20).Image = 12
'        '----
'        .Buttons(22).Image = 10 'Salir
'    End With
'
'    'Si el usuario tiene permiso para ver los balances, le dejo las graficas
'    Me.mnRatios.Visible = Toolbar1.Buttons(12).Visible
'
End Sub


Private Sub HabilitarSoloPrametros_o_Empresas(Habilitar As Boolean)
Dim T As Control
Dim cad As String
'
'    On Error Resume Next
'    For Each T In Me
'        Cad = T.Name
'        If Mid(T.Name, 1, 2) = "mn" Then
'            If LCase(Mid(T.Name, 1, 6)) <> "mnbarr" Then T.Enabled = Habilitar
'        End If
'    Next
'    Me.Toolbar1.Enabled = Habilitar
'    Me.Toolbar1.Visible = Habilitar
'    mnParametros.Enabled = True
'    mnEmpresa.Enabled = True
'    Me.mnParametros.Enabled = True
'    Me.mnConfiguracionAplicacion.Enabled = True
'    mnDatos.Enabled = True
'    Me.mnuSal.Enabled = True
'    Me.mnCambioUsuario.Enabled = True
End Sub

Private Sub PonerDatosVisiblesForm()
Dim cad As String
'    Cad = UCase(Mid(Format(Now, "dddd"), 1, 1)) & Mid(Format(Now, "dddd"), 2)
'    Cad = Cad & ", " & Format(Now, "d")
'    Cad = Cad & " de " & Format(Now, "mmmm")
'    Cad = Cad & " de " & Format(Now, "yyyy")
'    Cad = "    " & Cad & "    "
'    Me.StatusBar1.Panels(5).Text = Cad
'    If vEmpresa Is Nothing Then
'        Caption = "ARICONTA" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & "   Usuario: " & vUsu.Nombre & " FALTA CONFIGURAR"
'    Else
'        'Caption = "ARICONTA" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & vEmpresa.nomempre & "  -    Usuario: " & vUsu.Nombre
'        Caption = "ARICONTA" & " Ver. " & App.Major & "." & App.Minor & "." & App.Revision & "    " & vEmpresa.nomresum & "     Usuario: " & vUsu.Nombre
'    End If
End Sub


Private Sub PonerOpcionesUsuario()
    Dim B As Boolean

'
'    'SOLO ROOT
'    B = (vUsu.Codigo Mod 1000) = 0
'    Me.mnTraerDeCerrados.Visible = B
'    Me.mnUsuarios.Enabled = B
'
'    B = vUsu.Nivel < 2  'Administradores y root
'    Me.mnParametros.Enabled = B
'    Me.mnEmpresa.Enabled = B
'    Me.mnParametrosInmo.Enabled = B
'    Me.mnHerramientasAriadnaCC.Enabled = B
'    If B Then
'        'Si tiene permiso solo admin podra  subir ctas
'
'
'    End If
'
'
'
'    mnAsiePerdyGana.Enabled = B
'    mnRenumeracion.Enabled = B
'    mnTraspasoACerrados.Enabled = B
'    mnBorrarProveedores.Enabled = B
'    mnBorrarRegClientes.Enabled = B
'    mnDescierre.Enabled = B
'    mnVentaBajaInmo.Enabled = B
'    mnCaluloYContabilizacion.Enabled = B
'    mnDeshacerAmortizacion.Enabled = B
'    mnNuevaEmpresa.Enabled = B
'    mnRecalculoSaldos.Enabled = B
'    mnInformesScrystal.Enabled = B
'    Me.mnImportarDatosFiscales.Enabled = B
'
'    mnVerLog.Visible = B
'
'    'mnPedirPwd.Enabled = B
'    B = vUsu.Nivel = 3  'Es usuario de consultas
'    If B Then
'        mnBorreEjerciciosCerrados.Enabled = False
'        mnDiarioOficial.Enabled = False
'        mnActalizacionAsientos.Enabled = False
'        mnAsientosPredefinidos.Enabled = False
'        Me.mnConfigBalPeryGan.Enabled = False
'        Me.mnContFactCli.Enabled = False
'        Me.mnContFactProv.Enabled = False
'        Me.mnPunteoExtractos.Enabled = False
'        Me.mnImportarNorma43.Enabled = False
'        Me.mnPunteoBancario.Enabled = False
'        Me.mnImportarDatosFiscales.Enabled = False
'    End If
End Sub

'''ICONOS
Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
    Dim i As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = op
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub

Public Sub EstablecerSkin(QueSkin As Integer)

    'Añadido pero hay que preguntar ???
    If QueSkin = -1 Then Exit Sub


  FijaSkin QueSkin

  ' Cargando el archivo del Skin
  ' ============================
    frmPpal.SkinFramework.LoadSkin Skn$, ""
    frmPpal.SkinFramework.ApplyWindow frmPpal.hWnd
    frmPpal.SkinFramework.ApplyOptions = frmPpal.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
'
    
End Sub

Private Function FijaSkin(Cual As Integer)
  Select Case (Cual)
    Case 0:     ' Windows Luna XP Modificado
      Skn$ = CStr(App.Path & "\Styles\WinXP.Luna.cjstyles")
      frmPpal.SkinFramework.LoadSkin Skn$, "NormalBlue.ini"
    Case 1:     ' Windows Royale Modificado
      Skn$ = CStr(App.Path & "\Styles\WinXP.Royale.cjstyles")
      frmPpal.SkinFramework.LoadSkin Skn$, "NormalRoyale.ini"
    Case 2:     ' Microsoft Office 2007
      Skn$ = CStr(App.Path & "\Styles\Office2007.cjstyles")
      frmPpal.SkinFramework.LoadSkin Skn$, "NormalBlue.ini"
    Case 3:     ' Windows Vista Sencillo
      Skn$ = CStr(App.Path & "\Styles\Vista.cjstyles")
      frmPpal.SkinFramework.LoadSkin Skn$, "NormalBlue.ini"
  End Select

End Function

Private Sub BuscaEmpresas()
Dim Prohibidas As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim cad As String
Dim ItmX As ListItem
Dim SQL As String


'Cargamos las prohibidas
Prohibidas = DevuelveProhibidas

'Cargamos las empresas
Set RS = New ADODB.Recordset

'[Monica]11/04/2014: solo debe de salir las ariconta
RS.Open "Select * from usuarios.empresasariconta where conta like 'ariconta%' ORDER BY Codempre", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

While Not RS.EOF
    cad = "|" & RS!codempre & "|"
    If InStr(1, Prohibidas, cad) = 0 Then
        cad = RS!nomempre
        Set ItmX = ListView2.ListItems.Add()
        
        ItmX.Text = cad
        ItmX.SubItems(1) = RS!nomresum
        
        ' sacamos las fechas de inicio y fin
        SQL = "select fechaini, fechafin from " & Trim(RS!CONTA) & ".parametros"
        Set Rs2 = New ADODB.Recordset
        Rs2.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not Rs2.EOF Then
            ItmX.SubItems(2) = Rs2!fechaini & " - " & Rs2!fechafin
        End If
        Set Rs2 = Nothing
        
            
        cad = RS!CONTA & "|" & RS!nomresum '& "|" & Rs!Usuario & "|" & Rs!Pass & "|"
        ItmX.Tag = cad
        ItmX.ToolTipText = RS!CONTA
        
        
        'Si el codconta > 100 son empresas que viene del cambio del plan contable.
        'Atenuare su visibilidad
        If RS!codempre > 100 Then
            ItmX.ForeColor = &H808080
            ItmX.ListSubItems(1).ForeColor = &H808080
            ItmX.ListSubItems(2).ForeColor = &H808080
            ItmX.ListSubItems(3).ForeColor = &H808080
            ItmX.SmallIcon = 2
        Else
            'normal
            ItmX.SmallIcon = 1
        End If
    End If
    RS.MoveNext
Wend
RS.Close
End Sub


Private Function DevuelveProhibidas() As String
Dim RS As ADODB.Recordset
Dim cad As String
Dim i As Integer
    On Error GoTo EDevuelveProhibidas
    DevuelveProhibidas = ""
    Set RS = New ADODB.Recordset
    i = vUsu.Codigo Mod 1000
    RS.Open "Select * from usuarios.usuarioempresasariconta WHERE codusu =" & i, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    cad = ""
    While Not RS.EOF
        cad = cad & RS.Fields(1) & "|"
        RS.MoveNext
    Wend
    If cad <> "" Then cad = "|" & cad
    RS.Close
    DevuelveProhibidas = cad
EDevuelveProhibidas:
    Err.Clear
    Set RS = Nothing
End Function



Private Sub NumeroEmpresaMemorizar(Leer As Boolean)
Dim NF As Integer
Dim C1 As String
Dim cad As String
Dim Cad2 As String


On Error GoTo ENumeroEmpresaMemorizar

    If Leer Then
        If CadenaDesdeOtroForm <> "" Then
            'Ya estabamos trabajando con la aplicacion
            
            If Not (vEmpresa Is Nothing) Then
                 For NF = 1 To Me.ListView2.ListItems.Count
                    If ListView2.ListItems(NF).Text = vEmpresa.nomempre Then
                        Set ListView2.SelectedItem = ListView2.ListItems(NF)
                        ListView2.SelectedItem.EnsureVisible
                        Exit For
                    End If
                Next NF
            End If
            
                'El tercer pipe, si tiene es el ancho col1
                cad = AnchoLogin
                C1 = vControl.Ancho1
                If Val(C1) > 0 Then
                    NF = Val(C1)
                Else
                    NF = 4360
                End If
                ListView2.ColumnHeaders(1).Width = NF
                
                'El cuarto pipe si tiene es el ancho de col2
                C1 = vControl.Ancho2
                If Val(C1) > 0 Then
                    NF = Val(C1)
                Else
                    NF = 1400
                End If
                ListView2.ColumnHeaders(2).Width = NF
            
            
                'El quinto pipe si tiene es el ancho de col2
                C1 = vControl.Ancho3
                
                'DAVID
                'LO hablamos con calma
                C1 = 3000
                If Val(C1) > 0 Then
                    NF = Val(C1)
                Else
                    NF = 1400
                End If
                ListView2.ColumnHeaders(3).Width = NF
            
                vUsu.LeerFiltros "ariconta", 301 'asientos
                vUsu.LeerFiltros "ariconta", 401 'facturas de cliente
            
            
            CadenaDesdeOtroForm = ""
            Exit Sub
        End If
    End If
    
    If Leer Then
        If Not vControl Is Nothing Then
                'El primer pipe es el usuario. Como ya no lo necesito, no toco nada
                
                C1 = vControl.UltEmpre
                'el segundo es el
                If C1 <> "" Then
                    For NF = 1 To Me.ListView2.ListItems.Count
                        If ListView2.ListItems(NF).Text = C1 Then
                            Set ListView2.SelectedItem = ListView2.ListItems(NF)
                            ListView2.SelectedItem.EnsureVisible
                            Exit For
                        End If
                    Next NF
                End If
                
                'El tercer pipe, si tiene es el ancho col1
                C1 = vControl.Ancho1
                If Val(C1) > 0 Then
                    NF = Val(C1)
                Else
                    NF = 4360
                End If
                ListView2.ColumnHeaders(1).Width = NF
                'El cuarto pipe si tiene es el ancho de col2
                C1 = vControl.Ancho2
                If Val(C1) > 0 Then
                    NF = Val(C1)
                Else
                    NF = 1400
                End If
                ListView2.ColumnHeaders(2).Width = NF
                'El quinto pipe si tiene es el ancho de col3
                C1 = vControl.Ancho3
                If Val(C1) > 0 Then
                    NF = Val(C1)
                Else
                    NF = 1400
                End If
                ListView2.ColumnHeaders(3).Width = NF
            
                vUsu.LeerFiltros "ariconta", 301 'asientos
                vUsu.LeerFiltros "ariconta", 401 'facturas de cliente
                
        End If
    Else 'Escribir
'        cad = Cad2
        vControl.UltEmpre = ListView2.SelectedItem.ToolTipText
        vControl.Ancho1 = Int(Round(ListView2.ColumnHeaders(1).Width, 2))
        vControl.Ancho2 = Int(Round(ListView2.ColumnHeaders(2).Width, 2))
        vControl.Ancho3 = Int(Round(ListView2.ColumnHeaders(3).Width, 2))
        
        vControl.Grabar
        
        vUsu.CadenaConexion = vControl.UltEmpre
        
        AnchoLogin = cad
    End If
ENumeroEmpresaMemorizar:
    Err.Clear
End Sub

'Private Sub NumeroEmpresaMemorizar(Leer As Boolean)
'Dim NF As Integer
'Dim C1 As String
'Dim cad As String
'Dim Cad2 As String
'
'
'On Error GoTo ENumeroEmpresaMemorizar
'
'    If Leer Then
'        If CadenaDesdeOtroForm <> "" Then
'            'Ya estabamos trabajando con la aplicacion
'
'            If Not (vEmpresa Is Nothing) Then
'                 For NF = 1 To Me.ListView2.ListItems.Count
'                    If ListView2.ListItems(NF).Text = vEmpresa.nomempre Then
'                        Set ListView2.SelectedItem = ListView2.ListItems(NF)
'                        ListView2.SelectedItem.EnsureVisible
'                        Exit For
'                    End If
'                Next NF
'            End If
'
'                'El tercer pipe, si tiene es el ancho col1
'                cad = AnchoLogin
'                C1 = RecuperaValor(cad, 3)
'                If Val(C1) > 0 Then
'                    NF = Val(C1)
'                Else
'                    NF = 4360
'                End If
'                ListView2.ColumnHeaders(1).Width = NF
'
'                'El cuarto pipe si tiene es el ancho de col2
'                C1 = RecuperaValor(cad, 4)
'                If Val(C1) > 0 Then
'                    NF = Val(C1)
'                Else
'                    NF = 1400
'                End If
'                ListView2.ColumnHeaders(2).Width = NF
'
'
'                'El quinto pipe si tiene es el ancho de col2
'                C1 = RecuperaValor(cad, 5)
'
'                'DAVID
'                'LO hablamos con calma
'                C1 = 3000
'                If Val(C1) > 0 Then
'                    NF = Val(C1)
'                Else
'                    NF = 1400
'                End If
'                ListView2.ColumnHeaders(3).Width = NF
'
'                vUsu.LeerFiltros "ariconta", 301 'asientos
'                vUsu.LeerFiltros "ariconta", 401 'facturas de cliente
'
'
'            CadenaDesdeOtroForm = ""
'            Exit Sub
'        End If
'    End If
'    cad = App.Path & "\control.dat"
'    If Leer Then
'        If Dir(cad) <> "" Then
'            NF = FreeFile
'            Open cad For Input As #NF
'            Line Input #NF, cad
'            Close #NF
'            cad = Trim(cad)
'            If cad <> "" Then
'                'El primer pipe es el usuario. Como ya no lo necesito, no toco nada
'
'                C1 = RecuperaValor(cad, 2)
'                'el segundo es el
'                If C1 <> "" Then
'                    For NF = 1 To Me.ListView2.ListItems.Count
'                        If ListView2.ListItems(NF).Text = C1 Then
'                            Set ListView2.SelectedItem = ListView2.ListItems(NF)
'                            ListView2.SelectedItem.EnsureVisible
'                            Exit For
'                        End If
'                    Next NF
'                End If
'
'                'El tercer pipe, si tiene es el ancho col1
'                C1 = RecuperaValor(cad, 3)
'                If Val(C1) > 0 Then
'                    NF = Val(C1)
'                Else
'                    NF = 4360
'                End If
'                ListView2.ColumnHeaders(1).Width = NF
'                'El cuarto pipe si tiene es el ancho de col2
'                C1 = RecuperaValor(cad, 4)
'                If Val(C1) > 0 Then
'                    NF = Val(C1)
'                Else
'                    NF = 1400
'                End If
'                ListView2.ColumnHeaders(2).Width = NF
'                'El quinto pipe si tiene es el ancho de col3
'                C1 = RecuperaValor(cad, 5)
'                If Val(C1) > 0 Then
'                    NF = Val(C1)
'                Else
'                    NF = 1400
'                End If
'                ListView2.ColumnHeaders(3).Width = NF
'
'                vUsu.LeerFiltros "ariconta", 301 'asientos
'                vUsu.LeerFiltros "ariconta", 401 'facturas de cliente
'
'
'            End If
'        End If
'    Else 'Escribir
'        NF = FreeFile
'        Open cad For Output As #NF
'
'        Cad2 = CadenaControl
'
'        Cad2 = InsertaValor(Cad2, 2, ListView2.SelectedItem.ToolTipText)
'        Cad2 = InsertaValor(Cad2, 3, Int(Round(ListView2.ColumnHeaders(1).Width, 2)))
'        Cad2 = InsertaValor(Cad2, 4, Int(Round(ListView2.ColumnHeaders(2).Width, 2)))
'        Cad2 = InsertaValor(Cad2, 5, Int(Round(ListView2.ColumnHeaders(3).Width, 2)))
'
'        CadenaControl = Cad2
'
'        cad = Cad2
'
'        AnchoLogin = cad
'        Print #NF, Cad2
'        Close #NF
'    End If
'ENumeroEmpresaMemorizar:
'    Err.Clear
'End Sub



