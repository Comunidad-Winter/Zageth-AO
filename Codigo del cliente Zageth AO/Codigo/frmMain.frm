VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "Cswsk32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{B370EF78-425C-11D1-9A28-004033CA9316}#2.0#0"; "Captura.ocx"
Begin VB.Form FrmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8985
   ClientLeft      =   390
   ClientTop       =   300
   ClientWidth     =   11910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmMain.frx":030A
   Picture         =   "frmMain.frx":35525
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   5040
      Top             =   2040
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox MiniMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1515
      Left            =   10200
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   34
      Top             =   7320
      Width           =   1440
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   11
      Left            =   6120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   33
      Top             =   8400
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   10
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   29
      Top             =   8400
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   9
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   525
      TabIndex        =   28
      Top             =   8400
      Width           =   525
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   8
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   525
      TabIndex        =   27
      Top             =   8400
      Width           =   525
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   7
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   525
      TabIndex        =   26
      Top             =   8400
      Width           =   525
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   6
      Left            =   3120
      ScaleHeight     =   495
      ScaleWidth      =   525
      TabIndex        =   25
      Top             =   8400
      Width           =   525
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   2640
      ScaleHeight     =   495
      ScaleWidth      =   450
      TabIndex        =   24
      Top             =   8400
      Width           =   450
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   450
      TabIndex        =   23
      Top             =   8400
      Width           =   450
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   1440
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   22
      Top             =   8400
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   840
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   21
      Top             =   8400
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   20
      Top             =   8400
      Width           =   495
   End
   Begin VB.ListBox hlst 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2985
      ItemData        =   "frmMain.frx":36167
      Left            =   8640
      List            =   "frmMain.frx":36169
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   9000
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   159
      TabIndex        =   9
      Top             =   2280
      Width           =   2385
   End
   Begin VB.Timer AntiExternos 
      Interval        =   15000
      Left            =   5520
      Top             =   2040
   End
   Begin VB.Timer AntiEngine 
      Interval        =   100
      Left            =   6000
      Top             =   2040
   End
   Begin VB.Timer timerUclick 
      Interval        =   500
      Left            =   6480
      Top             =   2040
   End
   Begin Captura.wndCaptura Foto 
      Left            =   4080
      Top             =   2040
      _ExtentX        =   688
      _ExtentY        =   688
   End
   Begin VB.TextBox SendCMSTXT 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1785
      Visible         =   0   'False
      Width           =   8145
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4560
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   7440
      Top             =   2040
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6960
      Top             =   2040
   End
   Begin VB.Timer FPS 
      Interval        =   1000
      Left            =   7920
      Top             =   2040
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1785
      Visible         =   0   'False
      Width           =   8145
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1380
      Left            =   240
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   240
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   2434
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":3616B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Modhabla 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hablar"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7680
      TabIndex        =   32
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   9240
      Top             =   8040
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   150
      Index           =   2
      Left            =   9000
      Top             =   7800
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   1
      Left            =   8760
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   0
      Left            =   8880
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   150
      Index           =   1
      Left            =   8880
      TabIndex        =   30
      Top             =   840
      Width           =   1815
   End
   Begin VB.Shape ExpShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8880
      Top             =   870
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   19
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label lblAgi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   18
      Top             =   8640
      Width           =   375
   End
   Begin VB.Label lblFuerza 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   9120
      TabIndex        =   17
      Top             =   8640
      Width           =   375
   End
   Begin VB.Label GldLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100000000000"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10575
      TabIndex        =   16
      Top             =   5760
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   420
      Index           =   0
      Left            =   10200
      Top             =   5640
      Width           =   435
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Zageth AO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8520
      TabIndex        =   15
      Top             =   240
      Width           =   2865
   End
   Begin VB.Image CmdLanzar 
      Height          =   435
      Left            =   8640
      MouseIcon       =   "frmMain.frx":361E8
      MousePointer    =   99  'Custom
      Top             =   5040
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Image CmdInfo 
      Height          =   435
      Left            =   10920
      MouseIcon       =   "frmMain.frx":3633A
      MousePointer    =   99  'Custom
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   495
      Index           =   1
      Left            =   11280
      MouseIcon       =   "frmMain.frx":3648C
      MousePointer    =   99  'Custom
      Top             =   3360
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   495
      Index           =   0
      Left            =   11280
      MouseIcon       =   "frmMain.frx":365DE
      MousePointer    =   99  'Custom
      Top             =   2880
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   10080
      MouseIcon       =   "frmMain.frx":36730
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   1200
      Width           =   1725
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8520
      MouseIcon       =   "frmMain.frx":36882
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   1200
      Width           =   1605
   End
   Begin VB.Image InvEqu 
      Height          =   480
      Left            =   8520
      Picture         =   "frmMain.frx":369D4
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   9000
      TabIndex        =   12
      Top             =   4680
      Width           =   2475
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11280
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label HamBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   10320
      TabIndex        =   11
      Top             =   6240
      Width           =   1350
   End
   Begin VB.Label StaBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8760
      TabIndex        =   7
      Top             =   6720
      Width           =   1350
   End
   Begin VB.Label ManaBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8760
      TabIndex        =   6
      Top             =   6240
      Width           =   1350
   End
   Begin VB.Label HpBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8760
      TabIndex        =   5
      Top             =   5880
      Width           =   1350
   End
   Begin VB.Shape Hpshp 
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   8760
      Top             =   5880
      Width           =   1350
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   180
      Left            =   8760
      Top             =   6720
      Width           =   1350
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   180
      Left            =   8760
      Top             =   6240
      Width           =   1350
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   180
      Left            =   10320
      Top             =   6240
      Width           =   1350
   End
   Begin VB.Label AguBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   10320
      TabIndex        =   4
      Top             =   6720
      Width           =   1350
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   180
      Left            =   10320
      Top             =   6720
      Width           =   1350
   End
   Begin VB.Label LvlLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11040
      TabIndex        =   3
      Top             =   840
      Width           =   210
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "350/350"
      ForeColor       =   &H00FFFFFF&
      Height          =   75
      Left            =   8880
      TabIndex        =   2
      Top             =   9360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image cmdCerrar 
      Height          =   150
      Left            =   11640
      Top             =   240
      Width           =   120
   End
   Begin VB.Image cmdMinimizar 
      Height          =   150
      Left            =   11400
      Top             =   240
      Width           =   225
   End
   Begin VB.Shape MainViewShp 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   6300
      Left            =   240
      Top             =   2040
      Width           =   8130
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Public ActualSecond As Long
Public LastSecond As Long
Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Public SelM As Integer
Public MapMapa As Integer
Dim gDSB As DirectSoundBuffer
Dim gD As DSBUFFERDESC
Dim gW As WAVEFORMATEX
Dim gFileName As String
Dim dsE As DirectSoundEnum
Dim Pos(0) As DSBPOSITIONNOTIFY
Public IsPlaying As Byte
Dim endEvent As Long
Private TiempoActual As Long
Private Contador As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim PuedeMacrear As Boolean

'Anti Engine By NicoNZ
Private ElDeAhora As Double
Private Diferencia As Double
Private ElDeAntes As Double
Private Empezo As Boolean
Private Minimo As Double
Private Maximo As Double
Private Cont As Byte
Private EstuboDesbalanceado As Long
Private ContEngine As Byte
'/Anti Engine By NicoNZ

Implements DirectXEvent
Private Sub AntiEngine_Timer()
If logged Then
If Not logged Then Exit Sub
    If GetTickCount - TiempoActual > 110 Or GetTickCount - TiempoActual < 109 Then
        Contador = Contador + 1
    Else
        Contador = 0
    End If

    If Contador > 599 Then
        Call MsgBox("Has Sido Echado por uso de SH", vbCritical, "Chitero")
        Call SendData("BANEAMESpeedHack")
        End
        'Contador = 0 ' para que limpias la variable si el programa se cerro :s?
    End If
TiempoActual = GetTickCount()
End If

End Sub

Private Sub AntiExternos_Timer()
'If logged Then
'    ListApps
'    verify_cheats
'End If
'Nueva seguridad By NicoNZ :)

'Call enumProc Revisa procesos

If FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1.1")) Then
    Call HayExterno("CHEAT ENGINE 5.1.1")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.0")) Then
    Call HayExterno("CHEAT ENGINE 5.0")
ElseIf FindWindow(vbNullString, UCase$("Pts")) Then
    Call HayExterno("Auto Pots")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.2")) Then
    Call HayExterno("CHEAT ENGINE 5.2")
ElseIf FindWindow(vbNullString, UCase$("SOLOCOVO?")) Then
    Call HayExterno("SOLOCOVO?")
ElseIf FindWindow(vbNullString, UCase$("-=[ANUBYS RADAR]=-")) Then
    Call HayExterno("-=[ANUBYS RADAR]=-")
ElseIf FindWindow(vbNullString, UCase$("CRAZY SPEEDER 1.05")) Then
    Call HayExterno("CRAZY SPEEDER 1.05")
ElseIf FindWindow(vbNullString, UCase$("SET !XSPEED.NET")) Then
    Call HayExterno("SET !XSPEED.NET")
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.80 - UNREGISTERED")) Then
    Call HayExterno("SPEEDERXP V1.80 - UNREGISTERED")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.3")) Then
    Call HayExterno("CHEAT ENGINE 5.3")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1")) Then
    Call HayExterno("CHEAT ENGINE 5.1")
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER")) Then
    Call HayExterno("A SPEEDER")
ElseIf FindWindow(vbNullString, UCase$("MEMO :P")) Then
    Call HayExterno("MEMO :P")
ElseIf FindWindow(vbNullString, UCase$("ORK4M VERSION 1.5")) Then
    Call HayExterno("ORK4M VERSION 1.5")
ElseIf FindWindow(vbNullString, UCase$("BY FEDEX")) Then
    Call HayExterno("By Fedex")
ElseIf FindWindow(vbNullString, UCase$("!XSPEED.NET +4.59")) Then
    Call HayExterno("!Xspeeder")
ElseIf FindWindow(vbNullString, UCase$("CAMBIA TITULOS DE CHEATS BY FEDEX")) Then
    Call HayExterno("Cambia titulos")
ElseIf FindWindow(vbNullString, UCase$("NEWENG OCULTO")) Then
    Call HayExterno("Cambia titulos")
ElseIf FindWindow(vbNullString, UCase$("SERBIO ENGINE")) Then
    Call HayExterno("Serbio Engine")
ElseIf FindWindow(vbNullString, UCase$("REYMIX ENGINE 5.3 PUBLIC")) Then
    Call HayExterno("ReyMix Engine")
ElseIf FindWindow(vbNullString, UCase$("REY ENGINE 5.2")) Then
    Call HayExterno("ReyMix Engine")
ElseIf FindWindow(vbNullString, UCase$("AUTOCLICK - BY NIO_SHOOTER")) Then
    Call HayExterno("AutoClick")
ElseIf FindWindow(vbNullString, UCase$("TONNER MINER! :D [REG][SKLOV] 2.0")) Then
    Call HayExterno("Tonner")
ElseIf FindWindow(vbNullString, UCase$("Buffy The vamp Slayer")) Then
    Call HayExterno("Buffy The vamp Slayer")
ElseIf FindWindow(vbNullString, UCase$("Blorb Slayer 1.12.552 (BETA)")) Then
    Call HayExterno("Blorb Slayer 1.12.552 (BETA)")
ElseIf FindWindow(vbNullString, UCase$("PumaEngine3.0")) Then
    Call HayExterno("PumaEngine3.0")
ElseIf FindWindow(vbNullString, UCase$("Vicious Engine 5.0")) Then
    Call HayExterno("Vicious Engine 5.0")
ElseIf FindWindow(vbNullString, UCase$("AkumaEngine33")) Then
    Call HayExterno("AkumaEngine33")
ElseIf FindWindow(vbNullString, UCase$("Spuc3ngine")) Then
    Call HayExterno("Spuc3ngine")
ElseIf FindWindow(vbNullString, UCase$("Ultra Engine")) Then
    Call HayExterno("Ultra Engine")
ElseIf FindWindow(vbNullString, UCase$("Engine")) Then
    Call HayExterno("Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V5.4")) Then
    Call HayExterno("Cheat Engine V5.4")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4")) Then
    Call HayExterno("Cheat Engine V4.4")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4 German Add-On")) Then
    Call HayExterno("Cheat Engine V4.4 German Add-On")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.3")) Then
    Call HayExterno("Cheat Engine V4.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.2")) Then
    Call HayExterno("Cheat Engine V4.2")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.1.1")) Then
    Call HayExterno("Cheat Engine V4.1.1")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.3")) Then
    Call HayExterno("Cheat Engine V3.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.2")) Then
    Call HayExterno("Cheat Engine V3.2")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.1")) Then
    Call HayExterno("Cheat Engine V3.1")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("danza engine 5.2.150")) Then
    Call HayExterno("danza engine 5.2.150")
ElseIf FindWindow(vbNullString, UCase$("zenx engine")) Then
    Call HayExterno("zenx engine")
ElseIf FindWindow(vbNullString, UCase$("MACROMAKER")) Then
    Call HayExterno("Macro Maker")
ElseIf FindWindow(vbNullString, UCase$("MACREOMAKER - EDIT MACRO")) Then
    Call HayExterno("Macro Maker")
ElseIf FindWindow(vbNullString, UCase$("By Fedex")) Then
    Call HayExterno("Macro Fedex")
ElseIf FindWindow(vbNullString, UCase$("Macro Mage 1.0")) Then
    Call HayExterno("Macro Mage")
ElseIf FindWindow(vbNullString, UCase$("Auto* v0.4 (c) 2001 Pete Powa")) Then
    Call HayExterno("Macro Fisher")
ElseIf FindWindow(vbNullString, UCase$("Kizsada")) Then
    Call HayExterno("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Makro K33")) Then
    Call HayExterno("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Super Saiyan")) Then
    Call HayExterno("El Chit del Geri")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete")) Then
    Call HayExterno("Piringulete")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete 2003")) Then
    Call HayExterno("Piringulete 2003")
ElseIf FindWindow(vbNullString, UCase$("TUKY2005")) Then
    Call HayExterno("Makro Tuky")
End If

End Sub


Private Sub cmdCerrar_Click()
Call Audio.PlayWave(SND_CLICK)
        If MsgBox("¿Esta seguro que desea salir?", vbYesNo + vbQuestion, "Zageth AO") = vbYes Then
            Call SendData("/SALIR")
            End If
End Sub

Private Sub cmdMinimizar_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub cmdMoverHechi_Click(index As Integer)
If hlst.listIndex = -1 Then Exit Sub

Select Case index
Case 0 'subir
    If hlst.listIndex = 0 Then Exit Sub
Case 1 'bajar
    If hlst.listIndex = hlst.ListCount - 1 Then Exit Sub
End Select

Call SendData("DESPHE" & index + 1 & "," & hlst.listIndex + 1)

Select Case index
Case 0 'subir
    hlst.listIndex = hlst.listIndex - 1
Case 1 'bajar
    hlst.listIndex = hlst.listIndex + 1
End Select

End Sub



Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)

End Sub

Private Sub CreateEvent()
     endEvent = DirectX.CreateEvent(Me)
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbKeyRButton Then
    If PuedeUclickear = True Then
    Call UsarItem
    PuedeUclickear = False
    frmMain.timerUclick.Enabled = True
    End If
    End If
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If endEvent Then
        DirectX.DestroyEvent endEvent
    End If
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub FPS_Timer()

If logged And Not frmMain.Visible Then
    Unload frmConnect
    frmMain.Show
End If
    
End Sub

Private Sub lblBlues_Click()
End Sub


Private Sub LblCasco_Click()

End Sub





Private Sub Image1_Click(index As Integer)
  Call Audio.PlayWave(SND_CLICK)

    Select Case index
        Case 0
            '[MatuX] : 01 de Abril del 2002
                Call frmOpciones.Show(vbModeless, frmMain)
            '[END]
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            SendData "ATRI"
            SendData "ESKI"
            SendData "FEST"
            SendData "FAMA"
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        Case 2
            If Not frmGuildLeader.Visible Then _
                Call SendData("GLINFO")
    End Select
End Sub

Private Sub Label5_Click()
frmGameMaster.Visible = True
End Sub

Private Sub Image2_Click()
frmGameMaster.Visible = True
End Sub


Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    SendData "LC" & tX & "," & tY
    SendData "/COMERCIAR"
End Sub

Private Sub mnuNpcDesc_Click()
    SendData "LC" & tX & "," & tY
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub





Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Second_Timer()
    ActualSecond = mid(Time, 7, 2)
    ActualSecond = ActualSecond + 1
    If ActualSecond = LastSecond Then End
    LastSecond = ActualSecond
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            SendData "IT" & Inventario.SelectedItem & "," & 1
        Else
           If Inventario.Amount(Inventario.SelectedItem) > 1 Then
            frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    SendData "AG"
End Sub

Private Sub UsarItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then SendData "USA" & Inventario.SelectedItem
End Sub

Private Sub EquiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        SendData "EQUI" & Inventario.SelectedItem
End Sub
Private Sub cmdLanzar_Click()
    If hlst.List(hlst.listIndex) <> "(None)" Then
        Call SendData("HK" & hlst.listIndex + 1)
        Call SendData("UK" & Magia)
    End If
End Sub
Private Sub CmdInfo_Click()
    Call SendData("INFS" & hlst.listIndex + 1)
End Sub

''''''''''''''''''''''''''''''''''''''
'     OTROS                          '
''''''''''''''''''''''''''''''''''''''

Private Sub DespInv_Click(index As Integer)
    Inventario.ScrollInventory (index = 0)
End Sub

Private Sub Form_Click()

    If Cartel Then Cartel = False

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(MouseBoton, True)
#End If

    If Not Comerciando Then
        Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)

        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                If UsingSkill = 0 Then
                    SendData "LC" & tX & "," & tY
                Else
                    frmMain.MousePointer = vbDefault
                    If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
                    SendData "WLC" & tX & "," & tY & "," & UsingSkill
                    If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If MouseShift = vbLeftButton Then
                Call SendData("/TELEP YO " & UserMap & " " & tX & " " & tY)
        End If
        End If
    End If
    
End Sub

Private Sub Form_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
        Call SendData("/MOV")
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
#End If
        
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) And _
       ((KeyCode >= 65 And KeyCode <= 90) Or _
       (KeyCode >= 48 And KeyCode <= 57)) Then
        
            Select Case KeyCode
                Case vbKeyM:
                    If Not Audio.PlayingMusic Then
                        Musica = True
                        Audio.PlayMIDI CStr(currentMidi) & ".mid"
                    Else
                        Musica = False
                        Audio.StopMidi
                    End If
                Case vbKeyA:
                    Call AgarrarItem
                Case vbKeyE:
                    Call EquiparItem
                Case vbKeyN:
                    Nombres = Not Nombres
                Case vbKeyD
                    Call SendData("UK" & Domar)
                Case vbKeyR:
                    Call SendData("UK" & Robar)
                Case vbKeyS:
                    AddtoRichTextBox frmMain.RecTxt, "Para activar o desactivar el seguro utiliza la tecla '*' (asterisco)", 255, 255, 255, False, False, False
                Case vbKeyZ:
                    Call SendData("/SEGCLAN")
                Case vbKeyO:
                    Call SendData("UK" & Ocultarse)
                Case vbKeyT:
                    Call TirarItem
                Case vbKeyU:
                        Call UsarItem
                Case vbKeyL:
                    If UserPuedeRefrescar Then
                        Call SendData("RPU")
                        UserPuedeRefrescar = False
                        Beep
                    End If
            End Select
        End If
        
        Select Case KeyCode
            Case vbKeyReturn:
           
                If SendCMSTXT.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                End If
            Case vbKeyMultiply:
            Call SendData("/SEG")
            Case vbKeyDelete:
                If SendTxt.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendCMSTXT.Visible = True
                    SendCMSTXT.SetFocus
                End If
            Case vbKeyF2:
                Call SendData("/J1")
                Call SendData("/J2")
            Case vbKeyF3:
                Call SendData("/J3")
                Call SendData("/J4")
           Case vbKeyF1:
              Call DoAccionTecla("F1")
            Case vbKeyF2:
              Call DoAccionTecla("F2")
            Case vbKeyF3:
              Call DoAccionTecla("F3")
            Case vbKeyF4:
              Call DoAccionTecla("F4")
            Case vbKeyF5:
              Call DoAccionTecla("F5")
            Case vbKeyF6:
              Call DoAccionTecla("F6")
            Case vbKeyF7:
              Call DoAccionTecla("F7")
            Case vbKeyF8:
              Call DoAccionTecla("F8")
            Case vbKeyF9:
              Call DoAccionTecla("F9")
            Case vbKeyF10:
              Call DoAccionTecla("F10")
            Case vbKeyF11:
              Call DoAccionTecla("F11")
            Case vbKeyF12:
              Call DoAccionTecla("F12")
              Case vbKeyF12:
                Dim x As Integer
                    Foto.Area = Ventana
                    Foto.Captura
                    For x = 1 To 1000
                         If Not FileExist(App.Path & "/Fotos/" & x & frmOpciones.Combo1.Text, vbNormal) Then Exit For
                    Next
                     Call SavePicture(Foto.Imagen, App.Path & "/Fotos/" & x & frmOpciones.Combo1.Text)
                     Call AddtoRichTextBox(frmMain.RecTxt, "Fotos" & x & frmOpciones.Combo1.Text, 0, 255, 154, True, False, False)
           Case vbKeyControl:
                If (Not UserDescansar) And _
                   (Not UserMeditar) Then
                        SendData "KC"
                End If
   
        End Select
        
End Sub

Private Sub Form_Load()

Call SetWindowLong(RecTxt.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

SendTxt.Visible = False
SendCMSTXT.Visible = False
TiempoActual = GetTickCount()

    frmMain.Picture = LoadPicture(App.Path & "\Graficos\Main.jpg")

    InvEqu.Picture = LoadPicture(App.Path & _
    "\Graficos\Inventario.jpg")
    
   Me.Left = 0
   Me.Top = 0
   
    If AntiEngine.Interval <> 100 Or AntiEngine.Enabled = False Then
        Call CliEditado
    ElseIf AntiExternos.Interval <> 15000 Or AntiExternos.Enabled = False Then
        Call CliEditado
    ElseIf timerUclick.Interval <> 500 Or timerUclick.Enabled = False Then
        Call CliEditado
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MouseX = x
    MouseY = Y
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Image3_Click(index As Integer)
    Select Case index
        Case 0
            Inventario.SelectGold
            If UserGLD > 0 Then
             Call FrmTransferir.Show(vbModeless, frmMain)
            End If
    End Select
End Sub

Private Sub Label1_Click()
    Dim i As Integer
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    Alocados = SkillPoints
    frmSkills3.Puntos.Caption = "Puntos:" & SkillPoints
    frmSkills3.Show , frmMain
End Sub

Private Sub Label4_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Inventario.jpg")

    picInv.Visible = True

    hlst.Visible = False
    CmdInfo.Visible = False
    CmdLanzar.Visible = False
    lblNombre.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub Label7_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Hechizos.jpg")

    picInv.Visible = False
    hlst.Visible = True
    CmdInfo.Visible = True
    CmdLanzar.Visible = True
    lblNombre.Visible = False
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub picInv_DblClick()
    Call UsarItem
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    Else
      If (Not frmComerciar.Visible) And _
         (Not frmSkills3.Visible) And _
         (Not frmMSG.Visible) And _
         (Not frmForo.Visible) And _
         (Not frmEstadisticas.Visible) And _
         (Not frmCantidad.Visible) And _
         (picInv.Visible) Then
            picInv.SetFocus
      End If
    End If
    On Error GoTo 0
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If Left$(stxtbuffer, 1) = "/" Then
            If UCase(Left$(stxtbuffer, 8)) = "/PASSWD " Then
                    Dim j As String
#If SeguridadAlkon Then
                    j = md5.GetMD5String(Right$(stxtbuffer, Len(stxtbuffer) - 8))
                    Call md5.MD5Reset
#Else
                    j = Right$(stxtbuffer, Len(stxtbuffer) - 8)
#End If
                    stxtbuffer = "/PASSWD " & j
             ElseIf UCase$(stxtbuffer) = "/PANELGM" Then
                frmPanelGm.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
            ElseIf UCase$(stxtbuffer) = "/HACERTORNEO" Then
                FrmConsolaTorneo.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
            ElseIf UCase$(stxtbuffer) = "/FUNDARCLAN" Then
                frmEligeAlineacion.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                
                Exit Sub
            End If
            Call SendData(stxtbuffer)
    
       'Shout
        ElseIf Left$(stxtbuffer, 1) = "-" Then
            Call SendData("-" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

ElseIf Left$(stxtbuffer, 1) = "+" Then
Call SendData("+" & Right$(stxtbuffer, Len(stxtbuffer) - 1))


        'Whisper
        ElseIf Left$(stxtbuffer, 1) = "\" Then
            Call SendData("\" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Say
        ElseIf stxtbuffer <> "" Then
            Call SendData(";" & stxtbuffer)

        End If
    

        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call SendData("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub


Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub

#If UsarWrench = 1 Then

Private Sub Socket1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    
    ServerIp = Socket1.PeerAddress
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((mid$(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    
    Second.Enabled = True
    
    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")

    ElseIf EstadoLogin = E_MODO.Normal Then
        Call SendData("gIvEmEvAlcOde")

    ElseIf EstadoLogin = E_MODO.Dados Then
        Call SendData("gIvEmEvAlcOde")

    End If
End Sub

Private Sub Socket1_Disconnect()
    Dim i As Long
    
    LastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False
    
#If SegudidadAlkon Then
    LOGGING = False
    LOGSTRING = False
    LastPressed = 0
    LastMouse = False
    LastAmount = 0
#End If

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0
    LastSecond = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    

    If Not frmCrearPersonaje.Visible Then
       
            frmConnect.Show

        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer
    
    Socket1.Read RD, DataLength
    
    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        'Call LogCustom("HandleData: " & rBuffer(loopc))
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub


#End If

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).CharIndex > 0 Then
        If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
            Dim i As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If charlist(MapData(tX, tY).CharIndex).Nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            
            m.ListaFin
            m.Show , Me

        End If
    End If
End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
            Call UsarItem
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        SendData "LC" & tX & "," & tY
    Case 1 'Comerciar
        Call SendData("LC" & tX & "," & tY)
        Call SendData("/COMERCIAR")
    End Select
End Select
End Sub

Private Sub timerUclick_Timer()
PuedeUclickear = True
frmMain.timerUclick.Enabled = False
End Sub

'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim i As Long
    
    Debug.Print "WInsock Close"
    
    LastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    Debug.Print "Winsock Connect"
    
    ServerIp = Winsock1.RemoteHostIP
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((mid$(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    
    Second.Enabled = True
    
    'If frmCrearPersonaje.Visible Then
    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
    'ElseIf Not frmRecuperar.Visible Then
    ElseIf EstadoLogin = E_MODO.Normal Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.Dados Then
        Call SendData("gIvEmEvAlcOde")
    'Else
    ElseIf EstadoLogin = E_MODO.RecuperarPass Then
        Dim Cmd As String
        Cmd = "PASSRECO" & frmRecuperar.txtNombre.Text & "~" & frmRecuperar.txtCorreo
        'frmMain.Socket1.Write cmd$, Len(cmd$)
        'Call SendData(cmd$)
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer

    Debug.Print "Winsock DataArrival"
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD

    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    LastSecond = 0
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If

    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

#End If

Private Sub Macros_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
 
If Button = vbRightButton Then
    frmMacros.Show vbModeless, frmMain
Else
        Call DoAccionTecla("F" & index)
        Exit Sub
    End If

End Sub

