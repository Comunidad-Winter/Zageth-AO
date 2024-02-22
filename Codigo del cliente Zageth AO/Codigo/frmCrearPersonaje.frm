VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Crear Personaje"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   8040
      TabIndex        =   41
      Top             =   6720
      Width           =   3570
   End
   Begin VB.TextBox txtPasswdCheck 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   8040
      PasswordChar    =   "*"
      TabIndex        =   35
      Top             =   3000
      Width           =   3570
   End
   Begin VB.TextBox txtPasswd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   8040
      PasswordChar    =   "*"
      TabIndex        =   34
      Top             =   2070
      Width           =   3570
   End
   Begin VB.TextBox txtCorreoCheck 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   8040
      TabIndex        =   33
      Top             =   4875
      Width           =   3570
   End
   Begin VB.TextBox txtCorreo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   360
      Left            =   8040
      TabIndex        =   32
      Top             =   3930
      Width           =   3570
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
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
      ItemData        =   "frmCrearPersonaje.frx":0C42
      Left            =   1680
      List            =   "frmCrearPersonaje.frx":0C79
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   5760
      Width           =   1470
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
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
      ItemData        =   "frmCrearPersonaje.frx":0D13
      Left            =   1680
      List            =   "frmCrearPersonaje.frx":0D1D
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   6960
      Width           =   1470
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
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
      ItemData        =   "frmCrearPersonaje.frx":0D30
      Left            =   1680
      List            =   "frmCrearPersonaje.frx":0D43
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   6360
      Width           =   1470
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
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
      ItemData        =   "frmCrearPersonaje.frx":0D70
      Left            =   1680
      List            =   "frmCrearPersonaje.frx":0D77
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   7560
      Width           =   1470
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   330
      Left            =   8040
      TabIndex        =   0
      Top             =   1200
      Width           =   3570
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   8040
      TabIndex        =   42
      Top             =   5760
      Width           =   3615
   End
   Begin VB.Label modConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   40
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label modCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   39
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label modInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   38
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label modAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   37
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label modfuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   36
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image command1 
      Height          =   315
      Index           =   41
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":0D87
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   40
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":0ED9
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   225
   End
   Begin VB.Image command1 
      Height          =   315
      Index           =   39
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":102B
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   38
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":117D
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   225
   End
   Begin VB.Image command1 
      Height          =   315
      Index           =   37
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":12CF
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   36
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":1421
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   225
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   35
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":1573
      MousePointer    =   99  'Custom
      Top             =   6645
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   34
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":16C5
      MousePointer    =   99  'Custom
      Top             =   6600
      Width           =   225
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   33
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":1817
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   32
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":1969
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   225
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   31
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":1ABB
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   30
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":1C0D
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   29
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":1D5F
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   28
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":1EB1
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   26
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":2003
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   24
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":2155
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   22
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":22A7
      MousePointer    =   99  'Custom
      Top             =   4440
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   20
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":23F9
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   18
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":254B
      MousePointer    =   99  'Custom
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   16
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":269D
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   360
      Index           =   14
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":27EF
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   375
      Index           =   12
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":2941
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   375
      Index           =   10
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":2A93
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   375
      Index           =   8
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":2BE5
      MousePointer    =   99  'Custom
      Top             =   1920
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   375
      Index           =   6
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":2D37
      MousePointer    =   99  'Custom
      Top             =   1560
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   4
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":2E89
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   2
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":2FDB
      MousePointer    =   99  'Custom
      Top             =   960
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   0
      Left            =   7320
      MouseIcon       =   "frmCrearPersonaje.frx":312D
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   1
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":327F
      MousePointer    =   99  'Custom
      Top             =   600
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   27
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":33D1
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":3523
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":3675
      MousePointer    =   99  'Custom
      Top             =   4440
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   21
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":37C7
      MousePointer    =   99  'Custom
      Top             =   4155
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   19
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":3919
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   17
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":3A6B
      MousePointer    =   99  'Custom
      Top             =   3360
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   15
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":3BBD
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   13
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":3D0F
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":3E61
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   9
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":3FB3
      MousePointer    =   99  'Custom
      Top             =   2040
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   7
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":4105
      MousePointer    =   99  'Custom
      Top             =   1680
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   195
      Index           =   5
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":4257
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   6240
      MouseIcon       =   "frmCrearPersonaje.frx":43A9
      MousePointer    =   99  'Custom
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   31
      Top             =   8400
      Width           =   495
   End
   Begin VB.Image boton 
      Height          =   1485
      Index           =   2
      Left            =   1200
      MouseIcon       =   "frmCrearPersonaje.frx":44FB
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   1380
   End
   Begin VB.Image boton 
      Height          =   735
      Index           =   1
      Left            =   840
      MouseIcon       =   "frmCrearPersonaje.frx":464D
      MousePointer    =   99  'Custom
      Top             =   8160
      Width           =   2115
   End
   Begin VB.Image boton 
      Height          =   690
      Index           =   0
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":479F
      MousePointer    =   99  'Custom
      Top             =   7920
      Width           =   3000
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   20
      Left            =   6600
      TabIndex        =   27
      Top             =   7680
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   19
      Left            =   6600
      TabIndex        =   26
      Top             =   7320
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   18
      Left            =   6600
      TabIndex        =   25
      Top             =   6960
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   17
      Left            =   6600
      TabIndex        =   24
      Top             =   6600
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   16
      Left            =   6600
      TabIndex        =   23
      Top             =   6240
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   15
      Left            =   6600
      TabIndex        =   22
      Top             =   5880
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   14
      Left            =   6600
      TabIndex        =   21
      Top             =   5520
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   13
      Left            =   6600
      TabIndex        =   20
      Top             =   5160
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   12
      Left            =   6600
      TabIndex        =   19
      Top             =   4800
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   6600
      TabIndex        =   18
      Top             =   4440
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   6600
      TabIndex        =   17
      Top             =   4080
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   6600
      TabIndex        =   16
      Top             =   3720
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   6600
      TabIndex        =   15
      Top             =   3360
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   6600
      TabIndex        =   14
      Top             =   3000
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   13
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   6600
      TabIndex        =   12
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   6600
      TabIndex        =   11
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   6600
      TabIndex        =   10
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   9
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   8
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   7
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2520
      TabIndex        =   5
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2520
      TabIndex        =   4
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2520
      TabIndex        =   3
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lbFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2520
      TabIndex        =   1
      Top             =   1320
      Width           =   375
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private AntiKey As String

Public SkillPoints As Byte

Function CheckData() As Boolean
If UserRaza = "" Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = "" Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserHogar = "" Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True


End Function

Private Sub boton_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        
        Dim i As Integer
        Dim k As Object
        i = 1
        For Each k In Skill
            UserSkills(i) = k.Caption
            i = i + 1
        Next
        
        UserName = txtNombre.Text
        
        If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.List(lstRaza.listIndex)
        UserSexo = lstGenero.List(lstGenero.listIndex)
        UserClase = lstProfesion.List(lstProfesion.listIndex)
        
        UserAtributos(1) = Val(lbFuerza.Caption)
        UserAtributos(2) = Val(lbInteligencia.Caption)
        UserAtributos(3) = Val(lbAgilidad.Caption)
        UserAtributos(4) = Val(lbCarisma.Caption)
        UserAtributos(5) = Val(lbConstitucion.Caption)
        
        UserHogar = lstHogar.List(lstHogar.listIndex)
        
If CheckDatos() Then
#If SeguridadAlkon Then
    UserPassword = md5.GetMD5String(txtPasswd.Text)
    Call md5.MD5Reset
#Else
    UserPassword = txtPasswd.Text
#End If
    UserEmail = txtCorreo.Text
    
    If Not CheckMailString(UserEmail) Then
            MsgBox "Direccion de mail invalida."
            Exit Sub
    End If
    
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
#End If

    'SendNewChar = True
    EstadoLogin = CrearNuevoPj
    
    Me.MousePointer = 11

    EstadoLogin = CrearNuevoPj

#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State <> sckConnected Then
#End If
        MsgBox "Error: Se ha perdido la conexion con el server."
        Unload Me
        
    Else
        Call login(RandomCode)
    End If
End If
        
    Case 1
frmMain.Socket1.Disconnect
frmMain.Socket1.Cleanup
frmConnect.MousePointer = 1
Musica = False
Audio.StopMidi
        
        frmConnect.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
        Me.Visible = False
        
        
    Case 2
        Call Audio.PlayWave(SND_DICE)
        Call TirarDados
      
End Select


End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function


Private Sub TirarDados()

#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State = sckConnected Then
#End If
        Call SendData("TIRDAD")
    End If

End Sub

Private Sub Command1_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

Dim indice
If Index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
Call GenerateKey
SkillPoints = 10
puntos.Caption = SkillPoints
Me.Picture = LoadPicture(App.Path & "\graficos\CrearPersonaje.jpg")


Dim i As Integer
lstProfesion.Clear
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstProfesion.listIndex = 1

Call TirarDados
End Sub

Private Sub lstRaza_Change()
Select Case (lstRaza.List(lstRaza.listIndex))
    Case Is = "Humano"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = "+ 2"
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = ""
        modCarisma.Caption = ""
    Case Is = "Elfo"
        modfuerza.Caption = ""
        modConstitucion.Caption = "+ 1"
        modAgilidad.Caption = "+ 3"
        modInteligencia.Caption = "+ 1"
        modCarisma.Caption = "+ 2"
    Case Is = "Elfo Oscuro"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = "+ 2"
        modCarisma.Caption = "- 3"
    Case Is = "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+ 3"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "- 6"
        modCarisma.Caption = "- 3"
    Case Is = "Gnomo"
        modfuerza.Caption = "- 5"
        modAgilidad.Caption = "+ 4"
        modInteligencia.Caption = "+ 3"
        modCarisma.Caption = "+ 1"
End Select
End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
MsgBox "Sea cuidadoso al seleccionar el nombre de su personaje, Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Function CheckDatos() As Boolean

If txtPasswd.Text <> txtPasswdCheck.Text Then
    MsgBox "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
    Exit Function
End If

If text1.Text <> AntiKey Then
    MsgBox "El captcha ingresado es incorrecto."
    Call GenerateKey
    Exit Function
End If

If txtCorreo.Text <> txtCorreoCheck.Text Then
    MsgBox "Los Mails que tipeo no coinciden, por favor vuelva a ingresarlos."
    Exit Function
End If

CheckDatos = True

End Function

Private Function GenerateKey() As String
AntiKey = RandomNumber(62001, 99898) & Chr(RandomNumber(59, 95))
Label7.Caption = AntiKey
End Function
