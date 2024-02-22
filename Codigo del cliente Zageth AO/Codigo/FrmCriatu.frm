VERSION 5.00
Begin VB.Form FrmCriatura 
   BorderStyle     =   0  'None
   Caption         =   "Estadistica de Criatura"
   ClientHeight    =   3255
   ClientLeft      =   240
   ClientTop       =   4995
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmCriatu.frx":0000
   ScaleHeight     =   3255
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   135
      Left            =   1920
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Ataque 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "124 / 457"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Ataque"
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ataque"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Evasion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "58"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1200
      TabIndex        =   13
      ToolTipText     =   "Evasión"
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label asd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Evación"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Hit 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "20"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Hit"
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Defensa 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "58"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1200
      TabIndex        =   10
      ToolTipText     =   "Defensa"
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hit"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Defensa"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label GLD 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "10000000"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Monedas de Oro"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Exp 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "9999999999"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Experiencia"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Salud 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   " 10000000 / 10000000"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Salud"
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Namex 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Dragón de las Tinieblas"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Criatura"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Experiencia"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos de salud"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Monedas de oro"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la criatura"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FrmCriatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Unload FrmCriatura
End Sub
