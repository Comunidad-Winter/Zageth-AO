VERSION 5.00
Object = "{B370EF78-425C-11D1-9A28-004033CA9316}#2.0#0"; "CAPTURA.ocx"
Begin VB.Form FrmProcesos 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Captura.wndCaptura Foto 
      Left            =   0
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      Area            =   2
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Foto"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3540
      ItemData        =   "FrmProcesos.frx":0000
      Left            =   360
      List            =   "FrmProcesos.frx":0002
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "FrmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim x As Integer
Foto.Area = Ventana
Foto.Captura
For x = 1 To 1000
If Not FileExist(App.Path & "/Procesos/" & x & ".bmp", vbNormal) Then Exit For
Next
Call SavePicture(Foto.Imagen, App.Path & "/Procesos/" & x & ".bmp")
End Sub
