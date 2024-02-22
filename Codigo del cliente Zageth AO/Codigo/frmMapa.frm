VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mapa Twister AO"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5280
      Width           =   7335
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Unload Me
End Sub
Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\Graficos\Mapa.BMP")
End Sub
