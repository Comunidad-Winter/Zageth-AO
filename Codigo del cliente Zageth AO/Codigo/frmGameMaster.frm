VERSION 5.00
Begin VB.Form frmGameMaster 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Formulario de Contacto"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd 
      Caption         =   "Enviar"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   6
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Cancelar"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox Texto 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1920
      Width           =   3975
   End
   Begin VB.OptionButton Opcion 
      BackColor       =   &H00808080&
      Caption         =   "Reportar BUG"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.OptionButton Opcion 
      BackColor       =   &H00808080&
      Caption         =   "Denuncia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.OptionButton Opcion 
      BackColor       =   &H00808080&
      Caption         =   "Consulta Regular"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Utiliza este formulario para realizar consultas a los Game Masters, El Abuso de Dichas Consultas sera gravemente penado."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmGameMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            If Opcion(0).value = True Then
                Call SendData("/GAMEMASTER " & Texto.Text)
            ElseIf Opcion(1).value = True Then
                Call SendData("/DENUNCIARZAGETHAO " & Texto.Text)
            ElseIf Opcion(2).value = True Then
                Call SendData("/BUG " & Texto.Text)
            End If
        Exit Sub
    End Select
End Sub



