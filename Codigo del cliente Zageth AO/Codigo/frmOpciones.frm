VERSION 5.00
Begin VB.Form frmOpciones 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4950
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      BackColor       =   &H00808080&
      Caption         =   "Activar ""Mostrar lo que estoy jugando"" (MSN)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   4815
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmOpciones.frx":0152
      Left            =   2760
      List            =   "frmOpciones.frx":015C
      TabIndex        =   5
      Text            =   "Seleccionar"
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSound 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sonidos"
      Height          =   345
      Left            =   1080
      MouseIcon       =   "frmOpciones.frx":016A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   2790
   End
   Begin VB.CommandButton cmdMusica 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Musica"
      Height          =   345
      Left            =   1080
      MouseIcon       =   "frmOpciones.frx":02BC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   2790
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   1080
      MouseIcon       =   "frmOpciones.frx":040E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   2790
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sacar fotos  en"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   105
      Width           =   4935
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAlphaB_Click()
If ConAlfaB = True Then
ConAlfaB = False
cmdAlphaB.Caption = "AlphaBlending Desactivado"
Else
ConAlfaB = True
cmdAlphaB.Caption = "AlphaBlending Activado"
End If
End Sub

Private Sub Check1_Click()
If Check1.value = 1 Then
    Call WriteVar(App.Path & "\Init\opciones.zagethao", "INIT", "Active", "1")
    MSNshow = 1
Else
    Call WriteVar(App.Path & "\Init\opciones.zagethao", "INIT", "Active", "0")
    MSNshow = 0
End If
End Sub

Private Sub cmdCerrar_Click()
Me.Visible = False

End Sub

Private Sub CmdMapa_Click(index As Integer)
Call frmMapa.Show(vbModeless, frmMain)
End Sub



Private Sub cmdMusica_Click()
        If Musica Then
            Musica = False
            cmdMusica.Caption = "Musica Desactivada"
            Audio.StopMidi
        Else
            Musica = True
            cmdMusica.Caption = "Musica Activada"
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
        End If
End Sub
Private Sub cmdSound_Click()
If Sound Then
            Sound = False
            cmdSound.Caption = "Sonidos Desactivados"
            Call Audio.StopWave
            RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone
        Else
            Sound = True
            cmdSound.Caption = "Sonidos Activados"
        End If
End Sub

Private Sub CmdUclick_Click()
If Uclickear = True Then
Uclickear = False
CmdUclick.Caption = "U+Click Boton derecho Desactivado"
Else
Uclickear = True
CmdUclick.Caption = "U+Click Boton derecho Activado"
End If
End Sub

Private Sub creditos_Click()
frmcreditos.Visible = True
End Sub

Private Sub Form_Load()
    If Musica Then
        cmdMusica.Caption = "Musica Activada"
    Else
        cmdMusica.Caption = "Musica Desactivada"
    End If
    
    If Sound Then
        cmdSound.Caption = "Sonidos Activados"
    Else
        cmdSound.Caption = "Sonidos Desactivados"
         End If
   If MSNshow = 1 Then
        Check1.value = 1
    Else
        Check1.value = 0
    End If

End Sub

