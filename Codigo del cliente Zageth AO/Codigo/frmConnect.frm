VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11970
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox PasswordTXT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   720
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3240
      Width           =   2250
   End
   Begin VB.TextBox NombreTXT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   720
      TabIndex        =   4
      Top             =   2400
      Width           =   2250
   End
   Begin VB.ListBox lst_servers 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      ItemData        =   "frmConnect.frx":0C4E
      Left            =   -720
      List            =   "frmConnect.frx":0C55
      TabIndex        =   3
      Top             =   8955
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Text            =   "7666"
      Top             =   8970
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   -120
      TabIndex        =   2
      Text            =   "localhost"
      Top             =   8880
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   3720
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   465
      Index           =   0
      Left            =   600
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   1530
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1080
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1800
      TabIndex        =   6
      Top             =   7680
      Width           =   3495
   End
   Begin VB.Image imgServArgentina 
      Height          =   75
      Left            =   4440
      MousePointer    =   99  'Custom
      Top             =   -120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   75
      Left            =   -120
      TabIndex        =   1
      Top             =   9000
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
        frmCargando.Show
        frmCargando.Refresh
       

        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        
        Call WriteVar(App.Path & "\init\version.dat", "VERSION", "Graficos", "0")
        frmCargando.Refresh
        LiberarObjetosDX
        frmCargando.Refresh
        Call UnloadAllForms
End If
End Sub


Private Sub Form_Load()
   
    EngineRun = False
  
    
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next
 PortTxt.Text = Config_Inicio.Puerto
 
 frmConnect.Picture = LoadPicture(App.Path & "\Graficos\Conectar.jpg")
Timer1.Enabled = False



End Sub



Private Sub Image1_Click(index As Integer)

CurServer = 0

PuertoDelServidor = PortTxt
IPdelServidor = "127.0.0.1"

Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0
        
        If Musica Then
            Call Audio.PlayMIDI("7.mid")
        End If
        
    
        EstadoLogin = Dados
#If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
        End If
        frmMain.Socket1.HostName = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
#Else
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
        End If
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
        Me.MousePointer = 11

        
        

End Select
Exit Sub

End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)

#If UsarWrench = 1 Then
            If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
    #Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
    #End If
            If frmConnect.MousePointer = 11 Then
                Exit Sub
            End If
           
           
            UserName = NombreTXT.Text
        Dim aux As String
        aux = PasswordTXT.Text
#If SeguridadAlkon Then
        UserPassword = md5.GetMD5String(aux)
        Call md5.MD5Reset
#Else
        UserPassword = aux
#End If
            If CheckUserData(False) = True Then
                'SendNewChar = False
                EstadoLogin = Normal
                Me.MousePointer = 11
    #If UsarWrench = 1 Then
                frmMain.Socket1.HostName = CurServerIp
                frmMain.Socket1.RemotePort = CurServerPort
                frmMain.Socket1.Connect
    #Else
                If frmMain.Winsock1.State <> sckClosed Then _
                    frmMain.Winsock1.Close
                frmMain.Winsock1.Connect CurServerIp, CurServerPort
    #End If
            End If
End Sub

Private Sub imgGetPass_Click()
On Error GoTo errH

    Call Audio.PlayWave(SND_CLICK)
    Call Shell(App.Path & "\RECUPERAR.EXE", vbNormalFocus)
    'Call frmRecuperar.Show(vbModal, frmConnect)
    Exit Sub
errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")
End Sub

Private Sub Image3_Click()
Call Audio.PlayWave(SND_CLICK)
If MsgBox("¿Esta seguro que desea salir?", vbYesNo + vbQuestion, "Zageth AO") = vbYes Then Call SaveGameini
        Call UnloadAllForms

End Sub

Private Sub imgServArgentina_Click()
    Call Audio.PlayWave(SND_CLICK)
    IPTxt.Text = IPdelServidor
    PortTxt.Text = PuertoDelServidor
End Sub

Private Sub imgServEspana_Click()
    Call Audio.PlayWave(SND_CLICK)
    IPTxt.Text = "62.42.193.233"
    PortTxt.Text = "7666"
End Sub



Private Sub lst_servers_Click()
If ServersRecibidos Then
    CurServer = lst_servers.listIndex + 1
    IPTxt = ServersLst(CurServer).Ip
    PortTxt = ServersLst(CurServer).Puerto
End If

End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Image2_Click
    End If
End Sub

Private Sub Timer1_Timer()
frmConnect.Label1.Caption = " "
frmConnect.Timer1.Enabled = False
frmConnect.MousePointer = 1
End Sub
