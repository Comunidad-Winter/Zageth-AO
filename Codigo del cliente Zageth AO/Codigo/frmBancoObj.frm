VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBancoObj.frx":0000
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3930
      Index           =   0
      Left            =   795
      MouseIcon       =   "frmBancoObj.frx":0C42
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2595
      Width           =   2460
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3930
      Index           =   1
      Left            =   3750
      MouseIcon       =   "frmBancoObj.frx":190C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2595
      Width           =   2460
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   930
      ScaleHeight     =   465
      ScaleWidth      =   450
      TabIndex        =   1
      Top             =   1620
      Width           =   450
   End
   Begin VB.TextBox cantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3255
      TabIndex        =   0
      Text            =   "1"
      Top             =   6975
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   1575
      TabIndex        =   7
      Top             =   1560
      Width           =   2940
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   1575
      TabIndex        =   6
      Top             =   1980
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   1575
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   435
      Index           =   0
      Left            =   600
      MouseIcon       =   "frmBancoObj.frx":25D6
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6840
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   435
      Index           =   1
      Left            =   4245
      MouseIcon       =   "frmBancoObj.frx":32A0
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6840
      Width           =   2160
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   6495
      MouseIcon       =   "frmBancoObj.frx":3F6A
      MousePointer    =   99  'Custom
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   5535
      TabIndex        =   4
      Top             =   1560
      Width           =   60
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Public LastIndex1 As Integer
Public LastIndex2 As Integer




Private Sub cantidad_Change()
If Val(cantidad.Text) < 0 Then
    cantidad.Text = 1
End If

If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
    cantidad.Text = 1
End If

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Command2_Click()
SendData ("FINBAN")
End Sub



Private Sub Form_Deactivate()
'Me.SetFocus
End Sub


Private Sub Form_Load()
'Cargamos la interfase
Me.Picture = LoadPicture(App.Path & "\Graficos\Comerciar.bmp")
Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónComprar.bmp")
Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botónvender.bmp")

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónComprar.bmp")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botónvender.bmp")
    Image1(1).Tag = 1
End If
End Sub

Private Sub Image1_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

If List1(Index).List(List1(Index).listIndex) = "Nada" Or _
   List1(Index).listIndex < 0 Then Exit Sub

Select Case Index
    Case 0
        frmBancoObj.List1(0).SetFocus
        LastIndex1 = List1(0).listIndex
        
        SendData ("RETI" & "," & List1(0).listIndex + 1 & "," & cantidad.Text)
        
   Case 1
        LastIndex2 = List1(1).listIndex
        If Not Inventario.Equipped(List1(1).listIndex + 1) Then
            SendData ("DEPO" & "," & List1(1).listIndex + 1 & "," & cantidad.Text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No podes depositar el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If
                
End Select
List1(0).Clear

List1(1).Clear

NPCInvDim = 0
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        If Image1(0).Tag = 1 Then
                Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónComprarApretado.bmp")
                Image1(0).Tag = 0
                Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botónvender.bmp")
                Image1(1).Tag = 1
        End If
        
    Case 1
        If Image1(1).Tag = 1 Then
                Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botónvenderapretado.bmp")
                Image1(1).Tag = 0
                Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónComprar.bmp")
                Image1(0).Tag = 1
        End If
        
End Select
End Sub

Private Sub Image2_Click()
SendData ("FINBAN")
End Sub

Private Sub list1_Click(Index As Integer)
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

Select Case Index
    Case 0
        Label1(0).Caption = UserBancoInventory(List1(0).listIndex + 1).Name
        Label1(2).Caption = UserBancoInventory(List1(0).listIndex + 1).Amount
        Select Case UserBancoInventory(List1(0).listIndex + 1).OBJType
            Case 2
                Label1(3).Caption = "Max Golpe:" & UserBancoInventory(List1(0).listIndex + 1).MaxHit
                Label1(4).Caption = "Min Golpe:" & UserBancoInventory(List1(0).listIndex + 1).MinHit
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3, 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & UserBancoInventory(List1(0).listIndex + 1).Def
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
        Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hDC, UserBancoInventory(List1(0).listIndex + 1).GrhIndex, SR, DR)
    Case 1
        Label1(0).Caption = Inventario.ItemName(List1(1).listIndex + 1)
        Label1(2).Caption = Inventario.Amount(List1(1).listIndex + 1)
        Select Case Inventario.OBJType(List1(1).listIndex + 1)
            Case 2
                Label1(3).Caption = "Max Golpe:" & Inventario.MaxHit(List1(1).listIndex + 1)
                Label1(4).Caption = "Min Golpe:" & Inventario.MinHit(List1(1).listIndex + 1)
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3, 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & Inventario.Def(List1(1).listIndex + 1)
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
        Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hDC, Inventario.GrhIndex(List1(1).listIndex + 1), SR, DR)
End Select
Picture1.Refresh

End Sub
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónComprar.bmp")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botónvender.bmp")
    Image1(1).Tag = 1
End If
End Sub

