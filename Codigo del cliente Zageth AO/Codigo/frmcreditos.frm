VERSION 5.00
Begin VB.Form frmcreditos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   1110
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Subeposs 
      Interval        =   100
      Left            =   480
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   315
      Left            =   4440
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Subepos 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Label creditoss 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Atte: Staff Zageth AO"
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label creditos 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Staff Zageth AO (Agradece a www.gs-zone.org)"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "frmcreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmcreditos.Visible = False
End Sub

Private Sub Subepos_Timer()
creditos.Top = Val(creditos.Top) - 1
End Sub

Private Sub Subeposs_Timer()
creditoss.Top = Val(creditoss.Top) - 1
End Sub
