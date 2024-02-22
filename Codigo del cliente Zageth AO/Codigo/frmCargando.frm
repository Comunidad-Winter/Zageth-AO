VERSION 5.00
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":0000
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   801
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox imgProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   4080
      Picture         =   "frmCargando.frx":0C42
      ScaleHeight     =   585
      ScaleWidth      =   3720
      TabIndex        =   0
      Top             =   8040
      Width           =   3720
   End
   Begin VB.Image imgProgress2 
      Height          =   600
      Left            =   4080
      Picture         =   "frmCargando.frx":4EB0
      Top             =   8040
      Visible         =   0   'False
      Width           =   3750
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 
Private porcentajeActual As Integer
 
Private Const PROGRESS_DELAY = 10
Private Const PROGRESS_DELAY_BACKWARDS = 4
Private Const DEFAULT_PROGRESS_WIDTH = 250
Private Const DEFAULT_STEP_FORWARD = 1
Private Const DEFAULT_STEP_BACKWARDS = -3
 
Public Sub progresoConDelay(ByVal porcentaje As Integer)
 
If porcentaje = porcentajeActual Then Exit Sub
 
Dim step As Integer, stepInterval As Integer, timer As Long, tickCount As Long
 
If (porcentaje > porcentajeActual) Then
    step = DEFAULT_STEP_FORWARD
    stepInterval = PROGRESS_DELAY
Else
    step = DEFAULT_STEP_BACKWARDS
    stepInterval = PROGRESS_DELAY_BACKWARDS
End If
 
Do Until compararPorcentaje(porcentaje, porcentajeActual, step)
    Do Until (timer + stepInterval) <= GetTickCount()
        DoEvents
    Loop
    timer = GetTickCount()
    porcentajeActual = porcentajeActual + step
    Call establecerProgreso(porcentajeActual)
Loop
 
End Sub
 
 
Public Sub establecerProgreso(ByVal nuevoPorcentaje As Integer)
 
If nuevoPorcentaje >= 0 And nuevoPorcentaje <= 100 Then
    imgProgress.Width = DEFAULT_PROGRESS_WIDTH * CLng(nuevoPorcentaje) / 100
ElseIf nuevoPorcentaje > 100 Then
    imgProgress.Width = DEFAULT_PROGRESS_WIDTH
Else
    imgProgress.Width = 0
End If
porcentajeActual = nuevoPorcentaje
 
End Sub
 
Private Function compararPorcentaje(ByVal porcentajeTarget As Integer, ByVal porcentajeAct As Integer, ByVal step As Integer) As Boolean
 
If step = DEFAULT_STEP_FORWARD Then
    compararPorcentaje = (porcentajeAct >= porcentajeTarget)
Else
    compararPorcentaje = (porcentajeAct <= porcentajeTarget)
End If
 
End Function
 
Private Sub Form_Load()
frmCargando.Picture = LoadPicture(App.Path & "\Graficos\Cargando.jpg")
End Sub
