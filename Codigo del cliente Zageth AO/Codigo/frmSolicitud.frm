VERSION 5.00
Begin VB.Form frmGuildSol 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   240
      MouseIcon       =   "frmSolicitud.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enviar"
      Height          =   495
      Left            =   3360
      MouseIcon       =   "frmSolicitud.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   240
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSolicitud.frx":02A4
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmGuildSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Dim CName As String

Private Sub Command1_Click()
Dim f$

f$ = "SOLICITUD" & CName
f$ = f$ & "," & Replace(Replace(text1.Text, ",", ";"), vbCrLf, "º")

Call SendData(f$)

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Public Sub RecieveSolicitud(ByVal GuildName As String)

CName = GuildName

End Sub

