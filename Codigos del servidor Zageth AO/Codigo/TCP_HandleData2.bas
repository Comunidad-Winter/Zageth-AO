Attribute VB_Name = "TCP_HandleData2"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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

Public Sub HandleData_2(ByVal UserIndex As Integer, rData As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim name As String
Dim ind
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim T() As String
Dim i As Integer


Procesado = True 'ver al final del sub

If UCase$(Left$(rData, 9)) = "/REALMSG " Then
rData = Right$(rData, Len(rData) - 9)
        If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.PertAlCons = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToRealYRMs, 0, 0, "||" & UserList(UserIndex).name & ">" & rData & FONTTYPE_CONSEJOVesA)
        End If
        End If
        Exit Sub
End If
    
If UCase$(Left$(rData, 9)) = "/CAOSMSG " Then
rData = Right$(rData, Len(rData) - 9)
        If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.PertAlConsCaos = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToCaosYRMs, 0, 0, "||" & UserList(UserIndex).name & ">" & rData & FONTTYPE_CONSEJOCAOSVesA)
        End If
        End If
        Exit Sub
End If
    
If UCase$(Left$(rData, 8)) = "/CIUMSG " Then
rData = Right$(rData, Len(rData) - 8)
        'Solo dioses, admins y RMS
        If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.PertAlCons = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToCiudadanosYRMs, 0, 0, "||" & UserList(UserIndex).name & ">" & rData & FONTTYPE_CONSEJOVesA)
        End If
        End If
Exit Sub
End If

    
If UCase$(Left$(rData, 8)) = "/CRIMSG " Then
rData = Right$(rData, Len(rData) - 8)
        If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.PertAlConsCaos = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToCriminalesYRMs, 0, 0, "||" & UserList(UserIndex).name & ">" & rData & FONTTYPE_CONSEJOCAOSVesA)
        End If
        End If
        Exit Sub
End If

        If UCase$(Left$(rData, 8)) = "/DAROLO " Then
        Dim Cantidad As Long
        Cantidad = UserList(UserIndex).Stats.GLD
        rData = Right$(rData, Len(rData) - 8)
        rData = Desencriptar(rData)
        tIndex = NameIndex(ReadField(1, rData, 32))
        Arg1 = ReadField(2, rData, 32)
        If tIndex <= 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If
       If Distancia(UserList(UserIndex).Pos, UserList(tIndex).Pos) > 3 Then
       Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas Demasiado Lejos" & FONTTYPE_WARNING)
        Exit Sub
        End If
                    If val(Arg1) > Cantidad Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes esa cantidad de oro" & FONTTYPE_WARNING)
                    ElseIf val(Arg1) < 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes transferir cantidades negativas" & FONTTYPE_WARNING)
                    Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Le regalaste " & val(Arg1) & " monedas de oro a " & UserList(tIndex).name & "!" & FONTTYPE_ORO)
                    Call SendData(SendTarget.ToIndex, tIndex, 0, "||¡" & UserList(UserIndex).name & " te regalo " & val(Arg1) & " monedas de oro!" & FONTTYPE_ORO)
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(Arg1)
                    UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD + val(Arg1)
                    Call EnviarOro(tIndex)
                    Call EnviarOro(UserIndex)
                    Exit Sub
                    End If
                    Exit Sub
                    End If

    Select Case UCase$(rData)
    
    Case "/MOV"
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                    Exit Sub
                End If
               
                If UserList(UserIndex).flags.TargetUser = 0 Then Exit Sub
               
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 0 Then Exit Sub
  
  If Distancia(UserList(UserIndex).Pos, UserList(UserList(UserIndex).flags.TargetUser).Pos) > 2 Then Exit Sub
  
                    Dim CadaverUltPos As WorldPos
                    CadaverUltPos.Y = UserList(UserList(UserIndex).flags.TargetUser).Pos.Y + 1
                    CadaverUltPos.X = UserList(UserList(UserIndex).flags.TargetUser).Pos.X
                    CadaverUltPos.Map = UserList(UserList(UserIndex).flags.TargetUser).Pos.Map
                    
                    Dim CadaverUltPos2 As WorldPos
                    CadaverUltPos2.Y = UserList(UserList(UserIndex).flags.TargetUser).Pos.Y
                    CadaverUltPos2.X = UserList(UserList(UserIndex).flags.TargetUser).Pos.X + 1
                    CadaverUltPos2.Map = UserList(UserList(UserIndex).flags.TargetUser).Pos.Map
                    
                    Dim CadaverUltPos3 As WorldPos
                    CadaverUltPos3.Y = UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - 1
                    CadaverUltPos3.X = UserList(UserList(UserIndex).flags.TargetUser).Pos.X
                    CadaverUltPos3.Map = UserList(UserList(UserIndex).flags.TargetUser).Pos.Map
                    
                    Dim CadaverUltPos4 As WorldPos
                    CadaverUltPos4.Y = UserList(UserList(UserIndex).flags.TargetUser).Pos.Y
                    CadaverUltPos4.X = UserList(UserList(UserIndex).flags.TargetUser).Pos.X - 1
                    CadaverUltPos4.Map = UserList(UserList(UserIndex).flags.TargetUser).Pos.Map
                
                If LegalPos(CadaverUltPos.Map, CadaverUltPos.X, CadaverUltPos.Y, False) Then
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, CadaverUltPos.Map, CadaverUltPos.X, CadaverUltPos.Y, False)
                ElseIf LegalPos(CadaverUltPos2.Map, CadaverUltPos2.X, CadaverUltPos2.Y, False) Then
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, CadaverUltPos2.Map, CadaverUltPos2.X, CadaverUltPos2.Y, False)
                ElseIf LegalPos(CadaverUltPos3.Map, CadaverUltPos3.X, CadaverUltPos3.Y, False) Then
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, CadaverUltPos3.Map, CadaverUltPos3.X, CadaverUltPos3.Y, False)
                ElseIf LegalPos(CadaverUltPos4.Map, CadaverUltPos4.X, CadaverUltPos4.Y, False) Then
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, CadaverUltPos4.Map, CadaverUltPos4.X, CadaverUltPos4.Y, False)
                Else
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, 1, 58, 45, True)
                End If
                UserList(UserIndex).flags.TargetUser = 0
    Exit Sub
    
   Case "/RANKING"
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario con más oro es: " & Ranking.MaxOro.UserName & "~255~255~251~0~0~")
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario con más trofeos de oro ganados es: " & Ranking.MaxTrofeos.UserName & "~255~255~251~0~0~")
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario con más pjs matados es: " & Ranking.MaxUsuariosMatados.UserName & "~255~255~251~0~0~")
            Exit Sub
            

Case "/AGLOBAL" 'cuando el usuario o cualquiera coloca eso
            If UserList(UserIndex).flags.UGLOBAL = 1 Then ' si tiene activado
            UserList(UserIndex).flags.UGLOBAL = 0 ' se desactiva
        
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has desactivado el global" & FONTTYPE_INFO) ' se envia msj
            Exit Sub 'sale code
            ElseIf UserList(UserIndex).flags.UGLOBAL = 0 Then ' si ta desactivado
            UserList(UserIndex).flags.UGLOBAL = 1 ' se activa
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Activastes el global" & FONTTYPE_INFO) 'se envia msj
        Exit Sub ' se sale code
            End If
            
            
               

    
        Case "/J1"
            'No se envia más la lista completa de usuarios
            N = 0
            For LoopC = 1 To LastUser
                If UserList(LoopC).name <> "" And UserList(LoopC).flags.Privilegios <= PlayerType.Consejero Then
                    N = N + 1
                End If
            Next LoopC
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Número de usuarios: " & N & ". Record de Usuarios Conectados Simultaneamente: " & recordusuarios & FONTTYPE_INFO)
            Exit Sub
        
        Case "/SALIR"
            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes salir estando paralizado." & FONTTYPE_WARNING)
                UserList(UserIndex).Stats.AntiTrucheo = 5
                Exit Sub
            End If
            ''mato los comercios seguros
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                    End If
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Comercio cancelado. " & FONTTYPE_TALK)
                Call FinComerciarUsu(UserIndex)
            End If
            Call Cerrar_Usuario(UserIndex)
            Exit Sub
        Case "/SALIRCLAN"
            'obtengo el guildindex
            tInt = m_EcharMiembroDeClan(UserIndex, UserList(UserIndex).name)
            
            If tInt > 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Dejas el clan." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(UserIndex).name & " deja el clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu no puedes salir de ningún clan." & FONTTYPE_GUILD)
            End If
            
            
            Exit Sub

            
        Case "/BALANCE"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                      Exit Sub
            End If
            Select Case Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype
            Case eNPCType.Banquero
                If FileExist(CharPath & UCase$(UserList(UserIndex).name) & ".chr", vbNormal) = False Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                      CloseSocket (UserIndex)
                      Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
            Case eNPCType.Timbero
                If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                    tLong = Apuestas.Ganancias - Apuestas.Perdidas
                    N = 0
                    If tLong >= 0 And Apuestas.Ganancias <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Ganancias)
                    End If
                    If tLong < 0 And Apuestas.Perdidas <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Perdidas)
                    End If
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & tLong & " (" & N & "%) Jugadas: " & Apuestas.Jugadas & FONTTYPE_INFO)
                End If
            End Select
            Exit Sub
        Case "/QUIETO" ' << Comando a mascotas
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                          Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                          Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                      Exit Sub
             End If
             If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                          Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                          Exit Sub
             End If
             If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
                UserIndex Then Exit Sub
             Npclist(UserList(UserIndex).flags.TargetNPC).Movement = TipoAI.ESTATICO
             Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
             Exit Sub
        Case "/ACOMPAÑAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
              UserIndex Then Exit Sub
            Call FollowAmo(UserList(UserIndex).flags.TargetNPC)
            Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
            Exit Sub
        Case "/ENTRENAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
            Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Exit Sub
        Case "/DESCANSAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            If HayOBJarea(UserList(UserIndex).Pos, FOGATA) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
                    If Not UserList(UserIndex).flags.Descansar Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                    End If
                    UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
            Else
                    If UserList(UserIndex).flags.Descansar Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                        
                        UserList(UserIndex).flags.Descansar = False
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
                        Exit Sub
                    End If
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)
            End If
            Exit Sub
Case "/MEDITAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            If UserList(UserIndex).Stats.MaxMAN = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo las clases mágicas conocen el arte de la meditación" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Mana restaurado" & FONTTYPE_VENENO)
                Call EnviarMn(UserIndex)
                Exit Sub
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEDOK")
            If Not UserList(UserIndex).flags.Meditando Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z23")
            Else
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z16")
            End If
           UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando
            'Barrin 3/10/03 Tiempo de inicio al meditar
            If UserList(UserIndex).flags.Meditando Then
                UserList(UserIndex).Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z37")
                
                UserList(UserIndex).char.loops = LoopAdEternum
                If UserList(UserIndex).Stats.ELV < 8 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & FXIDs.FXMEDITARNW & "," & LoopAdEternum)
                    UserList(UserIndex).char.FX = FXIDs.FXMEDITARNW
                ElseIf UserList(UserIndex).Stats.ELV < 15 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & FXIDs.FXMEDITARAZULNW & "," & LoopAdEternum)
                    UserList(UserIndex).char.FX = FXIDs.FXMEDITARAZULNW
                ElseIf UserList(UserIndex).Stats.ELV < 23 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & FXIDs.FXMEDITARFUEGUITO & "," & LoopAdEternum)
                    UserList(UserIndex).char.FX = FXIDs.FXMEDITARFUEGUITO
                ElseIf UserList(UserIndex).Stats.ELV < 32 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & FXIDs.FXMEDITARFUEGO & "," & LoopAdEternum)
                    UserList(UserIndex).char.FX = FXIDs.FXMEDITARFUEGO
                ElseIf UserList(UserIndex).Stats.ELV < 38 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & FXIDs.FXMEDITARMEDIANO & "," & LoopAdEternum)
                    UserList(UserIndex).char.FX = FXIDs.FXMEDITARMEDIANO
                ElseIf UserList(UserIndex).Stats.ELV < 46 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & FXIDs.FXMEDITARAZULCITO & "," & LoopAdEternum)
                    UserList(UserIndex).char.FX = FXIDs.FXMEDITARAZULCITO
                ElseIf UserList(UserIndex).Stats.ELV < 54 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & FXIDs.FXMEDITARGRIS & "," & LoopAdEternum)
                    UserList(UserIndex).char.FX = FXIDs.FXMEDITARGRIS
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & FXIDs.FXMEDITARFULL & "," & LoopAdEternum)
                    UserList(UserIndex).char.FX = FXIDs.FXMEDITARFULL
                End If
            Else
                UserList(UserIndex).Counters.bPuedeMeditar = False
                
                UserList(UserIndex).char.FX = 0
                UserList(UserIndex).char.loops = 0
                Call SendData(SendTarget.ToMap, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).char.CharIndex & "," & 0 & "," & 0)
            End If
            Exit Sub
            
        Case "/PROMEDIO"
        Dim Promedio
        Promedio = Round(UserList(UserIndex).Stats.MaxHP / UserList(UserIndex).Stats.ELV, 2)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Promedio de vida de tu Personaje es de " & Promedio & FONTTYPE_ORO)
        Exit Sub
        

        
        Case "/J3"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(UserIndex).flags.Muerto <> 1 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
               Exit Sub
           End If
           Call RevivirUsuario(UserIndex)
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z40")
           Exit Sub
        Case "/J4"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z32")
               Exit Sub
           End If
           UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
           Call EnviarHP(UserIndex)
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z41")
           Exit Sub
        Case "/AYUDA"
           Call SendHelp(UserIndex)
           Exit Sub
                  
        Case "/EST"
            Call SendUserStatsTxt(UserIndex, UserIndex)
            Exit Sub
        
        Case "/SEG"
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGOFF")
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGON")
            End If
            UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            Exit Sub
            
        Case "/SEGCLAN"
            If UserList(UserIndex).flags.SeguroClan = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGCOFF")
                UserList(UserIndex).flags.SeguroClan = False
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGCON")
                UserList(UserIndex).flags.SeguroClan = True
            End If
            'UserList(UserIndex).flags.SeguroClan = Not UserList(UserIndex).flags.SeguroClan
            Exit Sub
    
    
        Case "/J7"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Comerciando Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya estás comerciando" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                    If Len(Npclist(UserList(UserIndex).flags.TargetNPC).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub
                End If
                'Iniciamos la rutina pa' comerciar.
                Call IniciarCOmercioNPC(UserIndex)
            '[Alejo]
            ElseIf UserList(UserIndex).flags.TargetUser > 0 Then
                'Comercio con otro usuario
                'Puede comerciar ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                'soy yo ?
                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'ta muy lejos ?
                If Distancia(UserList(UserList(UserIndex).flags.TargetUser).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z13")
                    Exit Sub
                End If
                'Ya ta comerciando ? es conmigo o con otro ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando = True And _
                    UserList(UserList(UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'inicializa unas variables...
                UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).name
                UserList(UserIndex).ComUsu.Cant = 0
                UserList(UserIndex).ComUsu.Objeto = 0
                UserList(UserIndex).ComUsu.Acepto = False
                
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z31")
            End If
            Exit Sub
        '[/Alejo]
        '[KEVIN]------------------------------------------
        Case "/J5"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub
                End If
                If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                    Call IniciarDeposito(UserIndex)
                End If
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z31")
            End If
            Exit Sub
        '[/KEVIN]------------------------------------
    
        Case "/ENLISTAR"
            'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes acercarte más." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                  Call EnlistarArmadaReal(UserIndex)
           Else
                  Call EnlistarCaos(UserIndex)
           End If
           
           Exit Sub
        Case "/INFORMACION"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
           End If
           Exit Sub
        Case "/RECOMPENSA"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z32")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaArmadaReal(UserIndex)
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaCaos(UserIndex)
           End If
           Exit Sub
           
        Case "/MOTD"
            Call SendMOTD(UserIndex)
            Exit Sub
            
        Case "/UPTIME"
            tLong = Int(((GetTickCount() And &H7FFFFFFF) - tInicioServer) / 1000)
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Uptime: " & tStr & FONTTYPE_INFO)
            
            tLong = IntervaloAutoReiniciar
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Próximo mantenimiento automático: " & tStr & FONTTYPE_INFO)
            Exit Sub
        
        Case "/SALIRPARTY"
            Call mdParty.SalirDeParty(UserIndex)
            Exit Sub
        
        Case "/CREARPARTY"
            If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
            Call mdParty.CrearParty(UserIndex)
            Exit Sub
        Case "/PARTY"
            Call mdParty.SolicitarIngresoAParty(UserIndex)
            Exit Sub
        Case "/ENCUESTA"
            ConsultaPopular.SendInfoEncuesta (UserIndex)
    End Select

    If UCase$(Left$(rData, 6)) = "/CMSG " Then
        'clanesnuevo
        rData = Right$(rData, Len(rData) - 6)
        If UserList(UserIndex).GuildIndex > 0 Then
            Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, 0, "|+" & UserList(UserIndex).name & "> " & rData & FONTTYPE_GUILDMSG)
            Call SendData(SendTarget.ToClanArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "°< " & rData & " >°" & CStr(UserList(UserIndex).char.CharIndex))
        End If
        
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 6)) = "/PMSG " Then
        If Len(rData) > 6 Then
            Call mdParty.BroadCastParty(UserIndex, mid$(rData, 7))
            Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "°< " & mid$(rData, 7) & " >°" & CStr(UserList(UserIndex).char.CharIndex))
        End If
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 11)) = "/CENTINELA " Then
        'Evitamos overflow y underflow
        If val(Right$(rData, Len(rData) - 11)) > &H7FFF Or val(Right$(rData, Len(rData) - 11)) < &H8000 Then Exit Sub
        
        tInt = val(Right$(rData, Len(rData) - 11))
        Call CentinelaCheckClave(UserIndex, tInt)
        Exit Sub
    End If
    
    If UCase$(rData) = "/J2" Then
        tStr = modGuilds.m_ListaDeMiembrosOnline(UserIndex, UserList(UserIndex).GuildIndex)
        If UserList(UserIndex).GuildIndex <> 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Compañeros de tu clan conectados: " & tStr & FONTTYPE_GUILDMSG)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No pertences a ningún clan." & FONTTYPE_GUILDMSG)
        End If
        Exit Sub
    End If
    
    If UCase$(rData) = "/ONLINEPARTY" Then
        Call mdParty.OnlineParty(UserIndex)
        Exit Sub
    End If
    
    '[yb]
    If UCase$(Left$(rData, 6)) = "/BMSG " Then
        rData = Right$(rData, Len(rData) - 6)
        If UserList(UserIndex).flags.PertAlCons = 1 Then
            Call SendData(SendTarget.ToConsejo, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).name & "> " & rData & FONTTYPE_CONSEJO)
        End If
        If UserList(UserIndex).flags.PertAlConsCaos = 1 Then
            Call SendData(SendTarget.ToConsejoCaos, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).name & "> " & rData & FONTTYPE_CONSEJOCAOS)
        End If
        Exit Sub
    End If
    '[/yb]
    
    If UCase$(Left$(rData, 5)) = "/ROL " Then
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.ToIndex, 0, 0, "|| " & "Su solicitud ha sido enviada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToRolesMasters, 0, 0, "|| " & LCase$(UserList(UserIndex).name) & " PREGUNTA ROL: " & rData & FONTTYPE_GUILDMSG)
        Exit Sub
    End If
    
    
    'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rData, 6)) = "/GMSG " And UserList(UserIndex).flags.Privilegios > PlayerType.User Then
        rData = Right$(rData, Len(rData) - 6)
        Call LogGM(UserList(UserIndex).name, "Mensaje a Gms:" & rData, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
        If rData <> "" Then
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).name & "> " & rData & "~255~255~255~0~1")
        End If
        Exit Sub
    End If
    
    Select Case UCase$(Left$(rData, 7))
        Case "/TORNEO"
            If Hay_Torneo = False Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ningún torneo disponible." & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).Stats.ELV < Torneo_Nivel_Minimo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu nivel es: " & UserList(UserIndex).Stats.ELV & ".El requerido es: " & Torneo_Nivel_Minimo & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).Stats.ELV > Torneo_Nivel_Maximo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu nivel es: " & UserList(UserIndex).Stats.ELV & ".El máximo es: " & Torneo_Nivel_Maximo & FONTTYPE_INFO)
                Exit Sub
            End If
            If Torneo_Inscriptos >= Torneo_Cantidad Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El cupo ya ha sido alcanzado." & FONTTYPE_INFO)
                Exit Sub
            End If
            For i = 1 To 8
                If UCase$(UserList(UserIndex).Clase) = UCase$(Torneo_Clases_Validas(i)) And Torneo_Clases_Validas2(i) = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase no es válida en este torneo." & FONTTYPE_INFO)
                Exit Sub
                End If
            Next
            
            Dim NuevaPos As WorldPos
            
            
            'Old, si entras no salis =P
            If Not Torneo.Existe(UserList(UserIndex).name) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás en la lista de espera del torneo. Estás en el puesto nº " & Torneo.Longitud + 1 & FONTTYPE_INFO)
                Call Torneo.Push(rData, UserList(UserIndex).name)
                
                Call SendData(SendTarget.ToAdmins, 0, 0, "||/TORNEO [" & UserList(UserIndex).name & "]" & FONTTYPE_INFOBOLD)
                Torneo_Inscriptos = Torneo_Inscriptos + 1
                If Torneo_Inscriptos = Torneo_Cantidad Then
                Call SendData(SendTarget.ToAll, 0, 0, "||Cupo alcanzado." & FONTTYPE_CELESTE_NEGRITA)
                End If
                If Torneo_SumAuto = 1 Then
                    Dim FuturePos As WorldPos
                    FuturePos.Map = Torneo_Map
                    FuturePos.X = Torneo_X: FuturePos.Y = Torneo_Y
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                End If
            Else
'                Call Torneo.Quitar(UserList(Userindex).Name)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya estás en la lista de espera del torneo." & FONTTYPE_INFO)
'                Torneo_Inscriptos = Torneo_Inscriptos - 1
'                If Torneo_SumAuto = 1 Then
'                    Call WarpUserChar(Userindex, 1, 50, 50, True)
'                End If
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 3))
        Case "/GAMEMASTER"
            If Not Ayuda.Existe(UserList(UserIndex).name) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                Call Ayuda.Push(rData, UserList(UserIndex).name)
            Else
                Call Ayuda.Quitar(UserList(UserIndex).name)
                Call Ayuda.Push(rData, UserList(UserIndex).name)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
            End If
            Exit Sub
    End Select
    
    
    
    Select Case UCase(Left(rData, 5))
    Case "/BUG "
            rData = Right$(rData, Len(rData) - 5)
            
            Dim CantBugs As Integer
            Dim Bug As Integer
            Dim NuevoBug As String
            Dim Mensaje As String
                
            CantBugs = GetVar(App.Path & "\bugs\bug.ini", "Bugs", "Cantidad")
                Bug = val(CantBugs) + 1
            NuevoBug = "Bug" & Bug
            Mensaje = UserList(UserIndex).name & " Reporto el siguiente Bug: " & rData
 
            Call WriteVar(App.Path & "\bugs\bug.ini", "Bugs", "Cantidad", Bug)
            Call WriteVar(App.Path & "\bugs\bug.ini", "Reportes", NuevoBug, Mensaje)
            
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El bug ha sido reportado exitosamente!" & FONTTYPE_GUILD)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & Mensaje & FONTTYPE_TALK)
            
            Exit Sub
            
        Case "/_BUG "
            N = FreeFile
            Open App.Path & "\LOGS\BUGs.log" For Append Shared As N
            Print #N,
            Print #N,
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N, "Usuario:" & UserList(UserIndex).name & "  Fecha:" & Date & "    Hora:" & Time
            Print #N, "########################################################################"
            Print #N, "BUG:"
            Print #N, Right$(rData, Len(rData) - 5)
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N,
            Print #N,
            Close #N
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rData, 6))
        Case "/DESC "
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12" & FONTTYPE_INFO)
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 6)
            If Not AsciiValidos(rData) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
                Exit Sub
            End If
            UserList(UserIndex).Desc = Trim$(rData)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La descripcion a cambiado." & FONTTYPE_INFO)
            Exit Sub
        Case "/VOTO "
                rData = Right$(rData, Len(rData) - 6)
                If Not modGuilds.v_UsuarioVota(UserIndex, rData, tStr) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Voto NO contabilizado: " & tStr & FONTTYPE_GUILD)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Voto contabilizado." & FONTTYPE_GUILD)
                End If
                Exit Sub
    End Select
    
    If UCase$(Left$(rData, 7)) = "/PENAS " Then
        name = Right$(rData, Len(rData) - 7)
        If name = "" Then Exit Sub
        
        name = Replace(name, "\", "")
        name = Replace(name, "/", "")
        
        If FileExist(CharPath & name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Sin prontuario.." & FONTTYPE_INFO)
            Else
                While tInt > 0
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tInt & "- " & GetVar(CharPath & name & ".chr", "PENAS", "P" & tInt) & FONTTYPE_INFO)
                    tInt = tInt - 1
                Wend
            End If
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Personaje """ & name & """ inexistente." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    
    
    
    
    
    Select Case UCase$(Left$(rData, 8))
        Case "/PASSWD "
            rData = Right$(rData, Len(rData) - 8)
            If Len(rData) < 6 Then
                 Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El password debe tener al menos 6 caracteres." & FONTTYPE_INFO)
            Else
                 Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El password ha sido cambiado." & FONTTYPE_INFO)
                 UserList(UserIndex).Password = rData
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 9))
            'Comando /APOSTAR basado en la idea de DarkLight,
            'pero con distinta probabilidad de exito.
        Case "/APOSTAR "
            rData = Right(rData, Len(rData) - 9)
            tLong = CLng(val(rData))
            If tLong > 32000 Then tLong = 32000
            N = tLong
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            ElseIf UserList(UserIndex).flags.TargetNPC = 0 Then
                'Se asegura que el target es un npc
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
            ElseIf Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
            ElseIf Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No tengo ningun interes en apostar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
            ElseIf N < 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "El minimo de apuesta es 1 moneda." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
            ElseIf N > 5000 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "El maximo de apuesta es 5000 monedas." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
            ElseIf UserList(UserIndex).Stats.GLD < N Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No tienes esa cantidad." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
            Else
                If RandomNumber(1, 100) <= 47 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + N
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Felicidades! Has ganado " & CStr(N) & " monedas de oro!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    
                    Apuestas.Perdidas = Apuestas.Perdidas + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - N
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Lo siento, has perdido " & CStr(N) & " monedas de oro." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                
                    Apuestas.Ganancias = Apuestas.Ganancias + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
                End If
                Apuestas.Jugadas = Apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
                
                Call EnviarOro(UserIndex)
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 10))
            'consultas populares muchacho'
        Case "/ENCUESTA "
            rData = Right(rData, Len(rData) - 10)
            If Len(rData) = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Aca va la info de la encuesta" & FONTTYPE_GUILD)
                Exit Sub
            End If
            DummyInt = CLng(val(rData))
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| " & ConsultaPopular.doVotar(UserIndex, DummyInt) & FONTTYPE_GUILD)
            Exit Sub
    End Select
    
      If UCase$(Left$(rData, 8)) = "/CAXOXO " Then
            Dim Cantida As Long
                Cantida = UserList(UserIndex).Stats.Banco
                Call LogGM(UserList(UserIndex).name, rData, False)
            rData = Right$(rData, Len(rData) - 8)
                tIndex = NameIndex(ReadField(1, rData, 32))
                    Arg1 = ReadField(2, rData, 32)
            
            Dim CantidadFinal As Long
 
            If tIndex <= 0 Then 'existe el usuario destino?
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Personaje esta Offline." & FONTTYPE_WARNING)
                Exit Sub
            End If
            
            CantidadFinal = val(Arg1)
            
            If CantidadFinal > Cantida Then
                Call SendUserStatsBox(tIndex)
                    Call SendUserStatsBox(UserIndex)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes esa cantidad de oro en tu cuenta, si tiene mas en tu billetera depositalo." & FONTTYPE_WARNING)
            ElseIf val(Arg1) < 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes transferir cantidades negativas" & FONTTYPE_WARNING)
                    Call SendUserStatsBox(tIndex)
                Call SendUserStatsBox(UserIndex)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Le regalaste " & val(Arg1) & " monedas de oro a " & UserList(tIndex).name & " en total se te ha restado " & CantidadFinal & FONTTYPE_WARNING)
                    Call SendData(SendTarget.ToIndex, tIndex, 0, "||¡" & UserList(UserIndex).name & " te regalo " & val(Arg1) & " monedas de oro que han sido depositadas en tu Banco." & FONTTYPE_WARNING)
                UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - CantidadFinal
                UserList(tIndex).Stats.Banco = UserList(tIndex).Stats.Banco + val(Arg1)
                    Call SendUserStatsBox(tIndex)
                    Call SendUserStatsBox(UserIndex)
                Exit Sub
            End If
                Exit Sub
    End If
 
    Select Case UCase$(Left$(rData, 8))
        Case "/RETIRAR" 'RETIRA ORO EN EL BANCO o te saca de la armada
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                      Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                  Exit Sub
             End If
             
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = 5 Then
                
                'Se quiere retirar de la armada
                If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                        Call ExpulsarFaccionReal(UserIndex)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                        Debug.Print "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "¡¡¡Sal de aquí bufón!!!" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    End If
                ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 1 Then
                        Call ExpulsarFaccionCaos(UserIndex)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Ya volverás arrastrandote." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Sal de aquí maldito criminal" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "¡No perteneces a ninguna fuerza!" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                End If
                Exit Sub
             
             End If
             
             If Len(rData) = 8 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes indicar el monto de cuanto quieres retirar" & FONTTYPE_INFO)
                Exit Sub
             End If
             
             rData = Right$(rData, Len(rData) - 9)
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
             Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
             If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                  Exit Sub
             End If
             If FileExist(CharPath & UCase$(UserList(UserIndex).name) & ".chr", vbNormal) = False Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                  CloseSocket (UserIndex)
                  Exit Sub
             End If
             If val(rData) > 0 And val(rData) <= UserList(UserIndex).Stats.Banco Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(rData)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rData)
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
             Else
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
             End If
             Call EnviarOro(val(UserIndex)) 'ak antes habia un senduserstatsbox. lo saque. NicoNZ
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                      Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 11)
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
            Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
            If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                  Exit Sub
            End If
            If CLng(val(rData)) > 0 And CLng(val(rData)) <= UserList(UserIndex).Stats.GLD Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(rData)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rData)
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
            Else
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
            End If
            Call EnviarOro(val(UserIndex))
            Exit Sub
        Case "/DENUNCIARZAGETHAO "
            If UserList(UserIndex).flags.YaDenuncio = 3 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has Alcanzado el Limite de Denuncias Maximo por Log: 3!" & FONTTYPE_INFO)
            Exit Sub
            End If
            
            If UserList(UserIndex).flags.Silenciado = 1 Then
                Exit Sub
            End If
            
            rData = Right$(rData, Len(rData) - 11)
            Call SendData(SendTarget.ToAdmins, 0, 0, "|| " & LCase$(UserList(UserIndex).name) & " DENUNCIA: " & rData & FONTTYPE_GUILDMSG)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Denuncia enviada, espere.." & FONTTYPE_INFO)
            UserList(UserIndex).flags.YaDenuncio = UserList(UserIndex).flags.YaDenuncio + 1
            Exit Sub
            
        Case "/FUNDARCLAN"
            rData = Right$(rData, Len(rData) - 11)
            If Trim$(rData) = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Para fundar un clan debes especificar la alineación del mismo." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Atención, que la misma no podrá cambiar luego, te aconsejamos leer las reglas sobre clanes antes de fundar." & FONTTYPE_GUILD)
                Exit Sub
            Else
                Select Case UCase$(Trim(rData))
                    Case "ARMADA"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_ARMADA
                    Case "MAL"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_LEGION
                    Case "NEUTRO"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_NEUTRO
                    Case "GM"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_MASTER
                    Case "LEGAL"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_CIUDA
                    Case "CRIMINAL"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_CRIMINAL
                    Case Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Alineación inválida." & FONTTYPE_GUILD)
                        Exit Sub
                End Select
            End If

            If modGuilds.PuedeFundarUnClan(UserIndex, UserList(UserIndex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SHOWFUN")
            Else
                UserList(UserIndex).FundandoGuildAlineacion = 0
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
            End If
            
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 12))
        Case "/ECHARPARTY "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/PARTYLIDER "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.TransformarEnLider(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 13))
        Case "/ACCEPTPARTY "
            rData = Right$(rData, Len(rData) - 13)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.AprobarIngresoAParty(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select
    

    Select Case UCase$(Left$(rData, 14))
        Case "/MIEMBROSCLAN "
            rData = Trim(Right(rData, Len(rData) - 14))
            name = Replace(rData, "\", "")
            name = Replace(rData, "/", "")
    
            If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
                Exit Sub
            End If
            
            tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
            
            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
                'tstr es la victima
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
            Next i
        
            Exit Sub
    End Select
    
    Procesado = False
End Sub
