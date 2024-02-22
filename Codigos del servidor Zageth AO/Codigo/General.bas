Attribute VB_Name = "General"


Option Explicit

Global LeerNPCs As New clsIniReader
Global LeerNPCsHostiles As New clsIniReader

Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal T As Long, ByVal r As String)

Public Function MD5String(p As String) As String
' compute MD5 digest on a given string, returning the result
    Dim r As String * 32, T As Long
    r = Space(32)
    T = Len(p)
    MDStringFix p, T, r
    MD5String = r
End Function

Public Function MD5File(f As String) As String
' compute MD5 digest on o given file, returning the result
    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
End Function
Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)

Select Case UCase$(UserList(UserIndex).Raza)
    Case "HUMANO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 21
                    Else
                        UserList(UserIndex).char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 39
                    Else
                        UserList(UserIndex).char.Body = 39
                    End If
      End Select
    Case "ELFO OSCURO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 32
                    Else
                        UserList(UserIndex).char.Body = 32
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 40
                    Else
                        UserList(UserIndex).char.Body = 40
                    End If
      End Select
    Case "ENANO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 53
                    Else
                        UserList(UserIndex).char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 60
                    Else
                        UserList(UserIndex).char.Body = 60
                    End If
      End Select
    Case "GNOMO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 53
                    Else
                        UserList(UserIndex).char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 60
                    Else
                        UserList(UserIndex).char.Body = 60
                    End If
      End Select
    Case Else
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 21
                    Else
                        UserList(UserIndex).char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 39
                    Else
                        UserList(UserIndex).char.Body = 39
                    End If
      End Select
    
End Select

UserList(UserIndex).flags.Desnudo = 1

End Sub


Sub Bloquear(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Map As Integer, ByVal X As Integer, ByVal Y As Integer, b As Byte)
'b=1 bloquea el tile en (x,y)
'b=0 desbloquea el tile indicado

Call SendData(sndRoute, sndIndex, sndMap, "BQ" & X & "," & Y & "," & b)

End Sub


Function HayAgua(Map As Integer, X As Integer, Y As Integer) As Boolean

If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
    If MapData(Map, X, Y).Graphic(1) >= 1505 And _
       MapData(Map, X, Y).Graphic(1) <= 1520 And _
       MapData(Map, X, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function

Sub LimpiarObjs()

Call SendData(SendTarget.ToAll, 0, 0, "||Zageth AO> Limpiando mundo..." & FONTTYPE_SERVER)
Dim i As Integer
Dim Y As Integer
Dim X As Integer
Dim tInt As String

For i = 1 To NumMaps
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(i, X, Y).OBJInfo.ObjIndex > 0 Then
                    tInt = ObjData(MapData(i, X, Y).OBJInfo.ObjIndex).OBJType
                    If tInt <> otArboles And tInt <> otPuertas And tInt <> otCONTENEDORES And _
                        tInt <> otCARTELES And tInt <> otFOROS And tInt <> otYacimiento And _
                        tInt <> otTELEPORT And tInt <> otYunque And tInt <> otFragua And _
                        tInt <> otMANCHAS Then
                        Call EraseObj(ToMap, 0, i, MapData(i, X, Y).OBJInfo.Amount, i, X, Y)
                    End If
                End If
            End If
            
        Next X
    Next Y
Next i

Call SendData(SendTarget.ToAll, 0, 0, "||Zageth AO> Limpieza de mundo terminada." & FONTTYPE_SERVER)

End Sub

Sub LimpiarMundo()

On Error Resume Next

Dim i As Integer


For i = 1 To TrashCollector.Count
    Dim d As cGarbage
    Set d = TrashCollector(1)
    Call EraseObj(SendTarget.ToMap, 0, d.Map, 1, d.Map, d.X, d.Y)
    Call TrashCollector.Remove(1)
    Set d = Nothing
Next i

Call SecurityIp.IpSecurityMantenimientoLista



End Sub

Sub LimpiarMundoEntero()
Call SendData(SendTarget.ToAll, 0, 0, "||Zageth AO> Limpiando Mundo." & FONTTYPE_SERVER)
Dim MapaActual As Long
Dim Y As Long
Dim X As Long
 
For MapaActual = 1 To NumMaps
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If MapData(MapaActual, X, Y).OBJInfo.ObjIndex = 378 Then Exit For
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex > 0 And MapData(MapaActual, X, Y).Blocked = 0 Then
                If ItemNoEsDeMapa(MapData(MapaActual, X, Y).OBJInfo.ObjIndex) Then Call EraseObj(SendTarget.ToMap, 0, MapaActual, 10000, MapaActual, X, Y)
            End If
        Next X
    Next Y
Next MapaActual
 
LimpiezaTimerMinutos = TimerCleanWorld
 
Call SendData(SendTarget.ToAll, 0, 0, "||Zageth AO> Limpieza del mundo realizada." & FONTTYPE_SERVER)
End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
Dim k As Integer, SD As String
SD = "SPL" & UBound(SpawnList) & ","

For k = 1 To UBound(SpawnList)
    SD = SD & SpawnList(k).NpcName & ","
Next k

Call SendData(SendTarget.ToIndex, UserIndex, 0, SD)
End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
#If UsarQueSocket = 0 Then

Obj.AddressFamily = AF_INET
Obj.Protocol = IPPROTO_IP
Obj.SocketType = SOCK_STREAM
Obj.Binary = False
Obj.Blocking = False
Obj.BufferSize = 1024
Obj.LocalPort = Port
Obj.backlog = 5
Obj.listen

#End If
End Sub




Sub Main()
On Error Resume Next
Dim f As Date

ChDir App.Path
ChDrive App.Path

Call LoadMotd
Call BanIpCargar

Prision.Map = 67
Libertad.Map = 1

Prision.X = 50
Prision.Y = 50
Libertad.X = 58
Libertad.Y = 45


LastBackup = Format(Now, "Short Time")
Minutos = Format(Now, "Short Time")



ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer
ReDim Parties(1 To MAX_PARTIES) As clsParty
ReDim Guilds(1 To MAX_GUILDS) As clsClan



IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"

ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"

Torneo_Clases_Validas(1) = "Guerrero"
Torneo_Clases_Validas(2) = "Mago"
Torneo_Clases_Validas(3) = "Paladin"
Torneo_Clases_Validas(4) = "Clerigo"
Torneo_Clases_Validas(5) = "Bardo"
Torneo_Clases_Validas(6) = "Asesino"
Torneo_Clases_Validas(7) = "Druida"
Torneo_Clases_Validas(8) = "Cazador"

Torneo_Alineacion_Validas(1) = "Criminal"
Torneo_Alineacion_Validas(2) = "Ciudadano"
Torneo_Alineacion_Validas(3) = "Armada CAOS"
Torneo_Alineacion_Validas(4) = "Armada REAL"

    ListaClases(1) = "Mago"
    ListaClases(2) = "Clerigo"
    ListaClases(3) = "Guerrero"
    ListaClases(4) = "Asesino"
    ListaClases(5) = "Bardo"
    ListaClases(6) = "Paladin"
    ListaClases(7) = "Cazador"

SkillsNames(1) = "Suerte"
SkillsNames(2) = "Magia"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Tacticas de combate"
SkillsNames(5) = "Combate con armas"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apuñalar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar arboles"
SkillsNames(11) = "Comercio"
SkillsNames(12) = "Defensa con escudos"
SkillsNames(13) = "Pesca"
SkillsNames(14) = "Mineria"
SkillsNames(15) = "Carpinteria"
SkillsNames(16) = "Herreria"
SkillsNames(17) = "Liderazgo"
SkillsNames(18) = "Domar animales"
SkillsNames(19) = "Armas de proyectiles"
SkillsNames(20) = "Wresterling"
SkillsNames(21) = "Navegacion"


frmCargando.Show

'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")

frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"

'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents

frmCargando.Label1(2).Caption = "Iniciando Arrays..."

Call LoadGuildsDB


Call CargarSpawnList

frmCargando.Label1(2).Caption = "Cargando Server.ini"

MaxUsers = 0
Call LoadSini
Call LoadOptions
Call CargaApuestas

'*************************************************
Call CargaNpcsDat
'*************************************************

frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
'Call LoadOBJData
Call LoadOBJData
    
frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
Call CargarHechizos
    
    
Call LoadArmasHerreria
Call LoadArmadurasHerreria
Call LoadObjCarpintero

If BootDelBackUp Then
    
    frmCargando.Label1(2).Caption = "Cargando BackUp"
    Call CargarBackUp
Else
    frmCargando.Label1(2).Caption = "Cargando Mapas"
    Call LoadMapData
End If


Call SonidosMapas.LoadSoundMapInfo


'Comentado porque hay worldsave en ese mapa!
'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Dim LoopC As Integer

'Resetea las conexiones de los usuarios
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
Next LoopC

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

With frmMain
    .AutoSave.Enabled = True
    .tPiqueteC.Enabled = True
    .Timer1.Enabled = True
    If ClientsCommandsQueue <> 0 Then
        .CmdExec.Enabled = True
    Else
        .CmdExec.Enabled = False
    End If
    .GameTimer.Enabled = True
    .FX.Enabled = True
    .Auditoria.Enabled = True
    .KillLog.Enabled = True
    .TIMER_AI.Enabled = True
    .npcataca.Enabled = True
End With

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Configuracion de los sockets

Call SecurityIp.InitIpTables(1000)

#If UsarQueSocket = 1 Then

Call IniciaWsApi(frmMain.hWnd)
SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 0 Then

frmCargando.Label1(2).Caption = "Configurando Sockets"

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

Call ConfigListeningSocket(frmMain.Socket1, Puerto)

#ElseIf UsarQueSocket = 2 Then

frmMain.Serv.Iniciar Puerto

#ElseIf UsarQueSocket = 3 Then

frmMain.TCPServ.Encolar True
frmMain.TCPServ.IniciarTabla 1009
frmMain.TCPServ.SetQueueLim 51200
frmMain.TCPServ.Iniciar Puerto

#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿




Unload frmCargando


'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
Close #N

'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

tInicioServer = GetTickCount() And &H7FFFFFFF
Call InicializaEstadisticas
Call ActualizarRanking
End Sub

Function FileExist(ByVal file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
    FileExist = Dir$(file, FileType) <> ""
End Function

Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'All these functions are much faster using the "$" sign
'after the function. This happens for a simple reason:
'The functions return a variant without the $ sign. And
'variants are very slow, you should never use them.

'*****************************************************************
'Devuelve el string del campo
'*****************************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String
  
Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = mid$(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i

FieldNum = FieldNum + 1
If FieldNum = Pos Then
    ReadField = mid$(Text, LastPos + 1)
End If

End Function
Function MapaValido(ByVal Map As Integer) As Boolean
MapaValido = Map >= 1 And Map <= NumMaps
End Function

Sub MostrarNumUsers()

frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers

End Sub


Public Sub LogCriticEvent(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogIndex(ByVal Index As Integer, ByVal Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & Index & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogError(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\errores.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogStatic(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogTarea(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile(1) ' obtenemos un canal
Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:


End Sub


Public Sub LogClanes(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\IP.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub


Public Sub LogDesarrollo(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\desarrollo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub



Public Sub LogGM(Nombre As String, texto As String, Consejero As Boolean)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
If Consejero Then
    Open App.Path & "\logs\consejeros\" & Nombre & ".log" For Append Shared As #nfile
Else
    Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
End If
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub SaveDayStats()
''On Error GoTo errhandler
''
''Dim nfile As Integer
''nfile = FreeFile ' obtenemos un canal
''Open App.Path & "\logs\" & Replace(Date, "/", "-") & ".log" For Append Shared As #nfile
''
''Print #nfile, "<stats>"
''Print #nfile, "<ao>"
''Print #nfile, "<dia>" & Date & "</dia>"
''Print #nfile, "<hora>" & Time & "</hora>"
''Print #nfile, "<segundos_total>" & DayStats.Segundos & "</segundos_total>"
''Print #nfile, "<max_user>" & DayStats.MaxUsuarios & "</max_user>"
''Print #nfile, "</ao>"
''Print #nfile, "</stats>"
''
''
''Close #nfile
Exit Sub

errhandler:

End Sub


Public Sub LogAsesinato(texto As String)
On Error GoTo errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:


End Sub
Public Sub LogHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogCheating(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CH.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, ""
Close #nfile

Exit Sub

errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
Dim Arg As String
Dim i As Integer


For i = 1 To 33

Arg = ReadField(i, cad, 44)

If Arg = "" Then Exit Function

Next i

ValidInputNP = True

End Function


Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

Dim LoopC As Integer
  
#If UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next

ReDim UserList(1 To MaxUsers)

For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
Next LoopC

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call LoadOptions
Call LoadOBJData

Call LoadMapData

Call CargarHechizos

#If UsarQueSocket = 0 Then

'*****************Setup socket
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = val(Puerto)
frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

'Log it
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " servidor reiniciado."
Close #N

'Ocultar

If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

  
End Sub


Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
    
    If MapInfo(UserList(UserIndex).Pos.Map).Zona <> "DUNGEON" Then
        If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 1 And _
           MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 2 And _
           MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
    
End Function


Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)

Dim modifi As Integer

If UserList(UserIndex).Counters.Frio < IntervaloFrio Then
  UserList(UserIndex).Counters.Frio = UserList(UserIndex).Counters.Frio + 1
Else
  If MapInfo(UserList(UserIndex).Pos.Map).Terreno = Nieve Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muriendo de frio, abrigate o moriras!!." & FONTTYPE_INFO)
    modifi = Porcentaje(UserList(UserIndex).Stats.MaxHP, 5)
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - modifi
    If UserList(UserIndex).Stats.MinHP < 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Has muerto de frio!!." & FONTTYPE_INFO)
            UserList(UserIndex).Stats.MinHP = 0
            Call UserDie(UserIndex)
    End If
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ASH" & UserList(UserIndex).Stats.MinHP)
  Else
    modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)
    Call QuitarSta(UserIndex, modifi)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ASS" & UserList(UserIndex).Stats.MinSta)
    'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Has perdido stamina, si no te abrigas rapido perderas toda!!." & FONTTYPE_INFO)
  End If
  
  UserList(UserIndex).Counters.Frio = 0
  
  
End If

End Sub

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Mimetismo < IntervaloInvisible Then
    UserList(UserIndex).Counters.Mimetismo = UserList(UserIndex).Counters.Mimetismo + 1
Else
    'restore old char
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Recuperas tu apariencia normal." & FONTTYPE_INFO)
    
    UserList(UserIndex).char.Body = UserList(UserIndex).CharMimetizado.Body
    UserList(UserIndex).char.Head = UserList(UserIndex).CharMimetizado.Head
    UserList(UserIndex).char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
    UserList(UserIndex).char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
    UserList(UserIndex).char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
        
    
    UserList(UserIndex).Counters.Mimetismo = 0
    UserList(UserIndex).flags.Mimetizado = 0
    Call ChangeUserChar(SendTarget.ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim)
End If
            
End Sub



Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Invisibilidad < IntervaloInvisible Then
    UserList(UserIndex).Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad + 1
Else
    UserList(UserIndex).Counters.Invisibilidad = 0
    UserList(UserIndex).flags.Invisible = 0
    If UserList(UserIndex).flags.Oculto = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z11")
        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).char.CharIndex & ",0")
    End If
End If

End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
Else
    Npclist(NpcIndex).flags.Paralizado = 0
    Npclist(NpcIndex).flags.Inmovilizado = 0
End If

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Ceguera > 0 Then
    UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
Else
    If UserList(UserIndex).flags.Ceguera = 1 Then
        UserList(UserIndex).flags.Ceguera = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "NSEGUE")
    End If
    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).flags.Estupidez = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "NESTUP")
    End If

End If


End Sub


Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Paralisis > 0 Then
    UserList(UserIndex).Counters.Paralisis = UserList(UserIndex).Counters.Paralisis - 1
Else
    UserList(UserIndex).flags.Paralizado = 0
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La Paralisis Desaparece" & FONTTYPE_GUILD)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PARADOK")
End If

End Sub

Public Sub RecStamina(UserIndex As Integer, EnviarStats As Boolean, Intervalo As Integer)

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub


Dim massta As Integer
If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta Then
   If UserList(UserIndex).Counters.STACounter < Intervalo Then
       UserList(UserIndex).Counters.STACounter = UserList(UserIndex).Counters.STACounter + 1
   Else
       EnviarStats = True
       UserList(UserIndex).Counters.STACounter = 0
       massta = RandomNumber(1, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))
       UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta + massta
       If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
        End If
    End If
End If

End Sub

Public Sub EfectoVeneno(UserIndex As Integer, EnviarStats As Boolean)
Dim N As Integer

If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
  UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
Else
  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z35")
  UserList(UserIndex).Counters.Veneno = 0
  N = RandomNumber(1, 5)
  UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - N
  If UserList(UserIndex).Stats.MinHP < 1 Then Call UserDie(UserIndex)
  Call SendData(SendTarget.ToIndex, UserIndex, 0, "ASH" & UserList(UserIndex).Stats.MinHP)
End If

End Sub

Public Sub DuracionPociones(UserIndex As Integer)

'Controla la duracion de las pociones
If UserList(UserIndex).flags.DuracionEfecto > 0 Then
   UserList(UserIndex).flags.DuracionEfecto = UserList(UserIndex).flags.DuracionEfecto - 1
   If UserList(UserIndex).flags.DuracionEfecto = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Los Atributos Fuerza y Agilidad vuelven a su estado Original." & FONTTYPE_GUILD)
        Call EnviarDopa(UserIndex)
        UserList(UserIndex).flags.TomoPocion = False
        UserList(UserIndex).flags.TipoPocion = 0
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
              UserList(UserIndex).Stats.UserAtributos(loopX) = UserList(UserIndex).Stats.UserAtributosBackUP(loopX)
        Next
   End If
End If

End Sub

Public Sub Sanar(UserIndex As Integer, EnviarStats As Boolean, Intervalo As Integer)

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub
       

Dim mashit As Integer
'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP Then
   If UserList(UserIndex).Counters.HPCounter < Intervalo Then
      UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
   Else
      mashit = RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))
                           
      UserList(UserIndex).Counters.HPCounter = 0
      UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + mashit
      If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z36")
      EnviarStats = True
    End If
End If

End Sub

Public Sub CargaNpcsDat()
'Dim NpcFile As String
'
'NpcFile = DatPath & "NPCs.dat"
'ANpc = INICarga(NpcFile)
'Call INIConf(ANpc, 0, "", 0)
'
'NpcFile = DatPath & "NPCs-HOSTILES.dat"
'Anpc_host = INICarga(NpcFile)
'Call INIConf(Anpc_host, 0, "", 0)

Dim npcfile As String

npcfile = DatPath & "NPCs.dat"
Call LeerNPCs.Initialize(npcfile)

npcfile = DatPath & "NPCs-HOSTILES.dat"
Call LeerNPCsHostiles.Initialize(npcfile)

End Sub

Public Sub DescargaNpcsDat()
'If ANpc <> 0 Then Call INIDescarga(ANpc)
'If Anpc_host <> 0 Then Call INIDescarga(Anpc_host)

End Sub

Sub PasarSegundo()
    Dim i As Integer
    For i = 1 To LastUser
        'Cerrar usuario
        If UserList(i).Counters.Saliendo Then
            UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
            If UserList(i).Counters.Salir <= 0 Then
                'If NumUsers <> 0 Then NumUsers = NumUsers - 1

                Call SendData(SendTarget.ToIndex, i, 0, "||Gracias por jugar Argentum Online" & FONTTYPE_INFO)
                Call SendData(SendTarget.ToIndex, i, 0, "FINOC")
                
                Call CloseSocket(i)
                Exit Sub
            End If
        
        'ANTIEMPOLLOS
        ElseIf UserList(i).flags.EstaEmpo = 1 Then
             UserList(i).EmpoCont = UserList(i).EmpoCont + 1
             If UserList(i).EmpoCont = 30 Then
                 
                 'If FileExist(CharPath & UserList(Z).Name & ".chr", vbNormal) Then
                 'esto siempre existe! sino no estaria logueado ;p
                 
                 'TmpP = val(GetVar(CharPath & UserList(Z).Name & ".chr", "PENAS", "Cant"))
                 'Call WriteVar(CharPath & UserList(Z).Name & ".chr", "PENAS", "Cant", TmpP + 1)
                 'Call WriteVar(CharPath & UserList(Z).Name & ".chr", "PENAS", "P" & TmpP + 1, LCase$(UserList(Z).Name) & ": CARCEL " & 30 & "m, MOTIVO: Empollando" & " " & Date & " " & Time)

                 'Call Encarcelar(Z, 30, "El sistema anti empollo")
                 Call SendData(SendTarget.ToIndex, i, 0, "!! Fuiste expulsado por permanecer muerto sobre un item")
                 'Call SendData(SendTarget.ToAdmins, Z, 0, "|| " & UserList(Z).Name & " Fue encarcelado por empollar" & FONTTYPE_INFO)
                 UserList(i).EmpoCont = 0
                 Call CloseSocket(i)
                 Exit Sub
             ElseIf UserList(i).EmpoCont = 15 Then
                 Call SendData(SendTarget.ToIndex, i, 0, "|| LLevas 15 segundos bloqueando el item, muévete o serás desconectado." & FONTTYPE_WARNING)
             End If
         End If
Next i
    


    If CuentaRegresiva > 0 Then
        If CuentaRegresiva > 1 Then
            Call SendData(SendTarget.ToAll, 0, 0, "||Contando..." & CuentaRegresiva - 1 & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToAll, 0, 0, "||YA!!!!!!!!!" & "~255~0~0~1~1")
        End If
        CuentaRegresiva = CuentaRegresiva - 1
    End If
    
    Dim UserIndex As Integer
    For UserIndex = 1 To LastUser
    If UserList(UserIndex).Stats.AntiTrucheo < 60 Then
    UserList(UserIndex).Stats.AntiTrucheo = UserList(UserIndex).Stats.AntiTrucheo + 1
    End If
Next

End Sub
 
Public Function ReiniciarAutoUpdate() As Double

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    'WorldSave
    Call DoBackUp

    'commit experiencias
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub

 
Sub GuardarUsuarios()
    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, 0, "BKW")
    Call SendData(SendTarget.ToAll, 0, 0, "||Zageth AO> Grabando Personajes" & FONTTYPE_SERVER)
    
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, CharPath & UCase$(UserList(i).name) & ".chr")
        End If
    Next i
    
    Call SendData(SendTarget.ToAll, 0, 0, "||Zageth AO> Personajes Grabados" & FONTTYPE_SERVER)
    Call SendData(SendTarget.ToAll, 0, 0, "BKW")

    haciendoBK = False
End Sub


Sub InicializaEstadisticas()
Dim Ta As Long
Ta = GetTickCount() And &H7FFFFFFF

Call EstadisticasWeb.Inicializa(frmMain.hWnd)
Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)

End Sub

Public Sub SwapObjects(ByVal UserIndex As Integer)
Dim tmpUserObj As UserOBJ
 
    With UserList(UserIndex)
               
        'Cambiamos si alguno es una herramienta
        If .Invent.HerramientaEqpSlot = ObjSlot1 Then
            .Invent.HerramientaEqpSlot = ObjSlot2
        ElseIf .Invent.HerramientaEqpSlot = ObjSlot2 Then
            .Invent.HerramientaEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un armor
        If .Invent.ArmourEqpSlot = ObjSlot1 Then
            .Invent.ArmourEqpSlot = ObjSlot2
        ElseIf .Invent.ArmourEqpSlot = ObjSlot2 Then
            .Invent.ArmourEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un barco
        If .Invent.BarcoSlot = ObjSlot1 Then
            .Invent.BarcoSlot = ObjSlot2
        ElseIf .Invent.BarcoSlot = ObjSlot2 Then
            .Invent.BarcoSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un casco
        If .Invent.CascoEqpSlot = ObjSlot1 Then
            .Invent.CascoEqpSlot = ObjSlot2
        ElseIf .Invent.CascoEqpSlot = ObjSlot2 Then
            .Invent.CascoEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un escudo
        If .Invent.EscudoEqpSlot = ObjSlot1 Then
            .Invent.EscudoEqpSlot = ObjSlot2
        ElseIf .Invent.EscudoEqpSlot = ObjSlot2 Then
            .Invent.EscudoEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es munición
        If .Invent.MunicionEqpSlot = ObjSlot1 Then
            .Invent.MunicionEqpSlot = ObjSlot2
        ElseIf .Invent.MunicionEqpSlot = ObjSlot2 Then
            .Invent.MunicionEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un arma
        If .Invent.WeaponEqpSlot = ObjSlot1 Then
            .Invent.WeaponEqpSlot = ObjSlot2
        ElseIf .Invent.WeaponEqpSlot = ObjSlot2 Then
            .Invent.WeaponEqpSlot = ObjSlot1
        End If
       
        'Hacemos el intercambio propiamente dicho
        tmpUserObj = .Invent.Object(ObjSlot1)
        .Invent.Object(ObjSlot1) = .Invent.Object(ObjSlot2)
        .Invent.Object(ObjSlot2) = tmpUserObj
 
        'Actualizamos los 2 slots que cambiamos solamente
        Call UpdateUserInv(False, UserIndex, ObjSlot1)
        Call UpdateUserInv(False, UserIndex, ObjSlot2)
    End With
End Sub


Sub ActualizarRanking()
    Dim EsUsuario As Boolean
    Dim Dioses() As String
    Dim Administradores() As String
    Dim SemiDioses() As String
    Dim Consejeros() As String
    Dim CantDioses As Integer
    Dim CantAdministradores As Integer
    Dim CantSemiDioses As Integer
    Dim CantConsejeros As Integer
    Dim NombreUsuario As String
    
    Dim Linea As String
    Dim file As String
    
    Dim MaxOro As Long
    Dim ActualOro As Long
    Dim ActualTrofOro As Integer
    Dim MaxTrofOro As Integer
    Dim ActualUsuariosMatados As Integer
    Dim ActualCiudasMatados As Integer
    Dim ActualCrimisMatados As Integer
    Dim MaxUsuariosMatados As Integer
    Dim ActualUserMatados As Integer
    
    
    Dim i As Integer
    
    CantDioses = 0
    MaxOro = 0
    
    
     Open (Trim(App.Path & "\Server.ini")) For Input As #1
        Do While Not (EOF(1))
            Line Input #1, Linea
            If mid(Linea, 1, 4) = "Dios" And mid(Linea, 6, 1) = "=" Then
                ReDim Preserve Dioses(CantDioses)
                Dioses(CantDioses) = mid(Linea, 7, Len(Linea) - 6)
                CantDioses = CantDioses + 1
            End If
            If mid(Linea, 1, 5) = "Admin" And mid(Linea, 7, 1) = "=" Then
                ReDim Preserve Administradores(CantAdministradores)
                Administradores(CantAdministradores) = mid(Linea, 8, Len(Linea) - 7)
                CantAdministradores = CantAdministradores + 1
            End If
            If mid(Linea, 1, 8) = "Semidios" And mid(Linea, 10, 1) = "=" Then
                ReDim Preserve SemiDioses(CantSemiDioses)
                SemiDioses(CantSemiDioses) = mid(Linea, 11, Len(Linea) - 10)
                CantSemiDioses = CantSemiDioses + 1
            End If
            If mid(Linea, 1, 9) = "Consejero" And mid(Linea, 11, 1) = "=" Then
                ReDim Preserve Consejeros(CantConsejeros)
                Consejeros(CantConsejeros) = mid(Linea, 12, Len(Linea) - 11)
                CantConsejeros = CantConsejeros + 1
            End If
        
        Loop
    Close #1
 
 
    file = Dir(Trim(App.Path & "\Charfile\"))
    Do While (Len(file) <> 0)
        EsUsuario = True
        NombreUsuario = mid(file, 1, Len(file) - 4)
        For i = 0 To CantDioses - 1
            If UCase(NombreUsuario) = UCase(Dioses(i)) Then
                EsUsuario = False
            End If
        Next
        For i = 0 To CantAdministradores - 1
            If UCase(NombreUsuario) = UCase(Administradores(i)) Then
                EsUsuario = False
            End If
        Next
        For i = 0 To CantSemiDioses - 1
            If UCase(NombreUsuario) = UCase(SemiDioses(i)) Then
                EsUsuario = False
            End If
        Next
         For i = 0 To CantConsejeros - 1
            If UCase(NombreUsuario) = UCase(Consejeros(i)) Then
                EsUsuario = False
            End If
        Next
        
        
        If EsUsuario = True Then
            Open (Trim(App.Path & "\Charfile\") & file) For Input As #1
             Do While Not (EOF(1))
                Line Input #1, Linea
                If mid(Linea, 1, 3) = "GLD" Then
                    ActualOro = CLng(mid(Linea, 5, Len(Linea) - 4))
                    If MaxOro < ActualOro Then
                        MaxOro = ActualOro
                        Ranking.MaxOro.value = MaxOro
                        Ranking.MaxOro.UserName = NombreUsuario
                    End If
                End If
                    
                If mid(Linea, 1, 7) = "TrofOro" Then
                    ActualTrofOro = CInt(mid(Linea, 9, Len(Linea) - 8))
                    If MaxTrofOro < ActualTrofOro Then
                        MaxTrofOro = ActualTrofOro
                        Ranking.MaxTrofeos.value = MaxTrofOro
                        Ranking.MaxTrofeos.UserName = NombreUsuario
                    End If
                End If
                    
                If mid(Linea, 1, 11) = "UserMuertes" Then
                    ActualUserMatados = CInt(mid(Linea, 13, Len(Linea) - 12))
                    If MaxUsuariosMatados < ActualUserMatados Then
                        MaxUsuariosMatados = ActualUserMatados
                        Ranking.MaxUsuariosMatados.value = MaxUsuariosMatados
                        Ranking.MaxUsuariosMatados.UserName = NombreUsuario
                    End If
                End If
             Loop
            Close #1
        End If
        file = Dir()
    Loop
    
End Sub
