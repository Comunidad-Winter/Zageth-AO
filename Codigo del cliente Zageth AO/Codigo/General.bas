Attribute VB_Name = "Mod_General"


Option Explicit

Public bK As Long
Public RandomCode As String

Public iplst As String
Public banners As String

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Public lFrameTimer As Long
Public sHKeys() As String

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

Public Function DirGraficos() As String
    DirGraficos = App.Path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
    DirSound = App.Path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.Path & "\" & Config_Inicio.DirMusica & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.Path & "\" & Config_Inicio.DirMapas & "\"
End Function

Public Function SumaDigitos(ByVal numero As Integer) As Integer
    'Suma digitos
    Do
        SumaDigitos = SumaDigitos + (numero Mod 10)
        numero = numero \ 10
    Loop While (numero > 0)
End Function

Public Function SumaDigitosMenos(ByVal numero As Integer) As Integer
    'Suma digitos, y resta el total de dígitos
    Do
        SumaDigitosMenos = SumaDigitosMenos + (numero Mod 10) - 1
        numero = numero \ 10
    Loop While (numero > 0)
End Function

Public Function Complex(ByVal numero As Integer) As Integer
    If numero Mod 2 <> 0 Then
        Complex = numero * SumaDigitos(numero)
    Else
        Complex = numero * SumaDigitosMenos(numero)
    End If
End Function

Public Function ValidarLoginMSG(ByVal numero As Integer) As Integer
    Dim AuxInteger As Integer
    Dim AuxInteger2 As Integer
    
    AuxInteger = SumaDigitos(numero)
    AuxInteger2 = SumaDigitosMenos(numero)
    ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.Path & "\init\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub CargarVersiones()
On Error GoTo errorH:

    Versiones(1) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Graficos", "Val"))
    Versiones(2) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Wavs", "Val"))
    Versiones(3) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Midis", "Val"))
    Versiones(4) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Init", "Val"))
    Versiones(5) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Mapas", "Val"))
    Versiones(6) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "E", "Val"))
    Versiones(7) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "O", "Val"))
Exit Sub

errorH:
    Call MsgBox("Error cargando versiones")
End Sub

Sub CargarColores()
    Dim archivoC As String
    
    archivoC = App.Path & "\init\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).r = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).G = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).b = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    
    ColoresPJ(50).r = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).G = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).b = CByte(GetVar(archivoC, "CR", "B"))
    ColoresPJ(49).r = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).G = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).b = CByte(GetVar(archivoC, "CI", "B"))
End Sub

#If SeguridadAlkon Then
Sub InitMI()
    Dim alternativos As Integer
    Dim CualMITemp As Integer
    
    alternativos = RandomNumber(1, 7368)
    CualMITemp = RandomNumber(1, 1233)
    

    Set MI(CualMITemp) = New clsManagerInvisibles
    Call MI(CualMITemp).Inicializar(alternativos, 10000)
    
    If CualMI <> 0 Then
        Call MI(CualMITemp).CopyFrom(MI(CualMI))
        Set MI(CualMI) = Nothing
    End If
    CualMI = CualMITemp
End Sub
#End If

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.Path & "\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************
    With RichTextBox
        If (Len(.Text)) > 10000 Then .Text = ""
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        
        .SelBold = Bold
        .SelItalic = Italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        RichTextBox.Refresh
    End With
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).Active = 1 Then
            MapData(charlist(loopc).Pos.x, charlist(loopc).Pos.Y).CharIndex = loopc
        End If
    Next loopc
End Sub

Sub SaveGameini()
    'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.Name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort
    
    Call EscribirGameIni(Config_Inicio)
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    

    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
  
    LegalCharacter = True
End Function

Sub SetConnected()

    Connected = True
    
    Call SaveGameini

    Unload frmConnect
    
    frmMain.Label8.Caption = UserName

    frmMain.Visible = True
Call SetMusicInfo("", "", "Jugando Zageth AO - www.zagethao.es.tl", , "{1}{0}", True)
End Sub


Sub MoveTo(ByVal Direccion As E_Heading)

    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = LegalPos(UserPos.x, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = LegalPos(UserPos.x + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = LegalPos(UserPos.x, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = LegalPos(UserPos.x - 1, UserPos.Y)
    End Select
    
    If LegalOk Then
        If Not UserMeditar And Not UserParalizado Then
        Call SendData("Ñ" & Direccion)
        Call DibujarMiniMapa(frmMain.MiniMap)
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call SendData("CHEA" & Direccion)
        End If
    End If
    
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************

    MoveTo RandomNumber(1, 4)
    
End Sub

Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error Resume Next
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(vbKeyUp) < 0 Then
                Call MoveTo(NORTH)
                Exit Sub
            End If
        
            'Move Right
            If GetKeyState(vbKeyRight) < 0 Then
                Call MoveTo(EAST)
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(vbKeyDown) < 0 Then
                Call MoveTo(SOUTH)
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(vbKeyLeft) < 0 Then
                Call MoveTo(WEST)
                Exit Sub
            End If
        Else
            Dim kp As Boolean
            kp = (GetKeyState(vbKeyUp) < 0) Or _
                GetKeyState(vbKeyRight) < 0 Or _
                GetKeyState(vbKeyDown) < 0 Or _
                GetKeyState(vbKeyLeft) < 0
            If kp Then Call RandomMove
        End If
    End If
End Sub

'TODO : esto no es del tileengine??
Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim x As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
    
        Case E_Heading.EAST
            x = 1
    
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            x = -1
            
    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.Y + Y

    If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.x, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.x, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.x, UserPos.Y).Trigger = 4, True, False)
        Exit Sub
    End If
End Sub

'TODO : esto no es del tileengine??
Function NextOpenChar()
'******************************************
'Finds next open Char
'******************************************
    Dim loopc As Long
    
    loopc = 1
    Do While charlist(loopc).Active And loopc < UBound(charlist)
        loopc = loopc + 1
    Loop
    
    NextOpenChar = loopc
End Function

'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Sub SwitchMap(ByVal Map As Integer)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
    Dim loopc As Long
    Dim Y As Long
    Dim x As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    
    Open DirMapas & "Mapa" & Map & ".map" For Binary As #1
    Seek #1, 1
            
    'map Header
    Get #1, , MapInfo.MapVersion
    Get #1, , MiCabecera
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            Get #1, , ByFlags
            
            MapData(x, Y).Blocked = (ByFlags And 1)
            
            Get #1, , MapData(x, Y).Graphic(1).GrhIndex
            InitGrh MapData(x, Y).Graphic(1), MapData(x, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get #1, , MapData(x, Y).Graphic(2).GrhIndex
                InitGrh MapData(x, Y).Graphic(2), MapData(x, Y).Graphic(2).GrhIndex
            Else
                MapData(x, Y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get #1, , MapData(x, Y).Graphic(3).GrhIndex
                InitGrh MapData(x, Y).Graphic(3), MapData(x, Y).Graphic(3).GrhIndex
            Else
                MapData(x, Y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get #1, , MapData(x, Y).Graphic(4).GrhIndex
                InitGrh MapData(x, Y).Graphic(4), MapData(x, Y).Graphic(4).GrhIndex
            Else
                MapData(x, Y).Graphic(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get #1, , MapData(x, Y).Trigger
            Else
                MapData(x, Y).Trigger = 0
            End If
            
            'Erase NPCs
            If MapData(x, Y).CharIndex > 0 Then
                Call EraseChar(MapData(x, Y).CharIndex)
            End If
            
            'Erase OBJs
            MapData(x, Y).ObjGrh.GrhIndex = 0
        Next x
    Next Y
    
    Close #1
    
    MapInfo.Name = ""
    MapInfo.Music = ""
    
    CurMap = Map
    Call GenerarMiniMapa
Call DibujarMiniMapa(frmMain.MiniMap)

End Sub

'TODO : Reemplazar por la nueva versión, esta apesta!!!
Public Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
    Dim i As Integer
    Dim LastPos As Integer
    Dim CurChar As String * 1
    Dim FieldNum As Integer
    Dim Seperator As String
    
    Seperator = Chr$(SepASCII)
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

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub WriteClientVer()
    Dim hFile As Integer
        
    hFile = FreeFile()
    Open App.Path & "\init\Ver.bin" For Binary Access Write Lock Read As #hFile
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    
    Put #hFile, , CInt(App.Major)
    Put #hFile, , CInt(App.Minor)
    Put #hFile, , CInt(App.Revision)
    
    Close #hFile
End Sub

Public Function IsIp(ByVal Ip As String) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ServersLst)
        If ServersLst(i).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next i
End Function


Public Function CurServerPasRecPort() As Integer
    If CurServer <> 0 Then
        CurServerPasRecPort = 7667
    Else
        CurServerPasRecPort = CInt(frmConnect.PortTxt)
    End If
End Function

Public Function CurServerIp() As String
CurServerIp = "127.0.0.1"
End Function

Public Function CurServerPort() As Integer
CurServerPort = "7666"
End Function


Sub Main()

On Error Resume Next



#If SeguridadAlkon Then
    InitSecurity
#End If

    Call WriteClientVer
    Call LeerLineaComandos
    
    Dim EstaBloqueado As Byte
    EstaBloqueado = Val(GetVar(App.Path & "\init\version.dat", "VERSION", "Graficos"))
    If EstaBloqueado = 1 Then
    Call MsgBox("Tu Cliente ha sido Bloqueado, Consulta a un Game Master para Solucionarlo", vbCritical + vbOKOnly)
    End
    End If
    
    If App.PrevInstance Then
        Call MsgBox("Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If
    
         If Not FileExist(App.Path & "\init\cabezas.ind", vbArchive) Then
        Call MsgBox("ERROR FATAL: sugió un error en los archivos, reinstale el juego e inicielo nuevamente", vbCritical + vbOKOnly)
        End
    End If
    
         If Not FileExist(App.Path & "\init\Ropas.ind", vbArchive) Then
        Call MsgBox("ERROR FATAL: sugió un error en los archivos, reinstale el juego e inicielo nuevamente", vbCritical + vbOKOnly)
        End
    End If
    

Uclickear = True
DialogosClanes.Activo = False
PuedeUclickear = True


Dim f As Boolean
Dim ulttick As Long, esttick As Long
Dim timers(1 To 2) As Integer

 
    
    ChDrive App.Path
    ChDir App.Path

Dim fMD5HushYo As String * 32
    fMD5HushYo = MD5File(App.Path & "\" & App.exeName & ".exe")
    MD5HushYo = txtOffset(hexMd52Asc(fMD5HushYo), 55)
    
    Debug.Print fMD5HushYo
    
    'Cargamos el archivo de configuracion inicial
    If FileExist(App.Path & "\init\Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()
    End If
    
    
    If FileExist(App.Path & "\init\ao.dat", vbArchive) Then
        Call LoadClientSetup
        
        If ClientSetup.bDinamic Then
            Set SurfaceDB = New clsSurfaceManDyn
        Else
            Set SurfaceDB = New clsSurfaceManStatic
        End If
    Else
    
        Set SurfaceDB = New clsSurfaceManDyn
    End If
    
    
    tipf = Config_Inicio.tip
    
    Call frmCargando.establecerProgreso(0)
    
    frmCargando.Show
    frmCargando.Refresh
    

#If UsarWrench = 1 Then
    frmMain.Socket1.Startup
#End If




    Call InicializarNombres
    
    Call frmCargando.progresoConDelay(15)
    Call frmCargando.progresoConDelay(20)
    

       
    IniciarObjetosDirectX
    


Dim loopc As Integer

lastTime = GetTickCount

    Call InitTileEngine(frmMain.hWnd, frmMain.MainViewShp.Top, frmMain.MainViewShp.Left, 32, 32, Round(frmMain.MainViewShp.Height / 32), Round(frmMain.MainViewShp.Width / 32), 9)
    
  
    
    Call CargarAnimsExtra
    Call frmCargando.progresoConDelay(85)
    Call CargarTips

UserMap = 1

    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarVersiones
    Call CargarColores
    Call frmCargando.progresoConDelay(95)
    
    If FileExist(App.Path & "\Init\opciones.zagethao", vbNormal) = False Then
        MSNshow = 1
        Call WriteVar(App.Path & "\Init\opciones.zagethao", "INIT", "Active", "1")
    Else
        MSNshow = GetVar(App.Path & "\Init\opciones.zagethao", "INIT", "Active")
    End If
    
#If SeguridadAlkon Then
    CualMI = 0
    Call InitMI
#End If

  
    
    Unload frmCargando
    

   
    Call Audio.Initialize(DirectX, frmMain.hWnd, App.Path & "\" & Config_Inicio.DirSonidos & "\", App.Path & "\" & Config_Inicio.DirMusica & "\")
    
    Call frmCargando.progresoConDelay(100)
    frmCargando.Visible = False
    Unload frmCargando
    
  
    Call Inventario.Initialize(DirectDraw, frmMain.picInv)
    

    
    frmConnect.Visible = True

    MainViewRect.Left = MainViewLeft
    MainViewRect.Top = MainViewTop
    MainViewRect.Right = MainViewRect.Left + MainViewWidth
    MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
    
    MainDestRect.Left = TilePixelWidth * TileBufferSize - TilePixelWidth
    MainDestRect.Top = TilePixelHeight * TileBufferSize - TilePixelHeight
    MainDestRect.Right = MainDestRect.Left + MainViewWidth
    MainDestRect.Bottom = MainDestRect.Top + MainViewHeight
    
    'Inicialización de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False
    
    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame
            
            'Play ambient sounds
            Call RenderSounds
        End If
        
'TODO : Porque el pausado de 20 ms???
        If GetTickCount - lastTime > 20 Then
            If Not pausa And frmMain.Visible And Not frmForo.Visible And Not frmComerciar.Visible And Not frmComerciarUsu.Visible And Not frmBancoObj.Visible Then
                CheckKeys
                lastTime = GetTickCount
            End If
        End If
        
        'Limitamos los FPS a 18 (con el nuevo engine 60 es un número mucho mejor)
        While (GetTickCount - lFrameTimer) \ 56 < FramesPerSecCounter
            Sleep 5
        Wend
        
        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            FramesPerSec = FramesPerSecCounter
            
            If FPSFLAG Then frmMain.Caption = FramesPerSec
            
            FramesPerSecCounter = 0
            lFrameTimer = GetTickCount
        End If
        
'TODO : Sería mejor comparar el tiempo desde la última vez que se hizo hasta el actual SOLO cuando se precisa. Además evitás el corte de intervalos con 2 golpes seguidos.
        'Sistema de timers renovado:
        esttick = GetTickCount
        For loopc = 1 To UBound(timers)
            timers(loopc) = timers(loopc) + (esttick - ulttick)
            'Timer de trabajo
            If timers(1) >= tUs Then
                timers(1) = 0
                NoPuedeUsar = False
            End If
            'timer de attaque (77)
            If timers(2) >= tAt Then
                timers(2) = 0
                UserCanAttack = 1
                UserPuedeRefrescar = True
            End If
        Next loopc
        ulttick = GetTickCount
        
#If SeguridadAlkon Then
        Call CheckSecurity
#End If
        
        DoEvents
    Loop

    EngineRun = False
    frmCargando.Show
   
    LiberarObjetosDX

'TODO : Esto debería ir en otro lado como al cambair a esta res
    If Not bNoResChange Then
        Dim typDevM As typDevMODE
        Dim lRes As Long
        
        lRes = EnumDisplaySettings(0, 0, typDevM)
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
        End With
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    End If

    'Destruimos los objetos públicos creados
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
#If SeguridadAlkon Then
    Set md5 = Nothing
#End If
    
    Call frmCargando.Show
    
    Call UnloadAllForms
    
    Call frmCargando.establecerProgreso(100)
    
    Call UnloadAllForms
    
    Call frmCargando.progresoConDelay(50)
    
    'Actualizar tip
    Config_Inicio.tip = tipf
    Call EscribirGameIni(Config_Inicio)
    
       Call frmCargando.progresoConDelay(0)


    Call UnloadAllForms
    
#If SeguridadAlkon Then
    DeinitSecurity
#End If
End

ManejadorErrores:
    MsgBox "Ha ocurrido un error irreparable, el cliente se cerrará."
    LogError "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.source
    End
End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lorelativo a mapas, no tiene anda que hacer acá....
Function HayAgua(ByVal x As Integer, ByVal Y As Integer) As Boolean

    HayAgua = MapData(x, Y).Graphic(1).GrhIndex >= 1505 And _
                MapData(x, Y).Graphic(1).GrhIndex <= 1520 And _
                MapData(x, Y).Graphic(2).GrhIndex = 0
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
    End If
End Sub
    
Public Sub LeerLineaComandos()
    Dim T() As String
    Dim i As Long
    
    'Parseo los comandos
    T = Split(Command, " ")
    
    For i = LBound(T) To UBound(T)
        Select Case UCase$(T(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
        End Select
    Next i
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    Open App.Path & "\init\ao.dat" For Binary Access Read Lock Write As fHandle
        Get fHandle, , ClientSetup
    Close fHandle
    
    Musica = Not ClientSetup.bNoMusic
    Sound = Not ClientSetup.bNoSound
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(1) = "Ullathorpe"
    Ciudades(2) = "Nix"
    Ciudades(3) = "Banderbill"

    CityDesc(1) = "Ullathorpe está establecida en el medio de los grandes bosques de Argentum, es principalmente un pueblo de campesinos y leñadores. Su ubicación hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares más legendarios de este mundo."
    CityDesc(2) = "Nix es una gran ciudad. Edificada sobre la costa oeste del principal continente de Argentum."
    CityDesc(3) = "Banderbill se encuentra al norte de Ullathorpe y Nix, es una de las ciudades más importantes de todo el imperio."

    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Elfo"
    ListaRazas(3) = "Elfo Oscuro"
    ListaRazas(4) = "Gnomo"
    ListaRazas(5) = "Enano"

    ListaClases(1) = "Mago"
    ListaClases(2) = "Clerigo"
    ListaClases(3) = "Guerrero"
    ListaClases(4) = "Asesino"
    ListaClases(5) = "Bardo"
    ListaClases(6) = "Paladin"
    ListaClases(7) = "Cazador"

    SkillsNames(Skills.Suerte) = "Suerte"
    SkillsNames(Skills.Magia) = "Magia"
    SkillsNames(Skills.Robar) = "Robar"
    SkillsNames(Skills.Tacticas) = "Tacticas de combate"
    SkillsNames(Skills.Armas) = "Combate con armas"
    SkillsNames(Skills.Meditar) = "Meditar"
    SkillsNames(Skills.Apuñalar) = "Apuñalar"
    SkillsNames(Skills.Ocultarse) = "Ocultarse"
    SkillsNames(Skills.Supervivencia) = "Supervivencia"
    SkillsNames(Skills.Talar) = "Talar árboles"
    SkillsNames(Skills.Comerciar) = "Comercio"
    SkillsNames(Skills.Defensa) = "Defensa con escudos"
    SkillsNames(Skills.Pesca) = "Pesca"
    SkillsNames(Skills.Mineria) = "Mineria"
    SkillsNames(Skills.Carpinteria) = "Carpinteria"
    SkillsNames(Skills.Herreria) = "Herreria"
    SkillsNames(Skills.Liderazgo) = "Liderazgo"
    SkillsNames(Skills.Domar) = "Domar animales"
    SkillsNames(Skills.Proyectiles) = "Armas de proyectiles"
    SkillsNames(Skills.Wresterling) = "Wresterling"
    SkillsNames(Skills.Navegacion) = "Navegacion"

    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"
End Sub


Public Function DoAccionTecla(ByVal Tecla As String)
 
Dim Accion As Byte
    Accion = GetVar(IniPath & "Macros.bin", Tecla, "Accion")
    
    If Accion = 1 Then
        Dim Comando As String
        Comando = GetVar(IniPath & "Macros.bin", Tecla, "Comando")
            Call SendData("/" & Comando)
            
    ElseIf Accion = 2 Then
        Dim Usar As Byte
        Usar = GetVar(IniPath & "Macros.bin", Tecla, "UsarItem")
            Call SendData("USA" & Usar)
            
    ElseIf Accion = 3 Then
        Dim Equipar As Byte
        Equipar = GetVar(IniPath & "Macros.bin", Tecla, "EquiparItem")
            Call SendData("EQUI" & Equipar)
            
    ElseIf Accion = 4 Then
        Dim Hechizo As Byte
        Hechizo = GetVar(IniPath & "Macros.bin", Tecla, "LanzarHechizo")
            Call SendData("HK" & Hechizo)
            Call SendData("UK" & Magia)
            
    ElseIf Accion <> 1 Or 2 Or 3 Or 4 Then
        Call frmMacros.Show(vbModeless, frmMain)
        
    ElseIf Accion = "" Then
        Exit Function
    End If
    
End Function

Public Function DibujarMacros(ByVal Tecla As Integer)
 
Dim Accion As Byte
    Accion = GetVar(IniPath & "Macros.bin", "F" & Tecla, "Accion")
 
    If Accion = 1 Then
        frmMain.Macros(Tecla).Picture = LoadPicture(App.Path & "\Graficos\531.bmp")
        
    ElseIf Accion = 2 Then
        Dim Usar As Byte
            Usar = GetVar(IniPath & "Macros.bin", "F" & Tecla, "UsarItem")
        Dim Grh As Integer
            Grh = Inventario.GrhIndex(Usar)
                Call DibujarMacrosUsarEquipar(Grh, Tecla)
         
    ElseIf Accion = 3 Then
        Dim Equipar As Byte
            Equipar = GetVar(IniPath & "Macros.bin", "F" & Tecla, "EquiparItem")
        Dim Grhs As Integer
            Grhs = Inventario.GrhIndex(Equipar)
                Call DibujarMacrosUsarEquipar(Grhs, Tecla)
                
    ElseIf Accion = 4 Then
        frmMain.Macros(Tecla).Picture = LoadPicture(App.Path & "\Graficos\617.bmp")
        
    ElseIf Accion <> 1 Or 2 Or 3 Or 4 Then
        Exit Function
    End If
End Function
Public Function DibujarMacrosUsarEquipar(ByVal Grh As Integer, ByVal Tecla As Integer)
Dim SR As RECT, DR As RECT
SR.Left = 0
SR.Top = 0
SR.Right = 34
SR.Bottom = 34
DR.Left = 0
DR.Top = 0
DR.Right = 34
DR.Bottom = 34
Call DrawGrhtoHdc(frmMain.Macros(Tecla).hWnd, frmMain.Macros(Tecla).hDC, Grh, SR, DR)
End Function
Public Function CargarMacros()
    Dim i As Byte
        For i = 1 To 12
            Call DibujarMacros(i)
        Next i
End Function

Public Sub DibujarMiniMapa(ByRef Pic As PictureBox)
    Dim DR As RECT
    DR.Left = 0
    DR.Top = 0
    DR.Bottom = 100
    DR.Right = 100
    SupMiniMap.BltFast 1, 1, SupBMiniMap, DR, DDBLTFAST_WAIT
    
    DR.Left = UserPos.x - 4
    DR.Top = UserPos.Y - 4
    DR.Bottom = UserPos.Y - 2
    DR.Right = UserPos.x - 2
    SupMiniMap.BltColorFill DR, &HFFFF00
    
    DR.Left = 0
    DR.Top = 0
    DR.Bottom = 100
    DR.Right = 100
    SupMiniMap.BltToDC Pic.hDC, DR, DR
    frmMain.MiniMap.Refresh
End Sub
 
Public Sub GenerarMiniMapa()
    Dim x As Integer
    Dim Y As Integer
    Dim i As Integer
    Dim DR As RECT
    Dim SR As RECT
    Dim aux As Integer
    
    SR.Left = 0
    SR.Top = 0
    SR.Bottom = 100
    SR.Right = 100
    'SupBMiniMap.BltColorFill SR, vbBlack
    
    For x = MinYBorder To MaxXBorder
        For Y = MinYBorder To MaxYBorder
            If MapData(x, Y).Graphic(1).GrhIndex > 0 Then
                With MapData(x, Y).Graphic(1)
                    i = GrhData(.GrhIndex).Frames(1)
                End With
                
                SR.Left = GrhData(i).sX
                SR.Top = GrhData(i).sY
                SR.Right = GrhData(i).sX + GrhData(i).pixelWidth
                SR.Bottom = GrhData(i).sY + GrhData(i).pixelHeight
                DR.Left = x - 5
                DR.Top = Y - 5
                DR.Bottom = Y - 3
                DR.Right = x - 3
                SupBMiniMap.Blt DR, SurfaceDB.Surface(GrhData(i).FileNum), SR, DDBLT_DONOTWAIT
            End If
            
            If MapData(x, Y).Graphic(2).GrhIndex > 0 Then
                With MapData(x, Y).Graphic(2)
                    i = GrhData(.GrhIndex).Frames(1)
                End With
            
                SR.Left = GrhData(i).sX
                SR.Top = GrhData(i).sY
                SR.Right = GrhData(i).sX + GrhData(i).pixelWidth
                SR.Bottom = GrhData(i).sY + GrhData(i).pixelHeight
                DR.Left = x - 5
                DR.Top = Y - 5
                DR.Bottom = Y - 3
                DR.Right = x - 3
                SupBMiniMap.Blt DR, SurfaceDB.Surface(GrhData(i).FileNum), SR, DDBLT_DONOTWAIT
            End If
            
            If MapData(x, Y).Graphic(3).GrhIndex > 0 Then
                With MapData(x, Y).Graphic(3)
                    i = GrhData(.GrhIndex).Frames(1)
                End With
            
                SR.Left = GrhData(i).sX
                SR.Top = GrhData(i).sY
                SR.Right = GrhData(i).sX + GrhData(i).pixelWidth
                SR.Bottom = GrhData(i).sY + GrhData(i).pixelHeight
                DR.Left = x - 5
                DR.Top = Y - 5
                DR.Bottom = Y - 3
                DR.Right = x - 3
                SupBMiniMap.Blt DR, SurfaceDB.Surface(GrhData(i).FileNum), SR, DDBLT_DONOTWAIT
            End If
            
        Next
    Next
    
End Sub
