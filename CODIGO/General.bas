Attribute VB_Name = "Mod_General"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
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

Public iplst As String

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Private lFrameTimer As Long

Public Function DirGraficos() As String
    DirGraficos = App.path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
    DirSound = App.path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.path & "\" & Config_Inicio.DirMusica & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.path & "\" & Config_Inicio.DirMapas & "\"
End Function

Public Function DirExtras() As String
    DirExtras = App.path & "\EXTRAS\"
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Function GetRawName(ByRef sName As String) As String
'***************************************************
'Author: ZaMa
'Last Modify Date: 13/01/2010
'Last Modified By: -
'Returns the char name without the clan name (if it has it).
'***************************************************

    Dim Pos As Integer
    
    Pos = InStr(1, sName, "<")
    
    If Pos > 0 Then
        GetRawName = Trim(Left(sName, Pos - 1))
    Else
        GetRawName = sName
    End If

End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim LoopC As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For LoopC = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(LoopC).WeaponWalk(1), Val(GetVar(arch, "ARMA" & LoopC, "Dir1")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(2), Val(GetVar(arch, "ARMA" & LoopC, "Dir2")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(3), Val(GetVar(arch, "ARMA" & LoopC, "Dir3")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(4), Val(GetVar(arch, "ARMA" & LoopC, "Dir4")), 0
    Next LoopC
End Sub

Sub CargarColores()
On Error Resume Next
    Dim archivoC As String
    
    archivoC = App.path & "\init\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).r = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).g = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).b = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    
    ' Crimi
    ColoresPJ(50).r = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).g = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).b = CByte(GetVar(archivoC, "CR", "B"))
    
    ' Ciuda
    ColoresPJ(49).r = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).g = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).b = CByte(GetVar(archivoC, "CI", "B"))
    
    ' Atacable
    ColoresPJ(48).r = CByte(GetVar(archivoC, "AT", "R"))
    ColoresPJ(48).g = CByte(GetVar(archivoC, "AT", "G"))
    ColoresPJ(48).b = CByte(GetVar(archivoC, "AT", "B"))
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

    Dim LoopC As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For LoopC = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(LoopC).ShieldWalk(1), Val(GetVar(arch, "ESC" & LoopC, "Dir1")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(2), Val(GetVar(arch, "ESC" & LoopC, "Dir2")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(3), Val(GetVar(arch, "ESC" & LoopC, "Dir3")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(4), Val(GetVar(arch, "ESC" & LoopC, "Dir4")), 0
    Next LoopC
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = True)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************r
    With RichTextBox
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
        
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
    Dim LoopC As Long
    
    For LoopC = 1 To LastChar
        If charlist(LoopC).Active = 1 Then
            MapData(charlist(LoopC).Pos.X, charlist(LoopC).Pos.Y).CharIndex = LoopC
        End If
    Next LoopC
End Sub

Sub SaveGameini()
    'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.name = "BetaTester"
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
    Dim LoopC As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For LoopC = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, LoopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next LoopC
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For LoopC = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, LoopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next LoopC
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

#If SeguridadAlkon Then
    Call UnprotectForm
#End If

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
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    Call SaveGameini
    
    frmLogin.Visible = False
    frmPersonaje.Visible = False
    If frmStart.Visible = True Then Unload frmStart
    
    FPSFLAG = True

    Call IniClient
End Sub


Sub MoveTo(ByVal Direccion As E_Heading)

    Dim LegalOk As Boolean
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If LegalOk And charlist(UserCharIndex).Inmo = 0 Then
        Call WriteWalk(Direccion)
        MoveCharbyHead UserCharIndex, Direccion
        MoveScreen Direccion
    Else
        If charlist(UserCharIndex).heading <> Direccion Then
            charlist(UserCharIndex).heading = Direccion
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call MoveTo(RandomNumber(NORTH, WEST))
End Sub

Private Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
    Static LastMovement As Long
    
    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If MirandoForo Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Or EnCuenta Then Exit Sub
    
    'TODO: Debería informarle por consola?
    If Traveling Then Exit Sub
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                Call MoveTo(NORTH)
                charlist(1).Quieto = 0
                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                Call MoveTo(EAST)
                charlist(1).Quieto = 0
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                Call MoveTo(SOUTH)
                charlist(1).Quieto = 0
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                Call MoveTo(WEST)
                charlist(1).Quieto = 0
                Exit Sub
            End If
            
            ' We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(UserPos.X, UserPos.Y)
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                Call Audio.MoveListener(UserPos.X, UserPos.Y)
            End If

        End If
    End If
End Sub

'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Sub SwitchMap(ByVal Map As Integer)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
    Dim Y As Long
    Dim X As Long
    Dim TempInt As Integer
    Dim ByFlags As Byte
    Dim Handle As Integer
    
    Handle = FreeFile()
    
    Open DirMapas & "Mapa" & Map & ".map" For Binary As Handle
    Seek Handle, 1
            
    'map Header
    Get Handle, , MapInfo.MapVersion
    Get Handle, , MiCabecera
    Get Handle, , TempInt
    Get Handle, , TempInt
    Get Handle, , TempInt
    Get Handle, , TempInt
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            Get Handle, , ByFlags
            
            MapData(X, Y).Blocked = (ByFlags And 1)
            
            Get Handle, , MapData(X, Y).Graphic(1).GrhIndex
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get Handle, , MapData(X, Y).Graphic(2).GrhIndex
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
            Else
                MapData(X, Y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get Handle, , MapData(X, Y).Graphic(3).GrhIndex
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
            Else
                MapData(X, Y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get Handle, , MapData(X, Y).Graphic(4).GrhIndex
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
            Else
                MapData(X, Y).Graphic(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get Handle, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0
            End If
            
            'Erase NPCs
            'If MapData(X, Y).CharIndex > 0 Then
            '    Call EraseChar(MapData(X, Y).CharIndex)
            'End If
            
            'Erase OBJs
            MapData(X, Y).ObjGrh.GrhIndex = 0
        Next X
    Next Y
    
    Close Handle
    
    MapInfo.name = ""
    MapInfo.Music = ""
    
    CurMap = Map
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub Main()

    PathBalance = App.path & "\Balance\"
    
    'Load config file
    If FileExist(App.path & "\init\Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()
    End If
    
    'Load ao.dat config file
    Call LoadClientSetup
    
    If ClientSetup.bDinamic Then
        Set SurfaceDB = New clsSurfaceManDyn
    Else
        Set SurfaceDB = New clsSurfaceManStatic
    End If
    
    'Read command line. Do it AFTER config file is loaded to prevent this from
    'canceling the effects of "/nores" option.
    Call LeerLineaComandos
    
    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.path & "\")
    
    ChDrive App.path
    ChDir App.path

#If SeguridadAlkon Then
    'Obtener el HushMD5
    Dim fMD5HushYo As String * 32
    
    fMD5HushYo = md5.GetMD5File(App.path & "\" & App.EXEName & ".exe")
    Call md5.MD5Reset
    MD5HushYo = txtOffset(hexMd52Asc(fMD5HushYo), 55)
    
    Debug.Print fMD5HushYo
#Else
    MD5HushYo = "0123456789abcdef"  'We aren't using a real MD5
#End If
    
    tipf = Config_Inicio.tip
    
    'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
    Call Resolution.SetResolution
    
    ' Mouse Pointer (Loaded before opening any form with buttons in it)
    If FileExist(DirExtras & "Hand.ico", vbArchive) Then _
        Set picMouseIcon = LoadPicture(DirExtras & "Hand.ico")

'TODO : esto de ServerRecibidos no se podría sacar???
    ServersRecibidos = True
    Call InicializarNombres
    
    If Not InitTileEngine(frmMain.hwnd, 149, 19, 32, 32, 13, 17, 9, 8, 8, 0.018) Then
        Call CloseClient
    End If
UserMap = 1
    
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    'Inicializamos el sonido
    Call Audio.Initialize(DirectX, frmMain.hwnd, App.path & "\" & Config_Inicio.DirSonidos & "\", App.path & "\" & Config_Inicio.DirMusica & "\")
    'Enable / Disable audio
    Audio.MusicActivated = Not ClientSetup.bNoMusic
    Audio.SoundActivated = Not ClientSetup.bNoSound
    Audio.SoundEffectsActivated = Not ClientSetup.bNoSoundEffects
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(DirectDraw, frmMain.picInv, MAX_INVENTORY_SLOTS)
    
    Call Audio.MusicMP3Play(App.path & "\MP3\" & MP3_Inicio & ".mp3")

    frmMain.Socket1.Startup
    
    Call Protocol.InitFonts
    
    'Inicialización de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Lan, INT_LAN)

   'Init timers
    Call MainTimer.Start(TimersIndex.Lan)

    'Set the dialog's font
    Dialogos.Font = frmMain.Font

    lFrameTimer = GetTickCount
    
    Dim i As Byte
    For i = 1 To MaxUsers
        charlist(i).ConnID = -1
        charlist(i).ConnIDValida = False
        Set charlist(i).incomingData = New clsByteQueue
        Set charlist(i).outgoingData = New clsByteQueue
    Next i
    
    ReDim Team1(1 To 1) As Integer
    ReDim Team2(1 To 1) As Integer
    
    frmLogin.Visible = True
        
    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            'Play ambient sounds
            Call RenderSounds
            
            Call CheckKeys
        End If

#If SeguridadAlkon Then
        Call CheckSecurity
#End If
        
        ' If there is anything to be sent, we send it
        Call FlushBuffer
        
        DoEvents
    Loop
    
    Call CloseClient
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
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
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

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 343 And MapData(X, Y).Graphic(1).GrhIndex <= 405))
                
End Function

Public Sub ShowSendTxt()
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
End Sub

' Checks the command line parameters, if you are running Ao with /nores command and checks the AoUpdate parameters
'
'

Public Sub LeerLineaComandos()
'*************************************************
'Author: Unknown
'Last modified: 25/11/2008 (BrianPr)
'
'*************************************************
    Dim T() As String
    Dim i As Long
    
    Dim UpToDate As Boolean
    Dim Patch As String
    
    'Parseo los comandos
    T = Split(Command, " ")
    For i = LBound(T) To UBound(T)
        Select Case UCase$(T(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
            Case "/UPTODATE"
                UpToDate = True
        End Select
    Next i
    
    'Call AoUpdate(UpToDate, NoRes) ' www.gs-zone.org
End Sub

''
' Runs AoUpdate if we haven't updated yet, patches aoupdate and runs Client normally if we are updated.
'
' @param UpToDate Specifies if we have checked for updates or not
' @param NoREs Specifies if we have to set nores arg when running the client once again (if the AoUpdate is executed).

Private Sub AoUpdate(ByVal UpToDate As Boolean, ByVal NoRes As Boolean)
'*************************************************
'Author: BrianPr
'Created: 25/11/2008
'Last modified: 25/11/2008
'
'*************************************************
On Error GoTo error
    Dim extraArgs As String
    If Not UpToDate Then
        'No recibe update, ejecutar AU
        'Ejecuto el AoUpdate, sino me voy
        If Dir(App.path & "\AoUpdate.exe", vbArchive) = vbNullString Then
            MsgBox "No se encuentra el archivo de actualización AoUpdate.exe por favor descarguelo y vuelva a intentar", vbCritical
            End
        Else
            FileCopy App.path & "\AoUpdate.exe", App.path & "\AoUpdateTMP.exe"
            
            If NoRes Then
                extraArgs = " /nores"
            End If
            
            Call ShellExecute(0, "Open", App.path & "\AoUpdateTMP.exe", App.EXEName & ".exe" & extraArgs, App.path, SW_SHOWNORMAL)
            End
        End If
    Else
        If FileExist(App.path & "\AoUpdateTMP.exe", vbArchive) Then Kill App.path & "\AoUpdateTMP.exe"
    End If
Exit Sub

error:
    If Err.Number = 75 Then 'Si el archivo AoUpdateTMP.exe está en uso, entonces esperamos 5 ms y volvemos a intentarlo hasta que nos deje.
        Sleep 5
        Resume
    Else
        MsgBox Err.Description & vbCrLf, vbInformation, "[ " & Err.Number & " ]" & " Error "
        End
    End If
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/19/09
'11/19/09: Pato - Is optional show the frmGuildNews form
'**************************************************************
    Dim fHandle As Integer
    
    If FileExist(App.path & "\init\ao.dat", vbArchive) Then
        fHandle = FreeFile
        
        Open App.path & "\init\ao.dat" For Binary Access Read Lock Write As fHandle
            Get fHandle, , ClientSetup
        Close fHandle
    Else
        'Use dynamic by default
        ClientSetup.bDinamic = True
    End If
    
    NoRes = ClientSetup.bNoRes
    
    If InStr(1, ClientSetup.sGraficos, "Graficos") Then
        GraphicsFile = ClientSetup.sGraficos
    Else
        GraphicsFile = "Graficos3.ind"
    End If
    
    ClientSetup.bGuildNews = Not ClientSetup.bGuildNews

End Sub

Private Sub SaveClientSetup()
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 03/11/10
'
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    
    ClientSetup.bNoMusic = Not Audio.MusicActivated
    ClientSetup.bNoSound = Not Audio.SoundActivated
    ClientSetup.bNoSoundEffects = Not Audio.SoundEffectsActivated
    ClientSetup.bGuildNews = Not ClientSetup.bGuildNews

    Open App.path & "\init\ao.dat" For Binary As fHandle
        Put fHandle, , ClientSetup
    Close fHandle
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cArghal) = "Arghâl"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eBotAi.Mago) = "Mago"
    ListaClases(eBotAi.Clerigo) = "Clerigo"
    ListaClases(eBotAi.Guerrero) = "Guerrero"
    ListaClases(eBotAi.Asesino) = "Asesino"
    ListaClases(eBotAi.Pirata) = "Pirata"
    ListaClases(eBotAi.Bardo) = "Bardo"
    ListaClases(eBotAi.Druida) = "Druida"
    ListaClases(eBotAi.Bandido) = "Bandido"
    ListaClases(eBotAi.Paladin) = "Paladin"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasión en combate"
    SkillsNames(eSkill.Armas) = "Combate cuerpo a cuerpo"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar árboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    
    Call Dialogos.RemoveAllDialogs
End Sub

Public Sub CloseClient()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 8/14/2007
'Frees all used resources, cleans up and leaves
'**************************************************************
    Dim i As Byte
    
    EngineRun = False

    Call Resolution.ResetResolution
    
    'Stop tile engine
    Call DeinitTileEngine
    
    Call SaveClientSetup
    
    'Destruimos los objetos públicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    'For i = 1 To MaxUsers
        'Set CharTimer(i) = Nothing
    'Next i
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
#If SeguridadAlkon Then
    Set md5 = Nothing
#End If
    
    Call UnloadAllForms
    
    'Actualizar tip
    Config_Inicio.tip = tipf
    Call EscribirGameIni(Config_Inicio)
    End
End Sub

Public Function esGM(CharIndex As Integer) As Boolean
esGM = False
If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or charlist(CharIndex).priv = 25 Then _
    esGM = True

End Function

Public Function getTagPosition(ByVal Nick As String) As Integer
Dim buf As Integer
buf = InStr(Nick, "<")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
buf = InStr(Nick, "[")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
getTagPosition = Len(Nick) + 2
End Function



Public Function getStrenghtColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getStrenghtColor = RGB(255 - (M * UserFuerza), (M * UserFuerza), 0)
End Function
Public Function getDexterityColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getDexterityColor = RGB(255, M * UserAgilidad, 0)
End Function

Public Function getCharIndexByName(ByVal name As String) As Integer
Dim i As Long
For i = 1 To LastChar
    If charlist(i).Nombre = name Then
        getCharIndexByName = i
        Exit Function
    End If
Next i
End Function

'---------------------------------------------------------------------------------------
' Procedure : IniBot
' Author    : Anagrama
' Date      : ???
' Purpose   : Inicializa el juego, los bots, el usuario, etc.
'---------------------------------------------------------------------------------------
'
Public Sub IniBot(Optional ByVal Balance As String)

    Dim i As Byte

    If UBound(Team1) > 1 Then
        For i = 1 To UBound(Team1) - 1
            Team1(i) = i
            charlist(i).Ai = TeamData1(i).Clase
            charlist(i).Raza = TeamData1(i).Raza
            charlist(i).Genero = TeamData1(i).Genero
            charlist(i).Lvl = TeamData1(i).Nivel
            charlist(i).Bot = 1
        Next i
    End If
    
    For i = 1 To UBound(Team2)
        If UBound(Team1) = 1 Then
            Team2(i) = Team1(UBound(Team1)) + i
            charlist(Team1(UBound(Team1)) + i).Ai = TeamData2(i).Clase
            charlist(Team1(UBound(Team1)) + i).Raza = TeamData2(i).Raza
            charlist(Team1(UBound(Team1)) + i).Genero = TeamData2(i).Genero
            charlist(Team1(UBound(Team1)) + i).Lvl = TeamData2(i).Nivel
            charlist(Team1(UBound(Team1)) + i).Bot = 1
        Else
            Team2(i) = Team1(UBound(Team1) - 1) + i
            charlist(Team1(UBound(Team1) - 1) + i).Ai = TeamData2(i).Clase
            charlist(Team1(UBound(Team1) - 1) + i).Raza = TeamData2(i).Raza
            charlist(Team1(UBound(Team1) - 1) + i).Genero = TeamData2(i).Genero
            charlist(Team1(UBound(Team1) - 1) + i).Lvl = TeamData2(i).Nivel
            charlist(Team1(UBound(Team1) - 1) + i).Bot = 1
        End If
    Next i
    
    If Balance <> vbNullString Then
        Call LoadBalance(Balance)
    End If
    
    ReDim CharTimer(1 To MaxUsers) As New clsTimer
    
    pausa = True
    
    Call CrearChars
        
    For i = 1 To MaxUsers
        Call CharTimer(i).Start(TimersIndex.Attack)
        Call CharTimer(i).Start(TimersIndex.Work)
        Call CharTimer(i).Start(TimersIndex.UseItemWithU)
        Call CharTimer(i).Start(TimersIndex.UseItemWithDblClick)
        Call CharTimer(i).Start(TimersIndex.SendRPU)
        Call CharTimer(i).Start(TimersIndex.CastSpell)
        Call CharTimer(i).Start(TimersIndex.Arrows)
        Call CharTimer(i).Start(TimersIndex.CastAttack)
        Call CharTimer(i).Start(TimersIndex.AttackCast)
        Call CharTimer(i).Start(TimersIndex.UseItem)
        Call CharTimer(i).Start(TimersIndex.GolpeU)
        Call CharTimer(i).Start(TimersIndex.WaitH)
        Call CharTimer(i).Start(TimersIndex.WaitP)
        Call CharTimer(i).Start(TimersIndex.Remo)
        Call CharTimer(i).Start(TimersIndex.Lan)
        Call CharTimer(i).Start(TimersIndex.RemoOtro)
        Call CharTimer(i).Start(TimersIndex.Resu)
        If Dificultad > 3 Then
            Call CharTimer(i).SetInterval(TimersIndex.Work, INT_WORK)
            Call CharTimer(i).SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
            Call CharTimer(i).SetInterval(TimersIndex.Arrows, INT_ARROWS)
            Call CharTimer(i).SetInterval(TimersIndex.UseItem, INT_MINU)
            Call CharTimer(i).SetInterval(TimersIndex.GolpeU, INT_GOLPEU)
            Call CharTimer(i).SetInterval(TimersIndex.Attack, INT_ATTACK + Dificultad * 50)
            Call CharTimer(i).SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK + Dificultad * 75)
            Call CharTimer(i).SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU + Dificultad * 25)
            Call CharTimer(i).SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL + Dificultad * 50)
            Call CharTimer(i).SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK + Dificultad * 50)
            Call CharTimer(i).SetInterval(TimersIndex.AttackCast, INT_ATTACK_CAST + Dificultad * 50)
            Call CharTimer(i).SetInterval(TimersIndex.WaitH, INT_PASOIH + Dificultad * 100)
            Call CharTimer(i).SetInterval(TimersIndex.WaitP, INT_PASOIH + Dificultad * 100)
            Call CharTimer(i).SetInterval(TimersIndex.Remo, INT_REMO)
            Call CharTimer(i).SetInterval(TimersIndex.Lan, INT_LAN)
        ElseIf Dificultad = 3 Then
            Call CharTimer(i).SetInterval(TimersIndex.Work, INT_WORK)
            Call CharTimer(i).SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
            Call CharTimer(i).SetInterval(TimersIndex.Arrows, INT_ARROWS)
            Call CharTimer(i).SetInterval(TimersIndex.UseItem, INT_MINU)
            Call CharTimer(i).SetInterval(TimersIndex.GolpeU, INT_GOLPEU)
            Call CharTimer(i).SetInterval(TimersIndex.Attack, INT_ATTACK)
            Call CharTimer(i).SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
            Call CharTimer(i).SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
            Call CharTimer(i).SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
            Call CharTimer(i).SetInterval(TimersIndex.AttackCast, INT_ATTACK_CAST)
            Call CharTimer(i).SetInterval(TimersIndex.WaitH, INT_PASOIH)
            Call CharTimer(i).SetInterval(TimersIndex.WaitP, INT_PASOIH)
            Call CharTimer(i).SetInterval(TimersIndex.Remo, INT_REMO)
            Call CharTimer(i).SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK + 100)
            Call CharTimer(i).SetInterval(TimersIndex.Lan, INT_LAN)
            Call CharTimer(i).SetInterval(TimersIndex.RemoOtro, INT_REMOOTRO)
            Call CharTimer(i).SetInterval(TimersIndex.Resu, INT_RESU)
        Else
            Call CharTimer(i).SetInterval(TimersIndex.Work, INT_WORK)
            Call CharTimer(i).SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
            Call CharTimer(i).SetInterval(TimersIndex.Arrows, INT_ARROWS)
            Call CharTimer(i).SetInterval(TimersIndex.UseItem, INT_MINU)
            Call CharTimer(i).SetInterval(TimersIndex.GolpeU, INT_GOLPEU)
            Call CharTimer(i).SetInterval(TimersIndex.Attack, INT_ATTACK - Dificultad * 25)
            Call CharTimer(i).SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK - Dificultad * 33)
            Call CharTimer(i).SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU - Dificultad * 12)
            Call CharTimer(i).SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL - Dificultad * 25)
            Call CharTimer(i).SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK - Dificultad * 25)
            Call CharTimer(i).SetInterval(TimersIndex.AttackCast, INT_ATTACK_CAST - Dificultad * 25)
            Call CharTimer(i).SetInterval(TimersIndex.WaitH, INT_PASOIH - Dificultad * 50)
            Call CharTimer(i).SetInterval(TimersIndex.WaitP, INT_PASOIH - Dificultad * 50)
            Call CharTimer(i).SetInterval(TimersIndex.Remo, INT_REMO)
            Call CharTimer(i).SetInterval(TimersIndex.Lan, INT_LAN)
            Call CharTimer(i).SetInterval(TimersIndex.RemoOtro, INT_REMOOTRO)
            Call CharTimer(i).SetInterval(TimersIndex.Resu, INT_RESU)
        End If
    Next i

End Sub

'---------------------------------------------------------------------------------------
' Procedure : IniClient
' Author    : Anagrama
' Date      : ???
' Purpose   : Inicializa el cliente con la información necesaria para funcionar.
'---------------------------------------------------------------------------------------
'
Public Sub IniClient()
    
    UserMap = 1
    UserPos.X = charlist(UserCharIndex).Pos.X
    UserPos.Y = charlist(UserCharIndex).Pos.Y
    UserLvl = charlist(UserCharIndex).Lvl
    Nombres = True
    MinLimiteX = 10
    MinLimiteY = 10
    MaxLimiteX = 90
    MaxLimiteY = 90

    UserParalizado = False

    frmMain.hlst.Clear
    frmMain.hlst.AddItem "(Nada)"
    frmMain.hlst.AddItem "(Nada)"
    frmMain.hlst.AddItem "(Nada)"
    frmMain.hlst.AddItem "(Nada)"
    frmMain.hlst.AddItem "(Nada)"
    frmMain.hlst.AddItem "(Nada)"
    frmMain.hlst.AddItem "Resucitar"
    frmMain.hlst.AddItem "Tormenta de Fuego"
    frmMain.hlst.AddItem "Descarga Eléctrica"
    frmMain.hlst.AddItem "Apocalípsis"
    frmMain.hlst.AddItem "Inmovilizar"
    frmMain.hlst.AddItem "Devolver Movilidad"
    
    frmMain.lblGanados = DGanados
    frmMain.lblPerdidos = DPerdidos
    If DGanados > 0 Then
        frmMain.lblPromedio = Int(DGanados * 100 / (DGanados + DPerdidos)) & "%"
    Else: frmMain.lblPromedio = "0%"
    End If
    
    Dim DifName As String
    Select Case Dificultad
        Case 1: DifName = "Imposible"
        Case 2: DifName = "Muy Dificil"
        Case 3: DifName = "Dificil"
        Case 4: DifName = "Fácil"
        Case 5: DifName = "Muy Fácil"
    End Select
    frmMain.lblDificultad = DifName

    If charlist(UserCharIndex).Ai = eBotAi.Guerrero Then
        Call Inventario.SetItem(1, 1, 0, 0, 599, 1, 0, 0, 0, 0, 0, "Poción roja")
        Call Inventario.SetItem(3, 3, 0, 1, 601, 1, 0, 0, 0, 0, 0, "Hacha de Guerra Dos Filos")
        Call Inventario.SetItem(2, 4, 0, 0, 602, 1, 0, 0, 0, 0, 0, "Arco de Cazador")
        Call Inventario.SetItem(5, 4, 0, 0, 602, 1, 0, 0, 0, 0, 0, "Arco de Cazador")
        Call Inventario.SetItem(4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
    ElseIf charlist(UserCharIndex).Ai = eBotAi.Bardo Then
        Call Inventario.SetItem(1, 1, 0, 0, 599, 1, 0, 0, 0, 0, 0, "Poción roja")
        Call Inventario.SetItem(2, 2, 0, 0, 600, 1, 0, 0, 0, 0, 0, "Poción azul")
        Call Inventario.SetItem(5, 2, 0, 0, 600, 1, 0, 0, 0, 0, 0, "Poción azul")
        Call Inventario.SetItem(3, 5, 0, 1, 603, 1, 0, 0, 0, 0, 0, "Laúd Élfico")
        Call Inventario.SetItem(4, 6, 0, 0, 604, 1, 0, 0, 0, 0, 0, "Anillo de Disolución Mágica")
    Else
        Call Inventario.SetItem(1, 1, 0, 0, 599, 1, 0, 0, 0, 0, 0, "Poción roja")
        Call Inventario.SetItem(2, 2, 0, 0, 600, 1, 0, 0, 0, 0, 0, "Poción azul")
        Call Inventario.SetItem(5, 2, 0, 0, 600, 1, 0, 0, 0, 0, 0, "Poción azul")
        Call Inventario.SetItem(4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
        Call Inventario.SetItem(3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
    End If
    
    Call SwitchMap(1)
    frmMain.Barras.Enabled = True
    
    If Hosting = 1 Then
        pausa = False
        frmMain.Bot.Enabled = False
        frmMain.Bot.Enabled = True
        frmMain.StaRecovery.Enabled = True
        EnCuenta = True
        CuentaR = 3
        frmMain.Cuenta.Enabled = True
        Call DarPrioridadTarget(1)
        Call DarPrioridadTarget(2)
        Call GetTargetBotTeam(1)
        Call GetTargetBotTeam(2)
    End If
    
    frmMain.Visible = True
End Sub

Public Sub CrearChars()
'---------------------------------------------------------------------------------------
' Procedure : CrearChars
' Author    : Anagrama
' Date      : ???
' Purpose   : Crea todos los personajes con la información otorgada.
'---------------------------------------------------------------------------------------
'
    Dim i As Byte

    For i = 1 To MaxUsers
        If charlist(i).ConnID <> -1 Or charlist(i).Bot = 1 Then
            Call CrearChar(i)
            Call ServerSendData(SendTarget.ToAllButIndex, UserCharIndex, PrepareMessageCharacterCreate(charlist(i).iBody, _
                                charlist(i).iHead, charlist(i).heading, i, charlist(i).Pos.X, charlist(i).Pos.Y, _
                                charlist(i).iArma, charlist(i).iEscudo, charlist(i).iCasco, charlist(i).Nombre, charlist(i).MinHP, EnTeam(i)))
            If charlist(i).Bot = 0 Then
                Call WritePosUpdate(i)
                Call WriteUpdateCharStats(i)
            End If
        End If
    Next i
    
End Sub

Public Sub CrearChar(ByVal CharIndex As Integer)
'---------------------------------------------------------------------------------------
' Procedure : CrearChar
' Author    : Anagrama
' Date      : ???
' Purpose   : Crea el personaje especifico.
'---------------------------------------------------------------------------------------
'
    Dim Head As Integer
    Dim body As Integer
    Dim Arma As Byte
    Dim Escudo As Byte
    Dim Casco As Byte
    Dim X As Byte
    Dim Y As Byte
    Dim N As Byte
    Dim tmpSta As Integer

    If RandomBots Then
        If charlist(CharIndex).Bot = 1 Then
            charlist(CharIndex).Raza = RandomNumber(0, 4)
            charlist(CharIndex).Genero = RandomNumber(0, 1)
        End If
    End If
    
    If RandomAiBots Then
        If charlist(CharIndex).Bot = 1 Then
            charlist(CharIndex).Ai = RandomNumber(0, 6)
        End If
    End If
    
    If EnTeam(CharIndex) = 1 Then
        charlist(CharIndex).Nombre = TeamData1(GetTeamIndex(CharIndex, 1)).Nombre
    Else
        charlist(CharIndex).Nombre = TeamData2(GetTeamIndex(CharIndex, 2)).Nombre
    End If
    charlist(CharIndex).MinHP = (VidaClase(charlist(CharIndex).Ai) + VidaRaza(charlist(CharIndex).Raza)) * charlist(CharIndex).Lvl
    charlist(CharIndex).MaxHP = (VidaClase(charlist(CharIndex).Ai) + VidaRaza(charlist(CharIndex).Raza)) * charlist(CharIndex).Lvl
    tmpSta = StaClase(charlist(CharIndex).Ai)
    charlist(CharIndex).MinSTA = tmpSta * charlist(CharIndex).Lvl + 60
    charlist(CharIndex).MaxSTA = tmpSta * charlist(CharIndex).Lvl + 60
    If charlist(CharIndex).Ai = eBotAi.Mago Then
        If ModBalance = 1 Then
            charlist(CharIndex).MinMAN = Int(ManaBaseClaseRaza(charlist(CharIndex).Ai, charlist(CharIndex).Raza) * ManaBaseModifClase(charlist(CharIndex).Ai)) * (charlist(CharIndex).Lvl - 1) + ManaBaseClaseRaza(charlist(CharIndex).Ai, charlist(CharIndex).Raza) * 3
            charlist(CharIndex).MaxMAN = Int(ManaBaseClaseRaza(charlist(CharIndex).Ai, charlist(CharIndex).Raza) * ManaBaseModifClase(charlist(CharIndex).Ai)) * (charlist(CharIndex).Lvl - 1) + ManaBaseClaseRaza(charlist(CharIndex).Ai, charlist(CharIndex).Raza) * 3
        ElseIf ModBalance = 2 Then
            charlist(CharIndex).MinMAN = (Int(ManaBaseClaseRaza(charlist(CharIndex).Ai, charlist(CharIndex).Raza) * ManaBaseModifClase(charlist(CharIndex).Ai)) - 20) * (charlist(CharIndex).Lvl - 1) + ManaStartClase(charlist(CharIndex).Ai)
            charlist(CharIndex).MaxMAN = (Int(ManaBaseClaseRaza(charlist(CharIndex).Ai, charlist(CharIndex).Raza) * ManaBaseModifClase(charlist(CharIndex).Ai)) - 20) * (charlist(CharIndex).Lvl - 1) + ManaStartClase(charlist(CharIndex).Ai)
        End If
    Else
        charlist(CharIndex).MinMAN = ManaBaseClaseRaza(charlist(CharIndex).Ai, charlist(CharIndex).Raza) * ManaBaseModifClase(charlist(CharIndex).Ai) * (charlist(CharIndex).Lvl - 1) + ManaStartClase(charlist(CharIndex).Ai)
        charlist(CharIndex).MaxMAN = ManaBaseClaseRaza(charlist(CharIndex).Ai, charlist(CharIndex).Raza) * ManaBaseModifClase(charlist(CharIndex).Ai) * (charlist(CharIndex).Lvl - 1) + ManaStartClase(charlist(CharIndex).Ai)
    End If
    charlist(CharIndex).Fuerza = FuerzaRaza(charlist(CharIndex).Raza)
    charlist(CharIndex).Agilidad = AgilidadRaza(charlist(CharIndex).Raza)
    body = BodyClaseRaza(charlist(CharIndex).Ai, charlist(CharIndex).Raza)
    Head = HeadRazaGenero(charlist(CharIndex).Raza, charlist(CharIndex).Genero)
    Arma = ArmaClase(charlist(CharIndex).Ai)
    Escudo = EscudoClase(charlist(CharIndex).Ai)
    Casco = CascoClase(charlist(CharIndex).Ai)
    charlist(CharIndex).MinRM = MinRMClase(charlist(CharIndex).Ai)
    charlist(CharIndex).MaxRM = MaxRMClase(charlist(CharIndex).Ai)
    charlist(CharIndex).ModEvasion = ModEvasionClase(charlist(CharIndex).Ai)
    charlist(CharIndex).ModAtaqueArma = ModAtaqueArmaClase(charlist(CharIndex).Ai)
    charlist(CharIndex).ModAtaqueProyectil = ModAtaqueProyectilClase(charlist(CharIndex).Ai)
    charlist(CharIndex).ModDañoArma = ModDañoArmaClase(charlist(CharIndex).Ai)
    charlist(CharIndex).ModDañoProyectil = ModDañoProyectilClase(charlist(CharIndex).Ai)
    charlist(CharIndex).ModEscudo = ModEscudoClase(charlist(CharIndex).Ai)
    charlist(CharIndex).MinDef = MinDefBClase(charlist(CharIndex).Ai)
    charlist(CharIndex).MaxDef = MaxDefBClase(charlist(CharIndex).Ai)
    charlist(CharIndex).MinDefH = MinDefHClase(charlist(CharIndex).Ai)
    charlist(CharIndex).MaxDefH = MaxDefHClase(charlist(CharIndex).Ai)
    charlist(CharIndex).ArmaMinHit = MinArmaClase(charlist(CharIndex).Ai)
    charlist(CharIndex).ArmaMaxHit = MaxArmaClase(charlist(CharIndex).Ai)
    charlist(CharIndex).DM = ModDMClase(charlist(CharIndex).Ai)
    If charlist(CharIndex).Ai = eBotAi.Paladin Or charlist(CharIndex).Ai = eBotAi.Asesino Then
        charlist(CharIndex).MinHit = 1 + MinHitClase(charlist(CharIndex).Ai) * (charlist(CharIndex).Lvl - 34) + MaxHitClase(charlist(CharIndex).Ai) * 34
        charlist(CharIndex).MaxHit = 2 + MinHitClase(charlist(CharIndex).Ai) * (charlist(CharIndex).Lvl - 34) + MaxHitClase(charlist(CharIndex).Ai) * 34
        charlist(CharIndex).TipoArma = 1
    ElseIf charlist(CharIndex).Ai = eBotAi.Guerrero Then
        charlist(CharIndex).MinHit = 1 + MinHitClase(charlist(CharIndex).Ai) * (charlist(CharIndex).Lvl - 34) + MaxHitClase(charlist(CharIndex).Ai) * 34
        charlist(CharIndex).MaxHit = 2 + MinHitClase(charlist(CharIndex).Ai) * (charlist(CharIndex).Lvl - 34) + MaxHitClase(charlist(CharIndex).Ai) * 34
        charlist(CharIndex).Refuerzo = ItemData(2).Refuerzo
        charlist(CharIndex).TipoArma = 2
    Else
        charlist(CharIndex).MinHit = 1 + MinHitClase(charlist(CharIndex).Ai) * (charlist(CharIndex).Lvl - 1)
        charlist(CharIndex).MaxHit = 2 + MinHitClase(charlist(CharIndex).Ai) * (charlist(CharIndex).Lvl - 1)
        charlist(CharIndex).TipoArma = 1
    End If
    
    If EnTeam(CharIndex) = 1 Then
        X = 35 + GetTeamIndex(CharIndex, 1)
        Y = 36
    Else
        X = 61 - GetTeamIndex(CharIndex, 2)
        Y = 57
    End If

    If charlist(CharIndex).Bot = 1 Then
        If EnTeam(CharIndex) = 2 Then
            If UBound(Team2) = 1 Then
                charlist(CharIndex).Nombre = "Anagrama"
            Else
                charlist(CharIndex).Nombre = "Anagrama " & CharIndex
            End If
            charlist(CharIndex).ComportamientoHechizos = IIf(RandomNumber(1, 10) <= 3, 1, 2)
            charlist(CharIndex).ComportamientoPotas = IIf(RandomNumber(1, 10) <= 3, 1, 2)
            
            If charlist(CharIndex).Ai = eBotAi.Paladin Or charlist(CharIndex).Ai = eBotAi.Druida Or charlist(CharIndex).Ai = eBotAi.Asesino Then
                charlist(CharIndex).ComportamientoCombo = 1
            ElseIf charlist(CharIndex).Ai = eBotAi.Bardo Then
                charlist(CharIndex).ComportamientoCombo = 2
            ElseIf charlist(CharIndex).Ai = eBotAi.Clerigo Then
                charlist(CharIndex).ComportamientoCombo = IIf(RandomNumber(1, 10) > 1, 1, 2)
            End If
            
            If charlist(CharIndex).Ai <> eBotAi.Guerrero Then
                charlist(CharIndex).TipoPocion = RandomNumber(1, 2)
                charlist(CharIndex).Lanzando = RandomNumber(0, 1)
            Else
                charlist(CharIndex).TipoPocion = 1
                charlist(CharIndex).Lanzando = 1
            End If
        Else
            If UBound(Team1) = 1 Then
                charlist(CharIndex).Nombre = "Anagrama"
            Else
                charlist(CharIndex).Nombre = "Anagrama " & CharIndex
            End If
            charlist(CharIndex).ComportamientoHechizos = IIf(RandomNumber(1, 10) <= 3, 1, 2)
            charlist(CharIndex).ComportamientoPotas = IIf(RandomNumber(1, 10) <= 3, 1, 2)
            
            If charlist(CharIndex).Ai = eBotAi.Paladin Or charlist(CharIndex).Ai = eBotAi.Druida Or charlist(CharIndex).Ai = eBotAi.Asesino Then
                charlist(CharIndex).ComportamientoCombo = 1
            ElseIf charlist(CharIndex).Ai = eBotAi.Bardo Then
                charlist(CharIndex).ComportamientoCombo = 2
            ElseIf charlist(CharIndex).Ai = eBotAi.Clerigo Then
                charlist(CharIndex).ComportamientoCombo = IIf(RandomNumber(1, 10) > 1, 1, 2)
            End If
            
            If charlist(CharIndex).Ai <> eBotAi.Guerrero Then
                charlist(CharIndex).TipoPocion = RandomNumber(1, 2)
                charlist(CharIndex).Lanzando = RandomNumber(0, 1)
            Else
                charlist(CharIndex).TipoPocion = 1
                charlist(CharIndex).Lanzando = 1
            End If
        End If
    End If
    
    Call MakeChar(CharIndex, body, Head, 3, X, Y, Arma, Escudo, Casco)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : LanzarSpell
' Author    : Anagrama
' Date      : ???
' Purpose   : Intenta lanzar un hechizo a la posición otorgada, en caso de cumplir los requisitos lo realiza.
'---------------------------------------------------------------------------------------
'
Public Sub LanzarSpell(ByVal CasterIndex As Integer, ByVal index As Byte, ByVal X As Byte, ByVal Y As Byte)
    Dim CharIndex As Integer
    Dim daño As Integer
    Dim i As Byte
    
    If charlist(CasterIndex).MinHP = 0 Then
        Call WriteConsoleMsg(CasterIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If charlist(CasterIndex).Ai <> eBotAi.Druida Or index = 1 Then
        If charlist(CasterIndex).MinMAN < Hechizo(index).Mana Then
            Call WriteConsoleMsg(CasterIndex, "No tienes suficiente mana.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        If charlist(CasterIndex).MinMAN < Hechizo(index).Mana * 0.9 Then
            Call WriteConsoleMsg(CasterIndex, "No tienes suficiente mana.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    If charlist(CasterIndex).MinSTA < Hechizo(index).Sta Then
        Call WriteConsoleMsg(CasterIndex, "Estás muy cansado.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
        
    CharIndex = IIf(MapData(X, Y + 1).CharIndex > 0, MapData(X, Y + 1).CharIndex, MapData(X, Y).CharIndex)
    
    If CharIndex > 0 Then
        If CharIndex = CasterIndex Then
            If index = 4 Then
                If charlist(CasterIndex).Inmo Then
                    charlist(CasterIndex).MinMAN = charlist(CasterIndex).MinMAN - IIf(charlist(CasterIndex).Ai <> eBotAi.Druida, Hechizo(index).Mana, Hechizo(index).Mana * 0.9)
                    charlist(CasterIndex).MinSTA = charlist(CasterIndex).MinSTA - Hechizo(index).Sta
                    charlist(CasterIndex).Inmo = 0
                    Call WriteParalizeOK(CharIndex, charlist(CasterIndex).Inmo)
                    Call WriteConsoleMsg(CasterIndex, "Has lanzado " & Hechizo(index).name & " sobre " & charlist(CharIndex).Nombre, FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteUpdateCharStats(CasterIndex)
                    Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageChatOverHead(Hechizo(index).Palabras, CasterIndex, vbCyan))
                    Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessagePlayWave(Hechizo(index).WAV, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
                    Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageCreateFX(CharIndex, Hechizo(index).FX, 0))
                End If
            End If
        Else
            If EnTeam(CharIndex) = EnTeam(CasterIndex) Then
                If index = 4 Then
                    charlist(CasterIndex).MinMAN = charlist(CasterIndex).MinMAN - IIf(charlist(CasterIndex).Ai <> eBotAi.Druida, Hechizo(index).Mana, Hechizo(index).Mana * 0.9)
                    charlist(CasterIndex).MinSTA = charlist(CasterIndex).MinSTA - Hechizo(index).Sta
                    charlist(CharIndex).Inmo = 0
                    If charlist(CharIndex).Bot = 0 Then Call WriteParalizeOK(CharIndex, charlist(CharIndex).Inmo)
                    Call WriteConsoleMsg(CasterIndex, "Has lanzado " & Hechizo(index).name & " sobre " & charlist(CharIndex).Nombre, FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteUpdateCharStats(CasterIndex)
                    Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageChatOverHead(Hechizo(index).Palabras, CasterIndex, vbCyan))
                    Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessagePlayWave(Hechizo(index).WAV, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
                    Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageCreateFX(CharIndex, Hechizo(index).FX, 0))
                    Exit Sub
                ElseIf index = 6 Then
                    If SinResu = 0 Then
                        If charlist(CharIndex).MinHP > 0 Then Exit Sub
                        charlist(CasterIndex).MinMAN = charlist(CasterIndex).MinMAN - IIf(charlist(CasterIndex).Ai <> eBotAi.Druida, Hechizo(index).Mana, Hechizo(index).Mana * 0.9)
                        charlist(CasterIndex).MinSTA = charlist(CasterIndex).MinSTA - Hechizo(index).Sta
                        charlist(CharIndex).MinHP = 1
                        charlist(CharIndex).MinMAN = 0
                        charlist(CharIndex).body = BodyData(BodyClaseRaza(charlist(CharIndex).Ai, charlist(CharIndex).Raza))
                        charlist(CharIndex).Head = HeadData(HeadRazaGenero(charlist(CharIndex).Raza, charlist(CharIndex).Genero))
                        charlist(CharIndex).Arma = WeaponAnimData(ArmaClase(charlist(CharIndex).Ai))
                        charlist(CharIndex).Escudo = ShieldAnimData(EscudoClase(charlist(CharIndex).Ai))
                        charlist(CharIndex).Casco = CascoAnimData(CascoClase(charlist(CharIndex).Ai))
                        charlist(CharIndex).ComportamientoPotas = 4
                        charlist(CharIndex).Lanzando = 0
                        daño = charlist(CasterIndex).MaxHP - charlist(CasterIndex).MinHP * (1 - charlist(CharIndex).Lvl * 0.015)
                        If ResuNoVida = 0 Then
                            charlist(CasterIndex).MinHP = charlist(CasterIndex).MinHP - daño
                            If charlist(CasterIndex).MinHP <= 0 Then
                                charlist(CasterIndex).MinHP = 0
                                Call CharDie(CasterIndex)
                            End If
                        End If
                        If charlist(CharIndex).ComportamientoCombo = 1 And charlist(CharIndex).Ai = eBotAi.Bardo Then charlist(CharIndex).ComportamientoCombo = IIf(RandomNumber(1, 10) > 1, 2, 1)
                        Call WriteConsoleMsg(CasterIndex, "Has lanzado " & Hechizo(index).name & " sobre " & charlist(CharIndex).Nombre, FontTypeNames.FONTTYPE_FIGHT)
                        Call WriteUpdateCharStats(CasterIndex)
                        Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageChatOverHead(Hechizo(index).Palabras, CasterIndex, vbCyan))
                        Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessagePlayWave(Hechizo(index).WAV, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
                        Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageCreateFX(CharIndex, Hechizo(index).FX, 0))
                        Call DarPrioridadTarget(1)
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(CasterIndex, "No está permitido resucitar.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                Else
                    Call WriteConsoleMsg(CasterIndex, "No puedes atacar a tus compañeros.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
            End If
            If charlist(CharIndex).MinHP = 0 Then Exit Sub
            If charlist(CharIndex).Inmo <> 1 Then CPega = CPega + 1
            If index = 3 Then
                If charlist(CharIndex).Inmo = 0 Then
                    charlist(CharIndex).Lanzando = 0
                    If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                    End If
                    If CharTimer(CharIndex).Check(TimersIndex.WaitH, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.WaitH)
                    End If
                    charlist(CasterIndex).MinMAN = charlist(CasterIndex).MinMAN - IIf(charlist(CasterIndex).Ai <> eBotAi.Druida, Hechizo(index).Mana, Hechizo(index).Mana * 0.9)
                    charlist(CasterIndex).MinSTA = charlist(CasterIndex).MinSTA - Hechizo(index).Sta
                    charlist(CharIndex).Inmo = 1
                    If charlist(CharIndex).Bot = 0 Then Call WriteParalizeOK(CharIndex, charlist(CharIndex).Inmo)
                    Call CharTimer(CharIndex).SetInterval(TimersIndex.Remo, INT_REMO + RandomNumber(100 * Dificultad, 200 * Dificultad))
                    Call CharTimer(CharIndex).Restart(TimersIndex.Remo)
                    For i = 1 To UBound(Team2)
                        Call CharTimer(Team2(i)).SetInterval(TimersIndex.RemoOtro, INT_REMOOTRO + RandomNumber(100 * Dificultad, 200 * Dificultad))
                        Call CharTimer(Team2(i)).Restart(TimersIndex.RemoOtro)
                    Next i
                    If charlist(CharIndex).Ai <> eBotAi.Mago Or charlist(CharIndex).Ai <> eBotAi.Druida Then charlist(CharIndex).ComportamientoCombo = 2
                    Call WriteConsoleMsg(CasterIndex, "Has lanzado " & Hechizo(index).name & " sobre " & charlist(CharIndex).Nombre, FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteUpdateCharStats(CasterIndex)
                    Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageChatOverHead(Hechizo(index).Palabras, CasterIndex, vbCyan))
                    Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessagePlayWave(Hechizo(index).WAV, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
                    Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageCreateFX(CharIndex, Hechizo(index).FX, 0))
                    If UBound(Team1) > 1 Then
                        Call GetTargetBotTeam(CasterIndex)
                    End If
                End If
            ElseIf index = 1 Then
                charlist(CasterIndex).MinMAN = charlist(CasterIndex).MinMAN - Hechizo(index).Mana
                charlist(CasterIndex).MinSTA = charlist(CasterIndex).MinSTA - Hechizo(index).Sta
                daño = RandomNumber(Hechizo(index).MinHP, Hechizo(index).MaxHP)
                daño = daño + (daño * 3 * UserLvl) / 100
                daño = daño * charlist(CasterIndex).DM
                daño = daño - RandomNumber(charlist(CharIndex).MinRM, charlist(CharIndex).MaxRM)
                charlist(CharIndex).MinHP = charlist(CharIndex).MinHP - daño
                Call WriteConsoleMsg(CasterIndex, "Has lanzado " & Hechizo(index).name & " sobre " & charlist(CharIndex).Nombre, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(CasterIndex, "Le has quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                If charlist(CharIndex).MinHP <= 0 Then
                    charlist(CharIndex).MinHP = 0
                    Call WriteConsoleMsg(CasterIndex, "Has matado a " & charlist(CharIndex).Nombre, FontTypeNames.FONTTYPE_FIGHT)
                    Call CharDie(CharIndex)
                    Call ResetDuelo(CasterIndex)
                    If UBound(Team1) > 1 Then
                        Call GetTargetBotTeam(CasterIndex)
                    End If
                End If
                Call WriteUpdateCharStats(CasterIndex)
                If charlist(CharIndex).Bot = 0 Then Call WriteUpdateCharStats(CharIndex)
                Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageChatOverHead(Hechizo(index).Palabras, CasterIndex, vbCyan))
                Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessagePlayWave(Hechizo(index).WAV, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
                Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageCreateFX(CharIndex, Hechizo(index).FX, 0))
            ElseIf index = 2 Then
                charlist(CasterIndex).MinMAN = charlist(CasterIndex).MinMAN - IIf(charlist(CasterIndex).Ai <> eBotAi.Druida, Hechizo(index).Mana, Hechizo(index).Mana * 0.9)
                charlist(CasterIndex).MinSTA = charlist(CasterIndex).MinSTA - Hechizo(index).Sta
                daño = RandomNumber(Hechizo(index).MinHP, Hechizo(index).MaxHP)
                daño = daño + (daño * 3 * UserLvl) / 100
                daño = daño * charlist(CasterIndex).DM
                daño = daño - RandomNumber(charlist(CharIndex).MinRM, charlist(CharIndex).MaxRM)
                charlist(CharIndex).MinHP = charlist(CharIndex).MinHP - daño
                Call WriteConsoleMsg(CasterIndex, "Has lanzado " & Hechizo(index).name & " sobre " & charlist(CharIndex).Nombre, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(CasterIndex, "Le has quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                If charlist(CharIndex).MinHP <= 0 Then
                    charlist(CharIndex).MinHP = 0
                    Call WriteConsoleMsg(CasterIndex, "Has matado a " & charlist(CharIndex).Nombre, FontTypeNames.FONTTYPE_FIGHT)
                    Call CharDie(CharIndex)
                    Call ResetDuelo(CasterIndex)
                    If UBound(Team1) > 1 Then
                        Call GetTargetBotTeam(CasterIndex)
                    End If
                End If
                Call WriteUpdateCharStats(CasterIndex)
                If charlist(CharIndex).Bot = 0 Then Call WriteUpdateCharStats(CharIndex)
                Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageChatOverHead(Hechizo(index).Palabras, CasterIndex, vbCyan))
                Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessagePlayWave(Hechizo(index).WAV, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
                Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageCreateFX(CharIndex, Hechizo(index).FX, 0))
            ElseIf index = 5 Then
                charlist(CasterIndex).MinMAN = charlist(CasterIndex).MinMAN - Hechizo(index).Mana
                charlist(CasterIndex).MinSTA = charlist(CasterIndex).MinSTA - Hechizo(index).Sta
                daño = RandomNumber(Hechizo(index).MinHP, Hechizo(index).MaxHP)
                daño = daño + (daño * 3 * UserLvl) / 100
                daño = daño - RandomNumber(charlist(CharIndex).MinRM, charlist(CharIndex).MaxRM)
                charlist(CharIndex).MinHP = charlist(CharIndex).MinHP - daño
                Call WriteConsoleMsg(CasterIndex, "Has lanzado " & Hechizo(index).name & " sobre " & charlist(CharIndex).Nombre, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(CasterIndex, "Le has quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                If charlist(CharIndex).MinHP <= 0 Then
                    charlist(CharIndex).MinHP = 0
                    Call WriteConsoleMsg(CasterIndex, "Has matado a " & charlist(CharIndex).Nombre, FontTypeNames.FONTTYPE_FIGHT)
                    Call CharDie(CharIndex)
                    Call ResetDuelo(CasterIndex)
                    If UBound(Team1) > 1 Then
                        Call GetTargetBotTeam(CasterIndex)
                    End If
                End If
                Call WriteUpdateCharStats(CasterIndex)
                If charlist(CharIndex).Bot = 0 Then Call WriteUpdateCharStats(CharIndex)
                Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageChatOverHead(Hechizo(index).Palabras, CasterIndex, vbCyan))
                Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessagePlayWave(Hechizo(index).WAV, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
                Call ServerSendData(SendTarget.ToAll, CasterIndex, PrepareMessageCreateFX(CharIndex, Hechizo(index).FX, 0))
                Else: Exit Sub
            End If
            If (charlist(CharIndex).ComportamientoPotas = 2 Or charlist(CharIndex).MinHP < charlist(CharIndex).MaxHP / 2) And index <> 3 Then
                charlist(CharIndex).ComportamientoPotas = 3
                If charlist(CharIndex).Lanzando = 1 Then
                    charlist(CharIndex).Lanzando = 0
                    If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                    End If
                End If
            End If
            If charlist(CharIndex).Ai = eBotAi.Paladin Or charlist(CharIndex).Ai = eBotAi.Asesino Or charlist(CharIndex).Ai = eBotAi.Clerigo Or charlist(CharIndex).Ai = eBotAi.Bardo Or charlist(CharIndex).Ai = eBotAi.Guerrero Then
                If UserParalizado And charlist(CharIndex).Inmo = 0 Then
                    charlist(CharIndex).ComportamientoPotas = 4
                    If charlist(CharIndex).Lanzando = 1 Then
                        charlist(CharIndex).Lanzando = 0
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                    End If
                ElseIf UserParalizado And charlist(CharIndex).Inmo Then
                    If charlist(CharIndex).ComportamientoPotas <> 4 And charlist(CharIndex).MinHP < charlist(CharIndex).MaxHP Then
                        charlist(CharIndex).ComportamientoPotas = 4
                        If charlist(CharIndex).Lanzando = 1 Then
                            charlist(CharIndex).Lanzando = 0
                            If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                                Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ResetDuelo
' Author    : Anagrama
' Date      : ???
' Purpose   : Reinicia el combate al morir uno de los lados.
'---------------------------------------------------------------------------------------
'
Public Sub ResetDuelo(ByVal CharIndex As Integer, Optional ByVal NumTeam1 As Byte, Optional ByVal NumTeam2 As Byte)
    Dim i As Byte
    
    If CharIndex <> -1 Then '-1 significa que se reinicia el duelo desde la configuracion
        If EnTeam(CharIndex) = 1 Then
            For i = 1 To UBound(Team2)
                If charlist(Team2(i)).MinHP > 0 Then
                    Call DarPrioridadTarget(2)
                    Exit Sub
                End If
            Next i
            DGanados = DGanados + 1
        Else
            For i = 1 To UBound(Team1)
                If charlist(Team1(i)).MinHP > 0 Then
                    Call DarPrioridadTarget(2)
                    Exit Sub
                End If
            Next i
            DPerdidos = DPerdidos + 1
        End If
    Else
        If UBound(Team1) > 0 Then
            For i = 1 To UBound(Team1)
                charlist(Team1(i)).Ai = 0
                charlist(Team1(i)).Raza = 0
                charlist(Team1(i)).Genero = 0
                charlist(Team1(i)).Lvl = 0
                NumChars = NumChars - 1
                charlist(Team1(i)).Bot = 0
                Call ServerSendData(SendTarget.ToAllButIndex, UserCharIndex, PrepareMessageCharacterRemove(Team1(i)))
                Call EraseChar(Team1(i))
                Team1(i) = 0
            Next i
        End If
        If NumTeam1 > 0 Then
            ReDim Team1(1 To NumTeam1) As Integer
        Else
            ReDim Team1(0 To NumTeam1) As Integer
        End If
        If UBound(Team2) > 0 Then
            For i = 1 To UBound(Team2)
                charlist(Team2(i)).Ai = 0
                charlist(Team2(i)).Raza = 0
                charlist(Team2(i)).Genero = 0
                charlist(Team2(i)).Lvl = 0
                NumChars = NumChars - 1
                charlist(Team2(i)).Bot = 0
                Call ServerSendData(SendTarget.ToAllButIndex, UserCharIndex, PrepareMessageCharacterRemove(Team2(i)))
                Call EraseChar(Team2(i))
                Team2(i) = 0
            Next i
        End If
        If NumTeam2 > 0 Then
            ReDim Team2(1 To NumTeam2) As Integer
        Else
            ReDim Team2(0 To NumTeam2) As Integer
        End If
        If UBound(Team1) > 0 Then
            For i = 1 To UBound(Team1)
                If TeamData1(i).index > 0 Then
                    Team1(i) = TeamData1(i).index
                    charlist(Team1(i)).Ai = TeamData1(i).Clase
                    charlist(Team1(i)).Raza = TeamData1(i).Raza
                    charlist(Team1(i)).Genero = TeamData1(i).Genero
                    charlist(Team1(i)).Lvl = TeamData1(i).Nivel
                    charlist(Team1(i)).Bot = TeamData1(i).Bot
                Else
                    Team1(i) = NextOpenTeamIndex
                    TeamData1(i).index = Team1(i)
                    charlist(Team1(i)).Ai = TeamData1(i).Clase
                    charlist(Team1(i)).Raza = TeamData1(i).Raza
                    charlist(Team1(i)).Genero = TeamData1(i).Genero
                    charlist(Team1(i)).Lvl = TeamData1(i).Nivel
                    charlist(Team1(i)).Bot = TeamData1(i).Bot
                End If
            Next i
        End If
        If UBound(Team2) > 0 Then
            For i = 1 To UBound(Team2)
                If TeamData2(i).index > 0 Then
                    Team2(i) = TeamData2(i).index
                    charlist(Team2(i)).Ai = TeamData2(i).Clase
                    charlist(Team2(i)).Raza = TeamData2(i).Raza
                    charlist(Team2(i)).Genero = TeamData2(i).Genero
                    charlist(Team2(i)).Lvl = TeamData2(i).Nivel
                    charlist(Team2(i)).Bot = TeamData2(i).Bot
                Else
                    Team2(i) = NextOpenTeamIndex
                    TeamData2(i).index = Team2(i)
                    charlist(Team2(i)).Ai = TeamData2(i).Clase
                    charlist(Team2(i)).Raza = TeamData2(i).Raza
                    charlist(Team2(i)).Genero = TeamData2(i).Genero
                    charlist(Team2(i)).Lvl = TeamData2(i).Nivel
                    charlist(Team2(i)).Bot = TeamData2(i).Bot
                End If
            Next i
        End If
    End If
    
    For i = 1 To MaxUsers
        If charlist(i).ConnID <> -1 Or charlist(i).Bot = 1 Then
            charlist(i).MinMAN = charlist(i).MaxMAN
            charlist(i).MinHP = charlist(i).MaxHP
            If charlist(i).Bot = 1 Then
                charlist(i).ComportamientoHechizos = IIf(RandomNumber(1, 10) <= 3, 1, 2)
                charlist(i).ComportamientoPotas = IIf(RandomNumber(1, 10) <= 3, 1, 2)
            Else
                If charlist(i).Inmo = 1 Then
                    charlist(i).Inmo = 0
                    Call WriteParalizeOK(i, charlist(i).Inmo)
                End If
                Call WriteUpdateCharStats(i)
            End If
            
            If CharIndex <> -1 Then
                Call ServerSendData(SendTarget.ToAllButIndex, UserCharIndex, PrepareMessageCharacterRemove(i))
                Call EraseChar(i)
            End If
        End If
    Next i
    Call CrearChars
    Call DarPrioridadTarget(1)
    Call DarPrioridadTarget(2)
    Call GetTargetBotTeam(1)
    Call GetTargetBotTeam(2)
    
    Call ServerSendData(SendTarget.ToAll, UserCharIndex, PrepareMessageCuentaToggle(1))
    EnCuenta = True
    CuentaR = 3
    frmMain.Cuenta.Enabled = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ActualizarStats
' Author    : Anagrama
' Date      : ???
' Purpose   : Actualiza la vida y la mana en el formulario.
'---------------------------------------------------------------------------------------
'
Public Sub ActualizarStats()
    frmMain.lblMana = charlist(UserCharIndex).MinMAN & "/" & charlist(UserCharIndex).MaxMAN
    frmMain.lblVida = charlist(UserCharIndex).MinHP & "/" & charlist(UserCharIndex).MaxHP
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnTeam
' Author    : Anagrama
' Date      : ???
' Purpose   : Devuelve en que team esta el char solicitado.
'---------------------------------------------------------------------------------------
'
Public Function EnTeam(ByVal CharIndex As Integer) As Byte
    Dim i As Byte
    If UBound(Team1) = 0 Then
        EnTeam = 2
        Exit Function
    End If
    For i = 1 To UBound(Team1)
        If Team1(i) = CharIndex Then
            EnTeam = 1
            Exit Function
        End If
    Next i
    EnTeam = 2
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetTeamIndex
' Author    : Anagrama
' Date      : ???
' Purpose   : Devuelve el indice que ocupa en el team.
'---------------------------------------------------------------------------------------
'
Public Function GetTeamIndex(ByVal CharIndex As Integer, ByVal Team As Byte) As Byte
    Dim i As Byte
    
    If Team = 1 Then
        For i = 1 To UBound(Team1)
            If Team1(i) = CharIndex Then
                GetTeamIndex = i
                Exit Function
            End If
        Next i
    ElseIf Team = 2 Then
        For i = 1 To UBound(Team2)
            If Team2(i) = CharIndex Then
                GetTeamIndex = i
                Exit Function
            End If
        Next i
    End If
End Function

Public Sub LoadBalance(ByVal file As String)
'---------------------------------------------------------------------------------------
' Procedure : LoadBalance
' Author    : Anagrama
' Date      : ???
' Purpose   : Carga el balance seleccionado.
'---------------------------------------------------------------------------------------
    Dim path As String
    Dim i As Byte
    Dim Leer As New clsIniReader
    
    path = PathBalance & file & ".dat"
    
    If Not FileExist(path, vbArchive) Then
        MsgBox "Error cargando archivo de balance, el archivo no existe."
        Exit Sub
    End If
    
    Call Leer.Initialize(path)
    
    ModBalance = Val(Leer.GetValue("INIT", "ModBalance"))
    
    StaClase(eBotAi.Mago) = Val(Leer.GetValue("StaClase", "StaMago"))
    StaClase(eBotAi.Paladin) = Val(Leer.GetValue("StaClase", "StaPaladin"))
    StaClase(eBotAi.Clerigo) = Val(Leer.GetValue("StaClase", "StaClerigo"))
    StaClase(eBotAi.Asesino) = Val(Leer.GetValue("StaClase", "StaAsesino"))
    StaClase(eBotAi.Bardo) = Val(Leer.GetValue("StaClase", "StaBardo"))
    StaClase(eBotAi.Druida) = Val(Leer.GetValue("StaClase", "StaDruida"))
    StaClase(eBotAi.Guerrero) = Val(Leer.GetValue("StaClase", "StaGuerrero"))

    PrioridadClase(eBotAi.Mago) = Val(Leer.GetValue("Prioridad", "PrioridadMago"))
    PrioridadClase(eBotAi.Paladin) = Val(Leer.GetValue("Prioridad", "PrioridadPaladin"))
    PrioridadClase(eBotAi.Clerigo) = Val(Leer.GetValue("Prioridad", "PrioridadClerigo"))
    PrioridadClase(eBotAi.Asesino) = Val(Leer.GetValue("Prioridad", "PrioridadAsesino"))
    PrioridadClase(eBotAi.Bardo) = Val(Leer.GetValue("Prioridad", "PrioridadBardo"))
    PrioridadClase(eBotAi.Druida) = Val(Leer.GetValue("Prioridad", "PrioridadDruida"))
    PrioridadClase(eBotAi.Guerrero) = Val(Leer.GetValue("Prioridad", "PrioridadGuerrero"))
    PrioridadRaza(eRazaAi.Humano) = Val(Leer.GetValue("Prioridad", "PrioridadH"))
    PrioridadRaza(eRazaAi.Enano) = Val(Leer.GetValue("Prioridad", "PrioridadEN"))
    PrioridadRaza(eRazaAi.Gnomo) = Val(Leer.GetValue("Prioridad", "PrioridadG"))
    PrioridadRaza(eRazaAi.Elfo) = Val(Leer.GetValue("Prioridad", "PrioridadE"))
    PrioridadRaza(eRazaAi.Drow) = Val(Leer.GetValue("Prioridad", "PrioridadEO"))
    
    VidaClase(eBotAi.Mago) = Val(Leer.GetValue("VidaClase", "Mago"))
    VidaClase(eBotAi.Paladin) = Val(Leer.GetValue("VidaClase", "Paladin"))
    VidaClase(eBotAi.Clerigo) = Val(Leer.GetValue("VidaClase", "Clerigo"))
    VidaClase(eBotAi.Asesino) = Val(Leer.GetValue("VidaClase", "Asesino"))
    VidaClase(eBotAi.Bardo) = Val(Leer.GetValue("VidaClase", "Bardo"))
    VidaClase(eBotAi.Druida) = Val(Leer.GetValue("VidaClase", "Druida"))
    VidaClase(eBotAi.Guerrero) = Val(Leer.GetValue("VidaClase", "Guerrero"))
    
    VidaRaza(eRazaAi.Humano) = Val(Leer.GetValue("VidaRaza", "Humano"))
    VidaRaza(eRazaAi.Enano) = Val(Leer.GetValue("VidaRaza", "Enano"))
    VidaRaza(eRazaAi.Gnomo) = Val(Leer.GetValue("VidaRaza", "Gnomo"))
    VidaRaza(eRazaAi.Elfo) = Val(Leer.GetValue("VidaRaza", "Elfo"))
    VidaRaza(eRazaAi.Drow) = Val(Leer.GetValue("VidaRaza", "ElfoDrow"))
    
    ManaBaseClaseRaza(eBotAi.Mago, eRazaAi.Humano) = Val(Leer.GetValue("ManaBaseClaseRaza", "MagoH"))
    ManaBaseClaseRaza(eBotAi.Mago, eRazaAi.Gnomo) = Val(Leer.GetValue("ManaBaseClaseRaza", "MagoG"))
    ManaBaseClaseRaza(eBotAi.Mago, eRazaAi.Elfo) = Val(Leer.GetValue("ManaBaseClaseRaza", "MagoE"))
    ManaBaseClaseRaza(eBotAi.Mago, eRazaAi.Enano) = Val(Leer.GetValue("ManaBaseClaseRaza", "MagoEN"))
    ManaBaseClaseRaza(eBotAi.Mago, eRazaAi.Drow) = Val(Leer.GetValue("ManaBaseClaseRaza", "MagoEO"))
    ManaBaseClaseRaza(eBotAi.Paladin, eRazaAi.Humano) = Val(Leer.GetValue("ManaBaseClaseRaza", "PaladinH"))
    ManaBaseClaseRaza(eBotAi.Paladin, eRazaAi.Gnomo) = Val(Leer.GetValue("ManaBaseClaseRaza", "PaladinG"))
    ManaBaseClaseRaza(eBotAi.Paladin, eRazaAi.Elfo) = Val(Leer.GetValue("ManaBaseClaseRaza", "PaladinE"))
    ManaBaseClaseRaza(eBotAi.Paladin, eRazaAi.Enano) = Val(Leer.GetValue("ManaBaseClaseRaza", "PaladinEN"))
    ManaBaseClaseRaza(eBotAi.Paladin, eRazaAi.Drow) = Val(Leer.GetValue("ManaBaseClaseRaza", "PaladinEO"))
    ManaBaseClaseRaza(eBotAi.Asesino, eRazaAi.Humano) = Val(Leer.GetValue("ManaBaseClaseRaza", "AsesinoH"))
    ManaBaseClaseRaza(eBotAi.Asesino, eRazaAi.Gnomo) = Val(Leer.GetValue("ManaBaseClaseRaza", "AsesinoG"))
    ManaBaseClaseRaza(eBotAi.Asesino, eRazaAi.Elfo) = Val(Leer.GetValue("ManaBaseClaseRaza", "AsesinoE"))
    ManaBaseClaseRaza(eBotAi.Asesino, eRazaAi.Enano) = Val(Leer.GetValue("ManaBaseClaseRaza", "AsesinoEN"))
    ManaBaseClaseRaza(eBotAi.Asesino, eRazaAi.Drow) = Val(Leer.GetValue("ManaBaseClaseRaza", "AsesinoEO"))
    ManaBaseClaseRaza(eBotAi.Bardo, eRazaAi.Humano) = Val(Leer.GetValue("ManaBaseClaseRaza", "BardoH"))
    ManaBaseClaseRaza(eBotAi.Bardo, eRazaAi.Gnomo) = Val(Leer.GetValue("ManaBaseClaseRaza", "BardoG"))
    ManaBaseClaseRaza(eBotAi.Bardo, eRazaAi.Elfo) = Val(Leer.GetValue("ManaBaseClaseRaza", "BardoE"))
    ManaBaseClaseRaza(eBotAi.Bardo, eRazaAi.Enano) = Val(Leer.GetValue("ManaBaseClaseRaza", "BardoEN"))
    ManaBaseClaseRaza(eBotAi.Bardo, eRazaAi.Drow) = Val(Leer.GetValue("ManaBaseClaseRaza", "BardoEO"))
    ManaBaseClaseRaza(eBotAi.Clerigo, eRazaAi.Humano) = Val(Leer.GetValue("ManaBaseClaseRaza", "ClerigoH"))
    ManaBaseClaseRaza(eBotAi.Clerigo, eRazaAi.Gnomo) = Val(Leer.GetValue("ManaBaseClaseRaza", "ClerigoG"))
    ManaBaseClaseRaza(eBotAi.Clerigo, eRazaAi.Elfo) = Val(Leer.GetValue("ManaBaseClaseRaza", "ClerigoE"))
    ManaBaseClaseRaza(eBotAi.Clerigo, eRazaAi.Enano) = Val(Leer.GetValue("ManaBaseClaseRaza", "ClerigoEN"))
    ManaBaseClaseRaza(eBotAi.Clerigo, eRazaAi.Drow) = Val(Leer.GetValue("ManaBaseClaseRaza", "ClerigoEO"))
    ManaBaseClaseRaza(eBotAi.Druida, eRazaAi.Humano) = Val(Leer.GetValue("ManaBaseClaseRaza", "DruidaH"))
    ManaBaseClaseRaza(eBotAi.Druida, eRazaAi.Gnomo) = Val(Leer.GetValue("ManaBaseClaseRaza", "DruidaG"))
    ManaBaseClaseRaza(eBotAi.Druida, eRazaAi.Elfo) = Val(Leer.GetValue("ManaBaseClaseRaza", "DruidaE"))
    ManaBaseClaseRaza(eBotAi.Druida, eRazaAi.Enano) = Val(Leer.GetValue("ManaBaseClaseRaza", "DruidaEN"))
    ManaBaseClaseRaza(eBotAi.Druida, eRazaAi.Drow) = Val(Leer.GetValue("ManaBaseClaseRaza", "DruidaEO"))
    ManaBaseClaseRaza(eBotAi.Guerrero, eRazaAi.Humano) = Val(Leer.GetValue("ManaBaseClaseRaza", "GuerreroH"))
    ManaBaseClaseRaza(eBotAi.Guerrero, eRazaAi.Gnomo) = Val(Leer.GetValue("ManaBaseClaseRaza", "GuerreroG"))
    ManaBaseClaseRaza(eBotAi.Guerrero, eRazaAi.Elfo) = Val(Leer.GetValue("ManaBaseClaseRaza", "GuerreroE"))
    ManaBaseClaseRaza(eBotAi.Guerrero, eRazaAi.Enano) = Val(Leer.GetValue("ManaBaseClaseRaza", "GuerreroEN"))
    ManaBaseClaseRaza(eBotAi.Guerrero, eRazaAi.Drow) = Val(Leer.GetValue("ManaBaseClaseRaza", "GuerreroEO"))
    
    ManaBaseModifClase(eBotAi.Mago) = Val(Leer.GetValue("ManaBaseModifClase", "Mago"))
    ManaBaseModifClase(eBotAi.Paladin) = Val(Leer.GetValue("ManaBaseModifClase", "Paladin"))
    ManaBaseModifClase(eBotAi.Clerigo) = Val(Leer.GetValue("ManaBaseModifClase", "Clerigo"))
    ManaBaseModifClase(eBotAi.Asesino) = Val(Leer.GetValue("ManaBaseModifClase", "Asesino"))
    ManaBaseModifClase(eBotAi.Bardo) = Val(Leer.GetValue("ManaBaseModifClase", "Bardo"))
    ManaBaseModifClase(eBotAi.Druida) = Val(Leer.GetValue("ManaBaseModifClase", "Druida"))
    ManaBaseModifClase(eBotAi.Guerrero) = Val(Leer.GetValue("ManaBaseModifClase", "Guerrero"))
    
    ManaStartClase(eBotAi.Mago) = Val(Leer.GetValue("ManaStartClase", "Mago"))
    ManaStartClase(eBotAi.Paladin) = Val(Leer.GetValue("ManaStartClase", "Paladin"))
    ManaStartClase(eBotAi.Clerigo) = Val(Leer.GetValue("ManaStartClase", "Clerigo"))
    ManaStartClase(eBotAi.Asesino) = Val(Leer.GetValue("ManaStartClase", "Asesino"))
    ManaStartClase(eBotAi.Bardo) = Val(Leer.GetValue("ManaStartClase", "Bardo"))
    ManaStartClase(eBotAi.Druida) = Val(Leer.GetValue("ManaStartClase", "Druida"))
    ManaStartClase(eBotAi.Guerrero) = Val(Leer.GetValue("ManaStartClase", "Guerrero"))
    
    MinHitClase(eBotAi.Mago) = Val(Leer.GetValue("MinHitClase", "Mago"))
    MinHitClase(eBotAi.Paladin) = Val(Leer.GetValue("MinHitClase", "Paladin"))
    MinHitClase(eBotAi.Clerigo) = Val(Leer.GetValue("MinHitClase", "Clerigo"))
    MinHitClase(eBotAi.Asesino) = Val(Leer.GetValue("MinHitClase", "Asesino"))
    MinHitClase(eBotAi.Bardo) = Val(Leer.GetValue("MinHitClase", "Bardo"))
    MinHitClase(eBotAi.Druida) = Val(Leer.GetValue("MinHitClase", "Druida"))
    MinHitClase(eBotAi.Guerrero) = Val(Leer.GetValue("MinHitClase", "Guerrero"))
    
    MaxHitClase(eBotAi.Mago) = Val(Leer.GetValue("MaxHitClase", "Mago"))
    MaxHitClase(eBotAi.Paladin) = Val(Leer.GetValue("MaxHitClase", "Paladin"))
    MaxHitClase(eBotAi.Clerigo) = Val(Leer.GetValue("MaxHitClase", "Clerigo"))
    MaxHitClase(eBotAi.Asesino) = Val(Leer.GetValue("MaxHitClase", "Asesino"))
    MaxHitClase(eBotAi.Bardo) = Val(Leer.GetValue("MaxHitClase", "Bardo"))
    MaxHitClase(eBotAi.Druida) = Val(Leer.GetValue("MaxHitClase", "Druida"))
    MaxHitClase(eBotAi.Guerrero) = Val(Leer.GetValue("MaxHitClase", "Guerrero"))
    
    FuerzaRaza(eRazaAi.Humano) = Val(Leer.GetValue("FuerzaRaza", "Humano"))
    FuerzaRaza(eRazaAi.Enano) = Val(Leer.GetValue("FuerzaRaza", "Enano"))
    FuerzaRaza(eRazaAi.Gnomo) = Val(Leer.GetValue("FuerzaRaza", "Gnomo"))
    FuerzaRaza(eRazaAi.Elfo) = Val(Leer.GetValue("FuerzaRaza", "Elfo"))
    FuerzaRaza(eRazaAi.Drow) = Val(Leer.GetValue("FuerzaRaza", "ElfoDrow"))
    
    AgilidadRaza(eRazaAi.Humano) = Val(Leer.GetValue("AgilidadRaza", "Humano"))
    AgilidadRaza(eRazaAi.Enano) = Val(Leer.GetValue("AgilidadRaza", "Enano"))
    AgilidadRaza(eRazaAi.Gnomo) = Val(Leer.GetValue("AgilidadRaza", "Gnomo"))
    AgilidadRaza(eRazaAi.Elfo) = Val(Leer.GetValue("AgilidadRaza", "Elfo"))
    AgilidadRaza(eRazaAi.Drow) = Val(Leer.GetValue("AgilidadRaza", "ElfoDrow"))
    
    BodyClaseRaza(eBotAi.Mago, eRazaAi.Humano) = Val(Leer.GetValue("BodyClaseRaza", "MagoH"))
    BodyClaseRaza(eBotAi.Mago, eRazaAi.Gnomo) = Val(Leer.GetValue("BodyClaseRaza", "MagoG"))
    BodyClaseRaza(eBotAi.Mago, eRazaAi.Elfo) = Val(Leer.GetValue("BodyClaseRaza", "MagoE"))
    BodyClaseRaza(eBotAi.Mago, eRazaAi.Enano) = Val(Leer.GetValue("BodyClaseRaza", "MagoEN"))
    BodyClaseRaza(eBotAi.Mago, eRazaAi.Drow) = Val(Leer.GetValue("BodyClaseRaza", "MagoEO"))
    BodyClaseRaza(eBotAi.Paladin, eRazaAi.Humano) = Val(Leer.GetValue("BodyClaseRaza", "PaladinH"))
    BodyClaseRaza(eBotAi.Paladin, eRazaAi.Gnomo) = Val(Leer.GetValue("BodyClaseRaza", "PaladinG"))
    BodyClaseRaza(eBotAi.Paladin, eRazaAi.Elfo) = Val(Leer.GetValue("BodyClaseRaza", "PaladinE"))
    BodyClaseRaza(eBotAi.Paladin, eRazaAi.Enano) = Val(Leer.GetValue("BodyClaseRaza", "PaladinEN"))
    BodyClaseRaza(eBotAi.Paladin, eRazaAi.Drow) = Val(Leer.GetValue("BodyClaseRaza", "PaladinEO"))
    BodyClaseRaza(eBotAi.Asesino, eRazaAi.Humano) = Val(Leer.GetValue("BodyClaseRaza", "AsesinoH"))
    BodyClaseRaza(eBotAi.Asesino, eRazaAi.Gnomo) = Val(Leer.GetValue("BodyClaseRaza", "AsesinoG"))
    BodyClaseRaza(eBotAi.Asesino, eRazaAi.Elfo) = Val(Leer.GetValue("BodyClaseRaza", "AsesinoE"))
    BodyClaseRaza(eBotAi.Asesino, eRazaAi.Enano) = Val(Leer.GetValue("BodyClaseRaza", "AsesinoEN"))
    BodyClaseRaza(eBotAi.Asesino, eRazaAi.Drow) = Val(Leer.GetValue("BodyClaseRaza", "AsesinoEO"))
    BodyClaseRaza(eBotAi.Bardo, eRazaAi.Humano) = Val(Leer.GetValue("BodyClaseRaza", "BardoH"))
    BodyClaseRaza(eBotAi.Bardo, eRazaAi.Gnomo) = Val(Leer.GetValue("BodyClaseRaza", "BardoG"))
    BodyClaseRaza(eBotAi.Bardo, eRazaAi.Elfo) = Val(Leer.GetValue("BodyClaseRaza", "BardoE"))
    BodyClaseRaza(eBotAi.Bardo, eRazaAi.Enano) = Val(Leer.GetValue("BodyClaseRaza", "BardoEN"))
    BodyClaseRaza(eBotAi.Bardo, eRazaAi.Drow) = Val(Leer.GetValue("BodyClaseRaza", "BardoEO"))
    BodyClaseRaza(eBotAi.Clerigo, eRazaAi.Humano) = Val(Leer.GetValue("BodyClaseRaza", "ClerigoH"))
    BodyClaseRaza(eBotAi.Clerigo, eRazaAi.Gnomo) = Val(Leer.GetValue("BodyClaseRaza", "ClerigoG"))
    BodyClaseRaza(eBotAi.Clerigo, eRazaAi.Elfo) = Val(Leer.GetValue("BodyClaseRaza", "ClerigoE"))
    BodyClaseRaza(eBotAi.Clerigo, eRazaAi.Enano) = Val(Leer.GetValue("BodyClaseRaza", "ClerigoEN"))
    BodyClaseRaza(eBotAi.Clerigo, eRazaAi.Drow) = Val(Leer.GetValue("BodyClaseRaza", "ClerigoEO"))
    BodyClaseRaza(eBotAi.Druida, eRazaAi.Humano) = Val(Leer.GetValue("BodyClaseRaza", "DruidaH"))
    BodyClaseRaza(eBotAi.Druida, eRazaAi.Gnomo) = Val(Leer.GetValue("BodyClaseRaza", "DruidaG"))
    BodyClaseRaza(eBotAi.Druida, eRazaAi.Elfo) = Val(Leer.GetValue("BodyClaseRaza", "DruidaE"))
    BodyClaseRaza(eBotAi.Druida, eRazaAi.Enano) = Val(Leer.GetValue("BodyClaseRaza", "DruidaEN"))
    BodyClaseRaza(eBotAi.Druida, eRazaAi.Drow) = Val(Leer.GetValue("BodyClaseRaza", "DruidaEO"))
    BodyClaseRaza(eBotAi.Guerrero, eRazaAi.Humano) = Val(Leer.GetValue("BodyClaseRaza", "GuerreroH"))
    BodyClaseRaza(eBotAi.Guerrero, eRazaAi.Gnomo) = Val(Leer.GetValue("BodyClaseRaza", "GuerreroG"))
    BodyClaseRaza(eBotAi.Guerrero, eRazaAi.Elfo) = Val(Leer.GetValue("BodyClaseRaza", "GuerreroE"))
    BodyClaseRaza(eBotAi.Guerrero, eRazaAi.Enano) = Val(Leer.GetValue("BodyClaseRaza", "GuerreroEN"))
    BodyClaseRaza(eBotAi.Guerrero, eRazaAi.Drow) = Val(Leer.GetValue("BodyClaseRaza", "GuerreroEO"))
    
    HeadRazaGenero(eRazaAi.Humano, 0) = Val(Leer.GetValue("HeadRazaGenero", "HumanoH"))
    HeadRazaGenero(eRazaAi.Humano, 1) = Val(Leer.GetValue("HeadRazaGenero", "HumanoM"))
    HeadRazaGenero(eRazaAi.Enano, 0) = Val(Leer.GetValue("HeadRazaGenero", "EnanoH"))
    HeadRazaGenero(eRazaAi.Enano, 1) = Val(Leer.GetValue("HeadRazaGenero", "EnanoM"))
    HeadRazaGenero(eRazaAi.Elfo, 0) = Val(Leer.GetValue("HeadRazaGenero", "ElfoH"))
    HeadRazaGenero(eRazaAi.Elfo, 1) = Val(Leer.GetValue("HeadRazaGenero", "ElfoM"))
    HeadRazaGenero(eRazaAi.Gnomo, 0) = Val(Leer.GetValue("HeadRazaGenero", "GnomoH"))
    HeadRazaGenero(eRazaAi.Gnomo, 1) = Val(Leer.GetValue("HeadRazaGenero", "GnomoM"))
    HeadRazaGenero(eRazaAi.Drow, 0) = Val(Leer.GetValue("HeadRazaGenero", "ElfoDrowH"))
    HeadRazaGenero(eRazaAi.Drow, 1) = Val(Leer.GetValue("HeadRazaGenero", "ElfoDrowM"))
    
    ArmaClase(eBotAi.Mago) = Val(Leer.GetValue("ArmaClase", "Mago"))
    ArmaClase(eBotAi.Paladin) = Val(Leer.GetValue("ArmaClase", "Paladin"))
    ArmaClase(eBotAi.Clerigo) = Val(Leer.GetValue("ArmaClase", "Clerigo"))
    ArmaClase(eBotAi.Asesino) = Val(Leer.GetValue("ArmaClase", "Asesino"))
    ArmaClase(eBotAi.Bardo) = Val(Leer.GetValue("ArmaClase", "Bardo"))
    ArmaClase(eBotAi.Druida) = Val(Leer.GetValue("ArmaClase", "Druida"))
    ArmaClase(eBotAi.Guerrero) = Val(Leer.GetValue("ArmaClase", "Guerrero"))
    
    CascoClase(eBotAi.Mago) = Val(Leer.GetValue("CascoClase", "Mago"))
    CascoClase(eBotAi.Paladin) = Val(Leer.GetValue("CascoClase", "Paladin"))
    CascoClase(eBotAi.Clerigo) = Val(Leer.GetValue("CascoClase", "Clerigo"))
    CascoClase(eBotAi.Asesino) = Val(Leer.GetValue("CascoClase", "Asesino"))
    CascoClase(eBotAi.Bardo) = Val(Leer.GetValue("CascoClase", "Bardo"))
    CascoClase(eBotAi.Druida) = Val(Leer.GetValue("CascoClase", "Druida"))
    CascoClase(eBotAi.Guerrero) = Val(Leer.GetValue("CascoClase", "Guerrero"))
    
    EscudoClase(eBotAi.Mago) = Val(Leer.GetValue("EscudoClase", "Mago"))
    EscudoClase(eBotAi.Paladin) = Val(Leer.GetValue("EscudoClase", "Paladin"))
    EscudoClase(eBotAi.Clerigo) = Val(Leer.GetValue("EscudoClase", "Clerigo"))
    EscudoClase(eBotAi.Asesino) = Val(Leer.GetValue("EscudoClase", "Asesino"))
    EscudoClase(eBotAi.Bardo) = Val(Leer.GetValue("EscudoClase", "Bardo"))
    EscudoClase(eBotAi.Druida) = Val(Leer.GetValue("EscudoClase", "Druida"))
    EscudoClase(eBotAi.Guerrero) = Val(Leer.GetValue("EscudoClase", "Guerrero"))
    
    MinArmaClase(eBotAi.Mago) = Val(Leer.GetValue("MinArmaClase", "Mago"))
    MinArmaClase(eBotAi.Paladin) = Val(Leer.GetValue("MinArmaClase", "Paladin"))
    MinArmaClase(eBotAi.Clerigo) = Val(Leer.GetValue("MinArmaClase", "Clerigo"))
    MinArmaClase(eBotAi.Asesino) = Val(Leer.GetValue("MinArmaClase", "Asesino"))
    MinArmaClase(eBotAi.Bardo) = Val(Leer.GetValue("MinArmaClase", "Bardo"))
    MinArmaClase(eBotAi.Druida) = Val(Leer.GetValue("MinArmaClase", "Druida"))
    MinArmaClase(eBotAi.Guerrero) = Val(Leer.GetValue("MinArmaClase", "Guerrero"))
    
    MaxArmaClase(eBotAi.Mago) = Val(Leer.GetValue("MaxArmaClase", "Mago"))
    MaxArmaClase(eBotAi.Paladin) = Val(Leer.GetValue("MaxArmaClase", "Paladin"))
    MaxArmaClase(eBotAi.Clerigo) = Val(Leer.GetValue("MaxArmaClase", "Clerigo"))
    MaxArmaClase(eBotAi.Asesino) = Val(Leer.GetValue("MaxArmaClase", "Asesino"))
    MaxArmaClase(eBotAi.Bardo) = Val(Leer.GetValue("MaxArmaClase", "Bardo"))
    MaxArmaClase(eBotAi.Druida) = Val(Leer.GetValue("MaxArmaClase", "Druida"))
    MaxArmaClase(eBotAi.Guerrero) = Val(Leer.GetValue("MaxArmaClase", "Guerrero"))
    
    MinDefBClase(eBotAi.Mago) = Val(Leer.GetValue("MinDefBClase", "Mago"))
    MinDefBClase(eBotAi.Paladin) = Val(Leer.GetValue("MinDefBClase", "Paladin"))
    MinDefBClase(eBotAi.Clerigo) = Val(Leer.GetValue("MinDefBClase", "Clerigo"))
    MinDefBClase(eBotAi.Asesino) = Val(Leer.GetValue("MinDefBClase", "Asesino"))
    MinDefBClase(eBotAi.Bardo) = Val(Leer.GetValue("MinDefBClase", "Bardo"))
    MinDefBClase(eBotAi.Druida) = Val(Leer.GetValue("MinDefBClase", "Druida"))
    MinDefBClase(eBotAi.Guerrero) = Val(Leer.GetValue("MinDefBClase", "Guerrero"))
    
    MaxDefBClase(eBotAi.Mago) = Val(Leer.GetValue("MaxDefBClase", "Mago"))
    MaxDefBClase(eBotAi.Paladin) = Val(Leer.GetValue("MaxDefBClase", "Paladin"))
    MaxDefBClase(eBotAi.Clerigo) = Val(Leer.GetValue("MaxDefBClase", "Clerigo"))
    MaxDefBClase(eBotAi.Asesino) = Val(Leer.GetValue("MaxDefBClase", "Asesino"))
    MaxDefBClase(eBotAi.Bardo) = Val(Leer.GetValue("MaxDefBClase", "Bardo"))
    MaxDefBClase(eBotAi.Druida) = Val(Leer.GetValue("MaxDefBClase", "Druida"))
    MaxDefBClase(eBotAi.Guerrero) = Val(Leer.GetValue("MaxDefBClase", "Guerrero"))
    
    MinDefHClase(eBotAi.Mago) = Val(Leer.GetValue("MinDefHClase", "Mago"))
    MinDefHClase(eBotAi.Paladin) = Val(Leer.GetValue("MinDefHClase", "Paladin"))
    MinDefHClase(eBotAi.Clerigo) = Val(Leer.GetValue("MinDefHClase", "Clerigo"))
    MinDefHClase(eBotAi.Asesino) = Val(Leer.GetValue("MinDefHClase", "Asesino"))
    MinDefHClase(eBotAi.Bardo) = Val(Leer.GetValue("MinDefHClase", "Bardo"))
    MinDefHClase(eBotAi.Druida) = Val(Leer.GetValue("MinDefHClase", "Druida"))
    
    MaxDefHClase(eBotAi.Mago) = Val(Leer.GetValue("MaxDefHClase", "Mago"))
    MaxDefHClase(eBotAi.Paladin) = Val(Leer.GetValue("MaxDefHClase", "Paladin"))
    MaxDefHClase(eBotAi.Clerigo) = Val(Leer.GetValue("MaxDefHClase", "Clerigo"))
    MaxDefHClase(eBotAi.Asesino) = Val(Leer.GetValue("MaxDefHClase", "Asesino"))
    MaxDefHClase(eBotAi.Bardo) = Val(Leer.GetValue("MaxDefHClase", "Bardo"))
    MaxDefHClase(eBotAi.Druida) = Val(Leer.GetValue("MaxDefHClase", "Druida"))
    MaxDefHClase(eBotAi.Guerrero) = Val(Leer.GetValue("MaxDefHClase", "Guerrero"))
    
    MinRMClase(eBotAi.Mago) = Val(Leer.GetValue("MinRMClase", "Mago"))
    MinRMClase(eBotAi.Paladin) = Val(Leer.GetValue("MinRMClase", "Paladin"))
    MinRMClase(eBotAi.Clerigo) = Val(Leer.GetValue("MinRMClase", "Clerigo"))
    MinRMClase(eBotAi.Asesino) = Val(Leer.GetValue("MinRMClase", "Asesino"))
    MinRMClase(eBotAi.Bardo) = Val(Leer.GetValue("MinRMClase", "Bardo"))
    MinRMClase(eBotAi.Druida) = Val(Leer.GetValue("MinRMClase", "Druida"))
    MinRMClase(eBotAi.Guerrero) = Val(Leer.GetValue("MinRMClase", "Guerrero"))
    
    MaxRMClase(eBotAi.Mago) = Val(Leer.GetValue("MaxRMClase", "Mago"))
    MaxRMClase(eBotAi.Paladin) = Val(Leer.GetValue("MaxRMClase", "Paladin"))
    MaxRMClase(eBotAi.Clerigo) = Val(Leer.GetValue("MaxRMClase", "Clerigo"))
    MaxRMClase(eBotAi.Asesino) = Val(Leer.GetValue("MaxRMClase", "Asesino"))
    MaxRMClase(eBotAi.Bardo) = Val(Leer.GetValue("MaxRMClase", "Bardo"))
    MaxRMClase(eBotAi.Druida) = Val(Leer.GetValue("MaxRMClase", "Druida"))
    MaxRMClase(eBotAi.Guerrero) = Val(Leer.GetValue("MaxRMClase", "Guerrero"))
    
    ModEvasionClase(eBotAi.Mago) = Val(Leer.GetValue("ModEvasionClase", "Mago"))
    ModEvasionClase(eBotAi.Paladin) = Val(Leer.GetValue("ModEvasionClase", "Paladin"))
    ModEvasionClase(eBotAi.Clerigo) = Val(Leer.GetValue("ModEvasionClase", "Clerigo"))
    ModEvasionClase(eBotAi.Asesino) = Val(Leer.GetValue("ModEvasionClase", "Asesino"))
    ModEvasionClase(eBotAi.Bardo) = Val(Leer.GetValue("ModEvasionClase", "Bardo"))
    ModEvasionClase(eBotAi.Druida) = Val(Leer.GetValue("ModEvasionClase", "Druida"))
    ModEvasionClase(eBotAi.Guerrero) = Val(Leer.GetValue("ModEvasionClase", "Guerrero"))
    
    ModAtaqueArmaClase(eBotAi.Mago) = Val(Leer.GetValue("ModAtaqueArmaClase", "Mago"))
    ModAtaqueArmaClase(eBotAi.Paladin) = Val(Leer.GetValue("ModAtaqueArmaClase", "Paladin"))
    ModAtaqueArmaClase(eBotAi.Clerigo) = Val(Leer.GetValue("ModAtaqueArmaClase", "Clerigo"))
    ModAtaqueArmaClase(eBotAi.Asesino) = Val(Leer.GetValue("ModAtaqueArmaClase", "Asesino"))
    ModAtaqueArmaClase(eBotAi.Bardo) = Val(Leer.GetValue("ModAtaqueArmaClase", "Bardo"))
    ModAtaqueArmaClase(eBotAi.Druida) = Val(Leer.GetValue("ModAtaqueArmaClase", "Druida"))
    ModAtaqueArmaClase(eBotAi.Guerrero) = Val(Leer.GetValue("ModAtaqueArmaClase", "Guerrero"))
    
    ModAtaqueProyectilClase(eBotAi.Mago) = Val(Leer.GetValue("ModAtaqueArmaClase", "Mago"))
    ModAtaqueProyectilClase(eBotAi.Paladin) = Val(Leer.GetValue("ModAtaqueArmaClase", "Paladin"))
    ModAtaqueProyectilClase(eBotAi.Clerigo) = Val(Leer.GetValue("ModAtaqueArmaClase", "Clerigo"))
    ModAtaqueProyectilClase(eBotAi.Asesino) = Val(Leer.GetValue("ModAtaqueArmaClase", "Asesino"))
    ModAtaqueProyectilClase(eBotAi.Bardo) = Val(Leer.GetValue("ModAtaqueArmaClase", "Bardo"))
    ModAtaqueProyectilClase(eBotAi.Druida) = Val(Leer.GetValue("ModAtaqueArmaClase", "Druida"))
    ModAtaqueProyectilClase(eBotAi.Guerrero) = Val(Leer.GetValue("ModAtaqueArmaClase", "Guerrero"))
    
    ModDañoArmaClase(eBotAi.Mago) = Val(Leer.GetValue("ModDañoArmaClase", "Mago"))
    ModDañoArmaClase(eBotAi.Paladin) = Val(Leer.GetValue("ModDañoArmaClase", "Paladin"))
    ModDañoArmaClase(eBotAi.Clerigo) = Val(Leer.GetValue("ModDañoArmaClase", "Clerigo"))
    ModDañoArmaClase(eBotAi.Asesino) = Val(Leer.GetValue("ModDañoArmaClase", "Asesino"))
    ModDañoArmaClase(eBotAi.Bardo) = Val(Leer.GetValue("ModDañoArmaClase", "Bardo"))
    ModDañoArmaClase(eBotAi.Druida) = Val(Leer.GetValue("ModDañoArmaClase", "Druida"))
    ModDañoArmaClase(eBotAi.Guerrero) = Val(Leer.GetValue("ModDañoArmaClase", "Guerrero"))
    
    ModDañoProyectilClase(eBotAi.Mago) = Val(Leer.GetValue("ModDañoArmaClase", "Mago"))
    ModDañoProyectilClase(eBotAi.Paladin) = Val(Leer.GetValue("ModDañoArmaClase", "Paladin"))
    ModDañoProyectilClase(eBotAi.Clerigo) = Val(Leer.GetValue("ModDañoArmaClase", "Clerigo"))
    ModDañoProyectilClase(eBotAi.Asesino) = Val(Leer.GetValue("ModDañoArmaClase", "Asesino"))
    ModDañoProyectilClase(eBotAi.Bardo) = Val(Leer.GetValue("ModDañoArmaClase", "Bardo"))
    ModDañoProyectilClase(eBotAi.Druida) = Val(Leer.GetValue("ModDañoArmaClase", "Druida"))
    ModDañoProyectilClase(eBotAi.Guerrero) = Val(Leer.GetValue("ModDañoArmaClase", "Guerrero"))
    
    ModEscudoClase(eBotAi.Mago) = Val(Leer.GetValue("ModEscudoClase", "Mago"))
    ModEscudoClase(eBotAi.Paladin) = Val(Leer.GetValue("ModEscudoClase", "Paladin"))
    ModEscudoClase(eBotAi.Clerigo) = Val(Leer.GetValue("ModEscudoClase", "Clerigo"))
    ModEscudoClase(eBotAi.Asesino) = Val(Leer.GetValue("ModEscudoClase", "Asesino"))
    ModEscudoClase(eBotAi.Bardo) = Val(Leer.GetValue("ModEscudoClase", "Bardo"))
    ModEscudoClase(eBotAi.Druida) = Val(Leer.GetValue("ModEscudoClase", "Druida"))
    ModEscudoClase(eBotAi.Guerrero) = Val(Leer.GetValue("ModEscudoClase", "Guerrero"))
    
    ModDMClase(eBotAi.Mago) = Val(Leer.GetValue("ModDMClase", "Mago"))
    ModDMClase(eBotAi.Paladin) = Val(Leer.GetValue("ModDMClase", "Paladin"))
    ModDMClase(eBotAi.Clerigo) = Val(Leer.GetValue("ModDMClase", "Clerigo"))
    ModDMClase(eBotAi.Asesino) = Val(Leer.GetValue("ModDMClase", "Asesino"))
    ModDMClase(eBotAi.Bardo) = Val(Leer.GetValue("ModDMClase", "Bardo"))
    ModDMClase(eBotAi.Druida) = Val(Leer.GetValue("ModDMClase", "Druida"))
    ModDMClase(eBotAi.Guerrero) = Val(Leer.GetValue("ModDMClase", "Guerrero"))
    
    For i = 1 To 6
        Hechizo(i).name = Leer.GetValue("Hechizo" & i, "Nombre")
        Hechizo(i).Mana = Val(Leer.GetValue("Hechizo" & i, "Mana"))
        Hechizo(i).Sta = Val(Leer.GetValue("Hechizo" & i, "Sta"))
        Hechizo(i).MinHP = Val(Leer.GetValue("Hechizo" & i, "MinHP"))
        Hechizo(i).MaxHP = Val(Leer.GetValue("Hechizo" & i, "MaxHP"))
        Hechizo(i).WAV = Val(Leer.GetValue("Hechizo" & i, "WAV"))
        Hechizo(i).FX = Val(Leer.GetValue("Hechizo" & i, "FX"))
        Hechizo(i).Palabras = Leer.GetValue("Hechizo" & i, "Palabras")
    Next i
    
    ItemData(1).MinHit = Val(Leer.GetValue("Item1", "MinHit"))
    ItemData(1).MaxHit = Val(Leer.GetValue("Item1", "MaxHit"))
    ItemData(2).MinHit = Val(Leer.GetValue("Item2", "MinHit"))
    ItemData(2).MaxHit = Val(Leer.GetValue("Item2", "MaxHit"))
    ItemData(2).Refuerzo = Val(Leer.GetValue("Item2", "Refuerzo"))
    ItemData(3).MinRM = Val(Leer.GetValue("Item3", "MinRM"))
    ItemData(3).MaxRM = Val(Leer.GetValue("Item3", "MaxRM"))
    ItemData(3).DM = Val(Leer.GetValue("Item3", "DM"))
    ItemData(4).MinRM = Val(Leer.GetValue("Item4", "MinRM"))
    ItemData(4).MaxRM = Val(Leer.GetValue("Item4", "MaxRM"))
    ItemData(4).DM = Val(Leer.GetValue("Item4", "DM"))
    Set Leer = Nothing
End Sub

Public Function RandomNum(ByVal max As Integer) As Integer
'---------------------------------------------------------------------------------------
' Procedure : RandomNum
' Author    : Anagrama
' Date      : ???
' Purpose   : Una funcion mucho mas aleatoria de la que habia.
'---------------------------------------------------------------------------------------
'
    Randomize

    RandomNum = CInt(Int((max * Rnd()) + 1))
End Function

Public Sub CharDie(ByVal CharIndex As Integer)
'---------------------------------------------------------------------------------------
' Procedure : CharDie
' Author    : Anagrama
' Date      : ???
' Purpose   : ¿El personaje muere?.
'---------------------------------------------------------------------------------------
'
    Dim i As Byte
    
    If charlist(CharIndex).Genero = 0 Then
        Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(11, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
    Else
        Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(74, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
    End If
    If EnTeam(CharIndex) = 2 Then
        For i = 1 To UBound(Team2)
            If charlist(i).Bot = 1 Then
                Call CharTimer(Team2(i)).SetInterval(TimersIndex.Resu, INT_RESU + RandomNumber(100 * Dificultad, 200 * Dificultad))
                Call CharTimer(Team2(i)).Restart(TimersIndex.Resu)
            End If
        Next i
    Else
        If UBound(Team1) > 1 Then
            For i = 1 To UBound(Team1)
                If charlist(i).Bot = 1 Then
                    Call CharTimer(Team1(i)).SetInterval(TimersIndex.Resu, INT_RESU + RandomNumber(100 * Dificultad, 200 * Dificultad))
                    Call CharTimer(Team1(i)).Restart(TimersIndex.Resu)
                End If
            Next i
        End If
    End If
    If charlist(CharIndex).Bot = 0 And charlist(CharIndex).Inmo = 1 Then
        charlist(CharIndex).Inmo = 0
        Call WriteParalizeOK(CharIndex, charlist(CharIndex).Inmo)
    End If
    charlist(CharIndex).Inmo = 0
    charlist(CharIndex).body = BodyData(7)
    charlist(CharIndex).Head = HeadData(11)
    charlist(CharIndex).Arma = WeaponAnimData(2)
    charlist(CharIndex).Casco = CascoAnimData(2)
    charlist(CharIndex).Escudo = ShieldAnimData(2)
    Call ServerSendData(SendTarget.ToAllButIndex, UserCharIndex, PrepareMessageCharacterChange(7, _
                                    11, charlist(CharIndex).heading, CharIndex, 2, _
                                    2, charlist(CharIndex).FxIndex, charlist(CharIndex).FX.Loops, 2))
End Sub

Public Sub EquiparItem(ByVal CharIndex As Integer, ByVal Slot As Byte)
'---------------------------------------------------------------------------------------
' Procedure : EquiparItem
' Author    : Anagrama
' Date      : ???
' Purpose   : Equipa el item seleccionado.
'---------------------------------------------------------------------------------------
'

    If Slot = 0 Then Exit Sub
    If Inventario.OBJIndex(Slot) < 3 Then Exit Sub
    Select Case Inventario.OBJIndex(Slot)
        Case 3
            Call Inventario.SetItem(3, 3, 0, 1, 601, 1, 0, 0, 0, 0, 0, "Hacha de Guerra Dos Filos")
            Call Inventario.SetItem(2, 4, 0, 0, 602, 1, 0, 0, 0, 0, 0, "Arco de Cazador")
            Call Inventario.SetItem(5, 4, 0, 0, 602, 1, 0, 0, 0, 0, 0, "Arco de Cazador")
        Case 4
            Call Inventario.SetItem(3, 3, 0, 0, 601, 1, 0, 0, 0, 0, 0, "Hacha de Guerra Dos Filos")
            Call Inventario.SetItem(2, 4, 0, 1, 602, 1, 0, 0, 0, 0, 0, "Arco de Cazador")
            Call Inventario.SetItem(5, 4, 0, 1, 602, 1, 0, 0, 0, 0, 0, "Arco de Cazador")
        Case 5
            Call Inventario.SetItem(3, 5, 0, 1, 603, 1, 0, 0, 0, 0, 0, "Laúd Élfico")
            Call Inventario.SetItem(4, 6, 0, 0, 604, 1, 0, 0, 0, 0, 0, "Anillo de Disolución Mágica")
        Case 6
            Call Inventario.SetItem(3, 5, 0, 0, 603, 1, 0, 0, 0, 0, 0, "Laúd Élfico")
            Call Inventario.SetItem(4, 6, 0, 1, 604, 1, 0, 0, 0, 0, 0, "Anillo de Disolución Mágica")
    End Select
    Call WriteEquipItem(Slot)
End Sub

Public Sub LimpiarPrioridadTarget(ByVal Team As Byte)
'---------------------------------------------------------------------------------------
' Procedure : LimpiarPrioridadTarget
' Author    : Anagrama
' Date      : ???
' Purpose   : Limpia el array de prioridades del team correspondiente.
'---------------------------------------------------------------------------------------
'
    Dim i As Byte
    
    If Team = 1 Then
        For i = 1 To 5
            Prioridad1(i).Char = 0
            Prioridad1(i).Probabilidad = 0
        Next
    Else
        For i = 1 To 5
            Prioridad2(i).Char = 0
            Prioridad2(i).Probabilidad = 0
        Next
    End If
End Sub

Public Sub DarPrioridadTarget(ByVal Team As Byte)
'---------------------------------------------------------------------------------------
' Procedure : DarPrioridadTarget
' Author    : Anagrama
' Date      : ???
' Purpose   : Llena los array de orden de prioridad de ataque segun el team y llama a calcular la probabilidad de target.
'---------------------------------------------------------------------------------------
'
    Dim i As Byte
    Dim a As Byte
    
    Call LimpiarPrioridadTarget(Team)
    
    If Team = 1 Then
        If UBound(Team1) > 0 Then
            For i = 1 To UBound(Team1)
                a = 1
                Do While a <= 5
                    If Prioridad1(a).Char > 0 Then
                        If PrioridadClase(charlist(Team1(i)).Ai) < PrioridadClase(charlist(Prioridad1(a).Char).Ai) Then
                            If a < 5 Then Call PushPrioridad(1, a)
                            Prioridad1(a).Char = Team1(i)
                            Exit Do
                        ElseIf PrioridadClase(charlist(Team1(i)).Ai) = PrioridadClase(charlist(Prioridad1(a).Char).Ai) Then
                            If PrioridadRaza(charlist(Team1(i)).Raza) < PrioridadRaza(charlist(Prioridad1(a).Char).Raza) Then
                                If a < 5 Then Call PushPrioridad(1, a)
                                Prioridad1(a).Char = Team1(i)
                                Exit Do
                            End If
                        End If
                    Else
                        Prioridad1(a).Char = Team1(i)
                        Exit Do
                    End If
                    a = a + 1
                Loop
            Next
            Call CalcProbTarget(1)
        End If
    Else
        If UBound(Team2) > 0 Then
            For i = 1 To UBound(Team2)
                a = 1
                Do While a <= 5
                    If Prioridad2(a).Char > 0 Then
                        If PrioridadClase(charlist(Team2(i)).Ai) < PrioridadClase(charlist(Prioridad2(a).Char).Ai) Then
                            If a < 5 Then Call PushPrioridad(2, a)
                            Prioridad2(a).Char = Team2(i)
                            Exit Do
                        ElseIf PrioridadClase(charlist(Team2(i)).Ai) = PrioridadClase(charlist(Prioridad2(a).Char).Ai) Then
                            If PrioridadRaza(charlist(Team2(i)).Raza) < PrioridadRaza(charlist(Prioridad2(a).Char).Raza) Then
                                If a < 5 Then Call PushPrioridad(2, a)
                                Prioridad2(a).Char = Team2(i)
                                Exit Do
                            End If
                        End If
                    Else
                        Prioridad2(a).Char = Team2(i)
                        Exit Do
                    End If
                    a = a + 1
                Loop
            Next
            Call CalcProbTarget(2)
        End If
    End If
    
End Sub

Public Sub PushPrioridad(ByVal Team As Byte, ByVal Finish As Byte)
'---------------------------------------------------------------------------------------
' Procedure : PushPrioridad
' Author    : Anagrama
' Date      : ???
' Purpose   : Empuja el orden de prioridad para atras dejando un lugar abuerto donde se lo solicita.
'---------------------------------------------------------------------------------------
'
    Dim i As Integer
    
    If Team = 1 Then
        For i = 5 To Finish + 1 Step -1
            Prioridad1(i).Char = Prioridad1(i - 1).Char
            Prioridad1(i).Probabilidad = Prioridad1(i - 1).Probabilidad
        Next
    Else
        For i = 5 To Finish + 1 Step -1
            Prioridad2(i).Char = Prioridad2(i - 1).Char
            Prioridad2(i).Probabilidad = Prioridad2(i - 1).Probabilidad
        Next
    End If
End Sub

Public Sub CalcProbTarget(ByVal Team As Byte)
'---------------------------------------------------------------------------------------
' Procedure : CalcProbTarget
' Author    : Anagrama
' Date      : ???
' Purpose   : Calcula la probabilidad de focusear a un personaje en base al orden de prioridad y la prioridad del segundo.
'---------------------------------------------------------------------------------------
'
    Dim i As Byte
    Dim a As Byte
    Dim Diferencia As Byte
    
    If Team = 1 Then
        For i = 1 To UBound(Team1)
            If charlist(Prioridad1(i).Char).MinHP > 0 Then
                If i < UBound(Team1) Then
                    a = i + 1
                    Do While a <= UBound(Team1)
                        If charlist(Prioridad1(a).Char).MinHP > 0 Then
                            Diferencia = PrioridadClase(charlist(Prioridad1(a).Char).Ai) - PrioridadClase(charlist(Prioridad1(i).Char).Ai)
                            If Diferencia = 0 Then 'Son la misma clase
                                Prioridad1(i).Probabilidad = 60
                                Prioridad1(a).Probabilidad = 30
                            ElseIf Diferencia = 1 Then
                                Prioridad1(i).Probabilidad = 65
                                Prioridad1(a).Probabilidad = 25
                            ElseIf Diferencia = 2 Then
                                Prioridad1(i).Probabilidad = 75
                                Prioridad1(a).Probabilidad = 15
                            ElseIf Diferencia = 3 Then
                                Prioridad1(i).Probabilidad = 80
                                Prioridad1(a).Probabilidad = 10
                            ElseIf Diferencia = 4 Then
                                Prioridad1(i).Probabilidad = 85
                                Prioridad1(a).Probabilidad = 10
                            ElseIf Diferencia = 5 Then
                                Prioridad1(i).Probabilidad = 90
                                Prioridad1(a).Probabilidad = 5
                            ElseIf Diferencia = 6 Then
                                Prioridad1(i).Probabilidad = 100
                                Prioridad1(a).Probabilidad = 0
                            End If
                            Exit Sub
                        End If
                        a = a + 1
                    Loop
                Else
                    Prioridad1(i).Probabilidad = 100
                End If
            End If
        Next
    Else
        For i = 1 To UBound(Team2)
            If charlist(Prioridad2(i).Char).MinHP > 0 Then
                If i < UBound(Team1) Then
                    a = i + 1
                    Do While a <= UBound(Team2)
                        If charlist(Prioridad2(a).Char).MinHP > 0 Then
                            Diferencia = PrioridadClase(charlist(Prioridad2(a).Char).Ai) - PrioridadClase(charlist(Prioridad2(i).Char).Ai)
                            If Diferencia = 0 Then 'Son la misma clase
                                Prioridad2(i).Probabilidad = 60
                                Prioridad2(a).Probabilidad = 30
                            ElseIf Diferencia = 1 Then
                                Prioridad2(i).Probabilidad = 65
                                Prioridad2(a).Probabilidad = 25
                            ElseIf Diferencia = 2 Then
                                Prioridad2(i).Probabilidad = 75
                                Prioridad2(a).Probabilidad = 15
                            ElseIf Diferencia = 3 Then
                                Prioridad2(i).Probabilidad = 80
                                Prioridad2(a).Probabilidad = 10
                            ElseIf Diferencia = 4 Then
                                Prioridad2(i).Probabilidad = 85
                                Prioridad2(a).Probabilidad = 10
                            ElseIf Diferencia = 5 Then
                                Prioridad2(i).Probabilidad = 90
                                Prioridad2(a).Probabilidad = 5
                            ElseIf Diferencia = 6 Then
                                Prioridad2(i).Probabilidad = 100
                                Prioridad2(a).Probabilidad = 0
                            End If
                            Exit Sub
                        End If
                        a = a + 1
                    Loop
                Else
                    Prioridad2(i).Probabilidad = 100
                End If
            End If
        Next
    End If
End Sub

Public Sub VerTile(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
    If MapData(X, Y + 1).CharIndex > 0 Then
        If EnTeam(MapData(X, Y + 1).CharIndex) = 1 Then
            Call WriteConsoleMsg(CharIndex, "Ves a " & charlist(MapData(X, Y + 1).CharIndex).Nombre & ".", FontTypeNames.FONTTYPE_CITIZEN)
        Else
            Call WriteConsoleMsg(CharIndex, "Ves a " & charlist(MapData(X, Y + 1).CharIndex).Nombre & ".", FontTypeNames.FONTTYPE_FIGHT)
        End If
    ElseIf MapData(X, Y).CharIndex > 0 Then
        If EnTeam(MapData(X, Y).CharIndex) = 1 Then
            Call WriteConsoleMsg(CharIndex, "Ves a " & charlist(MapData(X, Y).CharIndex).Nombre & ".", FontTypeNames.FONTTYPE_CITIZEN)
        Else
            Call WriteConsoleMsg(CharIndex, "Ves a " & charlist(MapData(X, Y).CharIndex).Nombre & ".", FontTypeNames.FONTTYPE_FIGHT)
        End If
    Else
        Call WriteConsoleMsg(CharIndex, "No ves nada interesante.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Public Sub UsarInvItem(ByVal CharIndex As Integer, ByVal Item As Byte)
    If pausa Or EnCuenta Then Exit Sub
    If charlist(CharIndex).MinHP = 0 Then
        Call WriteConsoleMsg(CharIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    Select Case Item
        Case 1
            charlist(CharIndex).MinHP = charlist(CharIndex).MinHP + 30
            If charlist(CharIndex).MinHP >= charlist(CharIndex).MaxHP Then charlist(CharIndex).MinHP = charlist(CharIndex).MaxHP
            Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(46, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
            Call WriteUpdateCharStats(CharIndex)
        Case 2
            charlist(CharIndex).MinMAN = charlist(CharIndex).MinMAN + (charlist(CharIndex).MaxMAN * 4) / 100 + charlist(CharIndex).Lvl \ 2 + 40 / charlist(CharIndex).Lvl
            If charlist(CharIndex).MinMAN >= charlist(CharIndex).MaxMAN Then charlist(CharIndex).MinMAN = charlist(CharIndex).MaxMAN
            Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(46, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
            Call WriteUpdateCharStats(CharIndex)
        Case 4
            If charlist(CharIndex).TipoArma = 1 Then
                Call WriteConsoleMsg(CharIndex, "Debes equipar el arco primero.", FontTypeNames.FONTTYPE_INFO)
            End If
    End Select
End Sub

Public Sub IniTimers(ByVal CharIndex As Integer)
    Call CharTimer(CharIndex).Start(TimersIndex.Attack)
    Call CharTimer(CharIndex).Start(TimersIndex.UseItemWithU)
    Call CharTimer(CharIndex).Start(TimersIndex.UseItemWithDblClick)
    Call CharTimer(CharIndex).Start(TimersIndex.SendRPU)
    Call CharTimer(CharIndex).Start(TimersIndex.CastSpell)
    Call CharTimer(CharIndex).Start(TimersIndex.Arrows)
    Call CharTimer(CharIndex).Start(TimersIndex.CastAttack)
    Call CharTimer(CharIndex).Start(TimersIndex.AttackCast)
    Call CharTimer(CharIndex).Start(TimersIndex.UseItem)
    Call CharTimer(CharIndex).Start(TimersIndex.GolpeU)

    Call CharTimer(CharIndex).SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call CharTimer(CharIndex).SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call CharTimer(CharIndex).SetInterval(TimersIndex.UseItem, INT_MINU)
    Call CharTimer(CharIndex).SetInterval(TimersIndex.GolpeU, INT_GOLPEU)
    Call CharTimer(CharIndex).SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call CharTimer(CharIndex).SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call CharTimer(CharIndex).SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call CharTimer(CharIndex).SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call CharTimer(CharIndex).SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    Call CharTimer(CharIndex).SetInterval(TimersIndex.AttackCast, INT_ATTACK_CAST)
End Sub

Public Sub EquiparInvItem(ByVal CharIndex As Integer, ByVal Slot As Byte)
'---------------------------------------------------------------------------------------
' Procedure : EquiparInvItem
' Author    : Anagrama
' Date      : ???
' Purpose   : Equipa el item seleccionado y envía el cambio a los usuarios.
'---------------------------------------------------------------------------------------
'

    If charlist(CharIndex).Ai = eBotAi.Guerrero Then
        Select Case Slot
            Case 3
                charlist(CharIndex).Arma = WeaponAnimData(3)
                charlist(CharIndex).iArma = 3
                charlist(CharIndex).ArmaMinHit = ItemData(1).MinHit
                charlist(CharIndex).ArmaMaxHit = ItemData(1).MaxHit
                charlist(CharIndex).Refuerzo = 0
                charlist(CharIndex).TipoArma = 1
                Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(25, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
                Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessageCharacterChange(charlist(CharIndex).iBody, _
                                    charlist(CharIndex).iHead, charlist(CharIndex).heading, CharIndex, charlist(CharIndex).iArma, _
                                    charlist(CharIndex).iEscudo, charlist(CharIndex).FxIndex, charlist(CharIndex).FX.Loops, charlist(CharIndex).iCasco))
            Case 2, 5
                charlist(CharIndex).Arma = WeaponAnimData(5)
                charlist(CharIndex).iArma = 5
                charlist(CharIndex).ArmaMinHit = ItemData(2).MinHit
                charlist(CharIndex).ArmaMaxHit = ItemData(2).MaxHit
                charlist(CharIndex).Refuerzo = ItemData(2).Refuerzo
                charlist(CharIndex).TipoArma = 2
                Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(25, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
                Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessageCharacterChange(charlist(CharIndex).iBody, _
                                    charlist(CharIndex).iHead, charlist(CharIndex).heading, CharIndex, charlist(CharIndex).iArma, _
                                    charlist(CharIndex).iEscudo, charlist(CharIndex).FxIndex, charlist(CharIndex).FX.Loops, charlist(CharIndex).iCasco))
        End Select
    ElseIf charlist(CharIndex).Ai = eBotAi.Bardo Then
        Select Case Slot
            Case 3
                charlist(CharIndex).MinRM = ItemData(3).MinRM
                charlist(CharIndex).MaxRM = ItemData(3).MaxRM
                charlist(CharIndex).DM = ItemData(3).DM
            Case 4
                charlist(CharIndex).MinRM = ItemData(4).MinRM
                charlist(CharIndex).MaxRM = ItemData(4).MaxRM
                charlist(CharIndex).DM = ItemData(4).DM
        End Select
    End If
End Sub

Public Function NextOpenTeamIndex() As Byte
    Dim i As Byte
    Dim tmpIndex As Byte
    Dim FoundIt As Byte
    
    tmpIndex = NextOpenUser
    
    Do While FoundIt = 0
        If UBound(Team1) > 0 Then
            For i = 1 To UBound(Team1)
                If TeamData1(i).index = tmpIndex Then
                    tmpIndex = tmpIndex + 1
                    FoundIt = 0
                    Exit For
                End If
                If i = UBound(Team1) Then FoundIt = 1
            Next i
        End If
        For i = 1 To UBound(Team2)
            If TeamData2(i).index = tmpIndex Then
                tmpIndex = tmpIndex + 1
                FoundIt = 0
                Exit For
            End If
            If i = UBound(Team2) And UBound(Team1) = 0 Then FoundIt = 1
        Next i
    Loop
    
    NextOpenTeamIndex = tmpIndex
End Function

Public Sub PushTeamData(ByVal Team As Byte, ByVal Limit As Byte)
    Dim i As Byte
    If Team = 1 Then
        For i = Limit To 4
            TeamData1(i).Clase = TeamData1(i + 1).Clase
            TeamData1(i).Raza = TeamData1(i + 1).Raza
            TeamData1(i).Genero = TeamData1(i + 1).Genero
            TeamData1(i).Nivel = TeamData1(i + 1).Nivel
            TeamData1(i).Bot = TeamData1(i + 1).Bot
            TeamData1(i).index = TeamData1(i + 1).index
            TeamData1(i).Nombre = TeamData1(i + 1).Nombre
        Next i
    Else
        For i = Limit To 4
            TeamData2(i).Clase = TeamData2(i + 1).Clase
            TeamData2(i).Raza = TeamData2(i + 1).Raza
            TeamData2(i).Genero = TeamData2(i + 1).Genero
            TeamData2(i).Nivel = TeamData2(i + 1).Nivel
            TeamData2(i).Bot = TeamData2(i + 1).Bot
            TeamData2(i).index = TeamData2(i + 1).index
            TeamData2(i).Nombre = TeamData2(i + 1).Nombre
        Next i
    End If
End Sub

Public Sub CambiarPj(ByVal CharIndex As Integer, ByVal Clase As Byte, ByVal Raza As Byte, ByVal Genero As Byte, ByVal Nivel As Byte, ByVal Team As Byte)
    Dim TeamIndex As Byte
    
    If EnTeam(CharIndex) = 1 Then
        TeamIndex = GetTeamIndex(CharIndex, 1)
        If Team <> 1 Then
            Call PushTeamData(1, TeamIndex)
            TeamData2(UBound(Team2) + 1).Clase = Clase
            TeamData2(UBound(Team2) + 1).Raza = Raza
            TeamData2(UBound(Team2) + 1).Genero = Genero
            TeamData2(UBound(Team2) + 1).Nivel = Nivel
            TeamData2(UBound(Team2) + 1).index = CharIndex
            TeamData2(UBound(Team2) + 1).Nombre = charlist(CharIndex).Nombre
            Call ResetDuelo(-1, UBound(Team1) - 1, UBound(Team2) + 1)
        Else
            TeamData1(TeamIndex).Clase = Clase
            TeamData1(TeamIndex).Raza = Raza
            TeamData1(TeamIndex).Genero = Genero
            TeamData1(TeamIndex).Nivel = Nivel
            Call ResetDuelo(-1, UBound(Team1), UBound(Team2))
        End If
    Else
        TeamIndex = GetTeamIndex(CharIndex, 2)
        If Team <> 2 Then
            Call PushTeamData(2, TeamIndex)
            TeamData1(UBound(Team1) + 1).Clase = Clase
            TeamData1(UBound(Team1) + 1).Raza = Raza
            TeamData1(UBound(Team1) + 1).Genero = Genero
            TeamData1(UBound(Team1) + 1).Nivel = Nivel
            TeamData1(UBound(Team1) + 1).index = CharIndex
            TeamData1(UBound(Team1) + 1).Nombre = charlist(CharIndex).Nombre
            Call ResetDuelo(-1, UBound(Team1) + 1, UBound(Team2) - 1)
        Else
            TeamData2(TeamIndex).Clase = Clase
            TeamData2(TeamIndex).Raza = Raza
            TeamData2(TeamIndex).Genero = Genero
            TeamData2(TeamIndex).Nivel = Nivel
            Call ResetDuelo(-1, UBound(Team1), UBound(Team2))
        End If
    End If
End Sub
