Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Protocol.bas
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517

Option Explicit

''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private auxiliarBuffer As New clsByteQueue

Private Type tFont
    red As Byte
    green As Byte
    blue As Byte
    bold As Boolean
    italic As Boolean
End Type

Private Enum ServerPacketID
    Logged                  ' LOGGED
    Disconnect              ' FINOK
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    ChangeMap               ' CM
    PosUpdate               ' PU
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    CharIndexInServer       ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterChangeNick
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    PlayMIDI                ' TM
    PlayWave                ' TW
    AreaChanged             ' CA
    PauseToggle             ' BKW
    CreateFX                ' CFX
    UpdateCharStats         ' EST
    ErrorMsg                ' ERR
    SetInvisible            ' NOVER
    ParalizeOK              ' PARADOK
    Pong
    CuentaToggle
End Enum

Private Enum ClientPacketID
    LoginChar               'OLOGIN
    Talk                    ';
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    CastSpell               'LH
    LeftClick               'LC
    UseItem                 'USA
    WorkLeftClick           'WLC
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    Online                  '/ONLINE
    Quit                    '/SALIR
    Ping                    '/PING
    LanzaFlecha
    ChangePj
End Enum

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
End Enum

Public FontTypes(20) As tFont

''
' Initializes the fonts array

Public Sub InitFonts()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .red = 255
        .green = 255
        .blue = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .red = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .red = 32
        .green = 51
        .blue = 223
        .bold = 1
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .red = 65
        .green = 190
        .blue = 156
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .red = 65
        .green = 190
        .blue = 156
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .red = 130
        .green = 130
        .blue = 130
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .red = 255
        .green = 180
        .blue = 250
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_VENENO).green = 255
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .red = 255
        .green = 255
        .blue = 255
        .bold = 1
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_SERVER).green = 185
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .red = 228
        .green = 199
        .blue = 27
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        .red = 130
        .green = 130
        .blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .red = 255
        .green = 60
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
        .green = 200
        .blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .red = 255
        .green = 50
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .green = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
        .red = 255
        .green = 255
        .blue = 255
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GM)
        .red = 30
        .green = 255
        .blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
        .blue = 200
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSE)
        .red = 30
        .green = 150
        .blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOS)
        .red = 250
        .green = 250
        .blue = 150
        .bold = 1
    End With
End Sub

Public Sub ServerHandleIncomingData(ByVal CharIndex As Integer)

On Error Resume Next
    Dim packetID As Byte
    
    packetID = charlist(CharIndex).incomingData.PeekByte()
    
    'Does the packet requires a logged user??
    If Not (packetID = ClientPacketID.LoginChar) Then
        
        'Is the user actually logged?
        If Not charlist(CharIndex).Logged Then
            Call CloseSocket(CharIndex)
            Exit Sub
        End If
    End If
    
    Select Case packetID
        Case ClientPacketID.LoginChar       'OLOGIN
            Call HandleLoginChar(CharIndex)

        Case ClientPacketID.Talk                    ';
            Call HandleTalk(CharIndex)

        Case ClientPacketID.Walk                    'M
            Call HandleWalk(CharIndex)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(CharIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(CharIndex)

        Case ClientPacketID.CastSpell               'LH
            Call HandleCastSpell(CharIndex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(CharIndex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(CharIndex)

        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(CharIndex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(CharIndex)

        Case ClientPacketID.Online                  '/ONLINE
            Call HandleOnline(CharIndex)
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(CharIndex)

        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(CharIndex)
            
        Case ClientPacketID.LanzaFlecha
            Call HandleLanzaFlecha(CharIndex)
            
        Case ClientPacketID.ChangePj
            Call HandleChangePj(CharIndex)
            
        Case Else
            'ERROR : Abort!
            Call CloseSocket(CharIndex)
    End Select
    
    'Done with this packet, move on to next one or send everything if no more packets found
    If charlist(CharIndex).incomingData.length > 0 And Err.Number = 0 Then
        Err.Clear
        Call ServerHandleIncomingData(CharIndex)
    
    ElseIf Err.Number <> 0 And Not Err.Number = charlist(CharIndex).incomingData.NotEnoughDataErrCode Then
        Call CloseSocket(CharIndex)
    Else
        'Flush buffer - send everything that has been written
        Call ServerFlushBuffer(CharIndex)
    End If
End Sub

Public Sub ClientHandleIncomingData()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error Resume Next

    Select Case incomingData.PeekByte()
        Case ServerPacketID.Logged                  ' LOGGED
            Call HandleLogged
        
        Case ServerPacketID.Disconnect              ' FINOK
            Call HandleDisconnect

        Case ServerPacketID.UpdateSta               ' ASS
            Call HandleUpdateSta
        
        Case ServerPacketID.UpdateMana              ' ASM
            Call HandleUpdateMana
        
        Case ServerPacketID.UpdateHP                ' ASH
            Call HandleUpdateHP

        Case ServerPacketID.ChangeMap               ' CM
            Call HandleChangeMap
        
        Case ServerPacketID.PosUpdate               ' PU
            Call HandlePosUpdate
        
        Case ServerPacketID.ChatOverHead            ' ||
            Call HandleChatOverHead
        
        Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
            Call HandleConsoleMessage

        Case ServerPacketID.CharIndexInServer       ' IP
            Call HandleCharIndexInServer
        
        Case ServerPacketID.CharacterCreate         ' CC
            Call HandleCharacterCreate
        
        Case ServerPacketID.CharacterRemove         ' BP
            Call HandleCharacterRemove
        
        Case ServerPacketID.CharacterChangeNick
            Call HandleCharacterChangeNick
            
        Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
            Call HandleCharacterMove
            
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        
        Case ServerPacketID.CharacterChange         ' CP
            Call HandleCharacterChange

        Case ServerPacketID.PlayMIDI                ' TM
            Call HandlePlayMIDI
        
        Case ServerPacketID.PlayWave                ' TW
            Call HandlePlayWave
        
        Case ServerPacketID.AreaChanged             ' CA
            Call HandleAreaChanged
        
        Case ServerPacketID.PauseToggle             ' BKW
            Call HandlePauseToggle

        Case ServerPacketID.CreateFX                ' CFX
            Call HandleCreateFX
        
        Case ServerPacketID.UpdateCharStats         ' EST
            Call HandleUpdateCharStats

        Case ServerPacketID.ErrorMsg                ' ERR
            Call HandleErrorMessage

        Case ServerPacketID.SetInvisible            ' NOVER
            Call HandleSetInvisible

        Case ServerPacketID.ParalizeOK              ' PARADOK
            Call HandleParalizeOK

        Case ServerPacketID.Pong
            Call HandlePong
        
        Case ServerPacketID.CuentaToggle
            Call HandleCuentaToggle
            
#If SeguridadAlkon Then
        Case Else
            Call HandleIncomingDataEx
#Else
        Case Else
            'ERROR : Abort!
            Exit Sub
#End If
    End Select
    
    'Done with this packet, move on to next one
    If incomingData.length > 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then
        Err.Clear
        Call ClientHandleIncomingData
    End If
End Sub

''
' Handles the Logged message.

Private Sub HandleLogged()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    ' Variable initialization
    EngineRun = True
    Nombres = True
    
    'Set connected state
    Call SetConnected
End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte

    frmMain.Socket1.Disconnect

    'Hide main form
    frmMain.Visible = False
    
    'Stop audio
    Call Audio.StopWave
    frmMain.IsPlaying = PlayLoop.plNone
    
    'Show connection form
    frmLogin.Visible = True
    
    'Reset global vars
    UserParalizado = False
    pausa = False

    'Delete all kind of dialogs
    Call CleanDialogs
    
    'Unload all forms except frmMain and frmConnect
    Dim frm As Form
    
    For Each frm In Forms
        If frm.name <> frmLogin.name And frm.name <> frmMain.name Then
            Unload frm
        End If
    Next
    
    For i = 1 To MAX_INVENTORY_SLOTS
        Call Inventario.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
    Next i

End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinSTA = incomingData.ReadInteger()
End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinMAN = incomingData.ReadInteger()
    
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    
    Dim bWidth As Byte
    
    If UserMaxMAN > 0 Then _
        bWidth = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 75)
        
    frmMain.shpMana.Width = 75 - bWidth
    frmMain.shpMana.Left = 584 + (75 - frmMain.shpMana.Width)
    
    frmMain.shpMana.Visible = (bWidth <> 75)
End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinHP = incomingData.ReadInteger()
    
    frmMain.lblVida = charlist(UserCharIndex).MinHP & "/" & charlist(UserCharIndex).MaxHP
    
    Dim bWidth As Byte
    
    bWidth = (((charlist(UserCharIndex).MinHP / 100) / (charlist(UserCharIndex).MaxHP / 100)) * 75)
    
    frmMain.shpVida.Width = 75 - bWidth
    frmMain.shpVida.Left = 584 + (75 - frmMain.shpVida.Width)
    
    frmMain.shpVida.Visible = (bWidth <> 75)
End Sub

''
' Handles the ChangeMap message.
Private Sub HandleChangeMap()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMap = incomingData.ReadInteger()
    
'TODO: Once on-the-fly editor is implemented check for map version before loading....
'For now we just drop it
    Call incomingData.ReadInteger
        
#If SeguridadAlkon Then
    Call InitMI
#End If
    
    If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
        Call SwitchMap(UserMap)
        If bLluvia(UserMap) = 0 Then
            If bRain Then
                Call Audio.StopWave(RainBufferIndex)
                RainBufferIndex = 0
                frmMain.IsPlaying = PlayLoop.plNone
            End If
        End If
    Else
        'no encontramos el mapa en el hd
        MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
        
        Call CloseClient
    End If
End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    If MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex Then
        MapData(UserPos.X, UserPos.Y).CharIndex = 0
    End If
    
    'Set new pos
    UserPos.X = incomingData.ReadByte()
    UserPos.Y = incomingData.ReadByte()
    charlist(UserCharIndex).Pos.X = UserPos.X
    charlist(UserCharIndex).Pos.Y = UserPos.Y

    'Set char
    MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex

End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim Chat As String
    Dim CharIndex As Integer
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    
    Chat = buffer.ReadASCIIString()
    CharIndex = buffer.ReadInteger()
    
    r = buffer.ReadByte()
    g = buffer.ReadByte()
    b = buffer.ReadByte()
    
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If charlist(CharIndex).Active Then _
        Call Dialogos.CreateDialog(Trim$(Chat), CharIndex, RGB(r, g, b))
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim Chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    
    Chat = buffer.ReadASCIIString()
    FontIndex = buffer.ReadByte()

    If InStr(1, Chat, "~") Then
        str = ReadField(2, Chat, 126)
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
            End If
            
            str = ReadField(3, Chat, 126)
            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)
            End If
            
            str = ReadField(4, Chat, 126)
            If Val(str) > 255 Then
                b = 255
            Else
                b = Val(str)
            End If
            
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(Chat, InStr(1, Chat, "~") - 1), r, g, b, Val(ReadField(5, Chat, 126)) <> 0, Val(ReadField(6, Chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, Chat, .red, .green, .blue, .bold, .italic)
        End With

    End If
'    Call checkText(chat)
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserIndex = incomingData.ReadInteger()
End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleCharIndexInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCharIndex = incomingData.ReadInteger()

End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 24 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim body As Integer
    Dim Head As Integer
    Dim heading As E_Heading
    Dim X As Byte
    Dim Y As Byte
    Dim weapon As Integer
    Dim shield As Integer
    Dim helmet As Integer
    Dim Nombre As String
    Dim MinHP As Integer
    Dim Team As Byte
    
    CharIndex = buffer.ReadInteger()
    body = buffer.ReadInteger()
    Head = buffer.ReadInteger()
    heading = buffer.ReadByte()
    X = buffer.ReadByte()
    Y = buffer.ReadByte()
    weapon = buffer.ReadInteger()
    shield = buffer.ReadInteger()
    helmet = buffer.ReadInteger()
    Nombre = buffer.ReadASCIIString()
    MinHP = buffer.ReadInteger()
    Team = buffer.ReadByte()
    
    Call MakeChar(CharIndex, body, Head, heading, X, Y, weapon, shield, helmet)
    charlist(CharIndex).MinHP = MinHP
    charlist(CharIndex).Nombre = Nombre
    charlist(CharIndex).Team = Team
    
    Call RefreshAllChars
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleCharacterChangeNick()
'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet id
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    CharIndex = incomingData.ReadInteger
    charlist(CharIndex).Nombre = incomingData.ReadASCIIString
    
End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    Call EraseChar(CharIndex)
    Call RefreshAllChars
End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim X As Byte
    Dim Y As Byte
    
    CharIndex = incomingData.ReadInteger()
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()

    Call DoPasosFx(CharIndex)

    Call MoveCharbyPos(CharIndex, X, Y)
    
    Call RefreshAllChars
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Direccion As Byte
    
    Direccion = incomingData.ReadByte()

    Call MoveCharbyHead(UserCharIndex, Direccion)
    Call MoveScreen(Direccion)
    
    Call RefreshAllChars
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2009
'25/08/2009: ZaMa - Changed a variable used incorrectly.
'***************************************************
    If incomingData.length < 18 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim TempInt As Integer
    Dim headIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    With charlist(CharIndex)
        TempInt = incomingData.ReadInteger()
        
        If TempInt < LBound(BodyData()) Or TempInt > UBound(BodyData()) Then
            .body = BodyData(0)
            .iBody = 0
        Else
            .body = BodyData(TempInt)
            .iBody = TempInt
        End If
        
        
        headIndex = incomingData.ReadInteger()
        
        If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .iHead = 0
        Else
            .Head = HeadData(headIndex)
            .iHead = headIndex
        End If

        .heading = incomingData.ReadByte()
        
        TempInt = incomingData.ReadInteger()
        If TempInt <> 0 Then
            .Arma = WeaponAnimData(TempInt)
            .iArma = TempInt
        End If
        
        TempInt = incomingData.ReadInteger()
        If TempInt <> 0 Then
            .Escudo = ShieldAnimData(TempInt)
            .iEscudo = TempInt
        End If
        
        TempInt = incomingData.ReadInteger()
        If TempInt <> 0 Then
            .Casco = CascoAnimData(TempInt)
            .iCasco = TempInt
        End If
        
        Call SetCharacterFx(CharIndex, incomingData.ReadInteger(), incomingData.ReadInteger())
    End With
    
    Call RefreshAllChars
End Sub

''
' Handles the PlayMIDI message.

Private Sub HandlePlayMIDI()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim currentMidi As Byte
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    currentMidi = incomingData.ReadByte()
    
    If currentMidi Then
        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", incomingData.ReadInteger())
    Else
        'Remove the bytes to prevent errors
        Call incomingData.ReadInteger
    End If
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Last Modified by: Rapsodius
'Added support for 3D Sounds.
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
    Dim wave As Byte
    Dim srcX As Byte
    Dim srcY As Byte
    
    wave = incomingData.ReadByte()
    srcX = incomingData.ReadByte()
    srcY = incomingData.ReadByte()
        
    Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
        
    Call CambioDeArea(X, Y)
End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    pausa = incomingData.ReadByte()
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim FX As Integer
    Dim Loops As Integer
    
    CharIndex = incomingData.ReadInteger()
    FX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call SetCharacterFx(CharIndex, FX, Loops)
End Sub


Private Sub HandleUpdateCharStats()

    If incomingData.length < 14 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim bWidth As Long
    Dim LastMP As Integer
    Dim LastHP As Integer
    Dim LastMaxHP As Integer
    Dim LastMaxMP As Integer
    Dim Dif As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    LastMP = charlist(UserCharIndex).MinMAN
    LastHP = charlist(UserCharIndex).MinHP
    LastMaxHP = charlist(UserCharIndex).MaxHP
    LastMaxMP = charlist(UserCharIndex).MaxMAN
    
    charlist(UserCharIndex).MaxHP = incomingData.ReadInteger()
    charlist(UserCharIndex).MinHP = incomingData.ReadInteger()
    charlist(UserCharIndex).MaxMAN = incomingData.ReadInteger()
    charlist(UserCharIndex).MinMAN = incomingData.ReadInteger()
    charlist(UserCharIndex).MaxSTA = incomingData.ReadInteger()
    charlist(UserCharIndex).MinSTA = incomingData.ReadInteger()
    charlist(UserCharIndex).Lvl = incomingData.ReadByte()
    
    frmMain.lblLvl = charlist(UserCharIndex).Lvl

    Dif = LastMP - charlist(UserCharIndex).MinMAN
    
    If LastMaxHP <> charlist(UserCharIndex).MaxHP Then
        frmMain.ActualHP = 0
    End If
    If LastMaxMP <> charlist(UserCharIndex).MaxMAN Then
        frmMain.ActualMP = 0
    End If
    
    If charlist(UserCharIndex).MaxMAN > 0 Then
        If frmMain.ActualMP < charlist(UserCharIndex).MinMAN Then
            frmMain.MoviendoMana = (charlist(UserCharIndex).MinMAN - frmMain.ActualMP + Dif) * 10 / 100
        ElseIf frmMain.ActualMP > charlist(UserCharIndex).MinMAN Then
            frmMain.MoviendoMana = (frmMain.ActualMP - charlist(UserCharIndex).MinMAN + Dif) * 10 / 100
        ElseIf frmMain.ActualMP = charlist(UserCharIndex).MinMAN Then
            frmMain.MoviendoMana = Dif * 10 / 100
        End If
    End If
    Dif = LastHP - charlist(UserCharIndex).MinHP
    If frmMain.ActualHP < charlist(UserCharIndex).MinHP Then
        frmMain.MoviendoVida = (charlist(UserCharIndex).MinHP - frmMain.ActualHP + Dif) * 10 / 100
    ElseIf frmMain.ActualHP > charlist(UserCharIndex).MinHP Then
        frmMain.MoviendoVida = (frmMain.ActualHP - charlist(UserCharIndex).MinHP + Dif) * 10 / 100
    ElseIf frmMain.ActualHP = charlist(UserCharIndex).MinHP Then
        frmMain.MoviendoVida = Dif * 10 / 100
    End If
    Call ActualizarStats

End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call MsgBox(buffer.ReadASCIIString())

    If frmLogin.Visible Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    charlist(CharIndex).invisible = incomingData.ReadBoolean()
    
#If SeguridadAlkon Then
    If charlist(CharIndex).invisible Then
        Call MI(CualMI).SetInvisible(CharIndex)
    Else
        Call MI(CualMI).ResetInvisible(CharIndex)
    End If
#End If

End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    charlist(UserCharIndex).Inmo = incomingData.ReadByte()

End Sub

''
' Handles the Pong message.

Private Sub HandlePong()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, "El ping es " & (GetTickCount - pingTime) & " ms.", 255, 0, 0, True, False, True)
    
    pingTime = 0
End Sub

Public Sub WriteLoginChar()

    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginChar)
        
        Call .WriteASCIIString(UserName)
        Call .WriteByte(UserClase)
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserGenero)
        Call .WriteByte(UserLvl)
        Call .WriteByte(UserTeam)
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)

    End With
End Sub

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal Chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Talk" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        
        Call .WriteASCIIString(Chat)
    End With
End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Walk" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Walk)
        
        Call .WriteByte(heading)
    End With
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestPositionUpdate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)
End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Attack" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Attack)
End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal X As Byte, ByVal Y As Byte, ByVal Spell As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CastSpell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteByte(Spell)
    End With
End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal Slot As Byte, ByVal Click As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        
        Call .WriteByte(Slot)
        Call .WriteByte(Click)
    End With
End Sub

''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkLeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .WriteByte(Skill)
    End With
End Sub

' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EquipItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        
        Call .WriteByte(Slot)
    End With
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeHeading" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(heading)
    End With
End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Online" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Online)
End Sub

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/16/08
'Writes the "Quit" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Quit)
End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/01/2007
'Writes the "Ping" message to the outgoing data buffer
'***************************************************
    'Prevent the timer from being cut
    If pingTime <> 0 Then Exit Sub
    
    Call outgoingData.WriteByte(ClientPacketID.Ping)
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
    DoEvents
    
    pingTime = GetTickCount
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    
    With outgoingData
        If .length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call SendData(sndData)
    End With
End Sub

Public Function PrepareMessageErrorMsg(ByVal Message As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ErrorMsg" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ErrorMsg)
        Call .WriteASCIIString(Message)
        
        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

Private Sub HandleLoginChar(ByVal CharIndex As Integer)

    If charlist(CharIndex).incomingData.length < 12 Then
        Err.Raise charlist(CharIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(charlist(CharIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim CharName As String
    Dim Clase As Byte
    Dim Raza As Byte
    Dim Genero As Byte
    Dim Lvl As Byte
    Dim Team As Byte
    
    Dim version As String
    
    CharName = buffer.ReadASCIIString()
    Clase = buffer.ReadByte()
    Raza = buffer.ReadByte()
    Genero = buffer.ReadByte()
    Lvl = buffer.ReadByte()
    Team = buffer.ReadByte()
    
    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    If Not AsciiValidos(CharName) Then
        Call WriteErrorMsg(CharIndex, "Nombre inválido.")
        Call ServerFlushBuffer(CharIndex)
        Call CloseSocket(CharIndex)
        
        Exit Sub
    End If
    
    If Clase < 0 Or Clase > 6 Then
        Call WriteErrorMsg(CharIndex, "Clase inválida.")
        Call ServerFlushBuffer(CharIndex)
        Call CloseSocket(CharIndex)
        
        Exit Sub
    End If

    If Genero < 0 Or Genero > 1 Then
        Call WriteErrorMsg(CharIndex, "Genero inválido.")
        Call ServerFlushBuffer(CharIndex)
        Call CloseSocket(CharIndex)
        
        Exit Sub
    End If
    
    If Raza < 0 Or Raza > 4 Then
        Call WriteErrorMsg(CharIndex, "Raza inválida.")
        Call ServerFlushBuffer(CharIndex)
        Call CloseSocket(CharIndex)
        
        Exit Sub
    End If
    
    If Lvl < 35 Or Lvl > 50 Then
        Call WriteErrorMsg(CharIndex, "Nivel inválido.")
        Call ServerFlushBuffer(CharIndex)
        Call CloseSocket(CharIndex)
        
        Exit Sub
    End If
    
    If Not VersionOK(version) Then
        Call WriteErrorMsg(CharIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". La misma se encuentra disponible en www.argentumonline.com.ar")
    Else
        Call ConnectChar(CharIndex, CharName, Clase, Raza, Genero, Lvl, Team)
    End If

    
    'If we got here then packet is complete, copy data back to original queue
    Call charlist(CharIndex).incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleTalk(ByVal CharIndex As Integer)

    If charlist(CharIndex).incomingData.length < 3 Then
        Err.Raise charlist(CharIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    With charlist(CharIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessageChatOverHead(Chat, CharIndex, RGB(255, 255, 255)))
            Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessageConsoleMsg(charlist(CharIndex).Nombre & "> " & Chat, FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleWalk(ByVal CharIndex As Integer)

    If charlist(CharIndex).incomingData.length < 2 Then
        Err.Raise charlist(CharIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim heading As E_Heading
    Dim nPos As WorldPos
    Dim X As Integer
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer
    
    With charlist(CharIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        heading = .incomingData.ReadByte()
        
        If .Inmo = 0 Then
            .Quieto = 0
            If CharIndex <> UserCharIndex Then
                Select Case heading
                    Case E_Heading.NORTH
                        If MoveToLegalPos(.Pos.X, .Pos.Y - 1) Then
                            nPos.X = .Pos.X
                            nPos.Y = .Pos.Y - 1
                        Else
                            Call WritePosUpdate(CharIndex)
                            Exit Sub
                        End If
                    Case E_Heading.EAST
                        If MoveToLegalPos(.Pos.X + 1, .Pos.Y) Then
                            nPos.X = .Pos.X + 1
                            nPos.Y = .Pos.Y
                        Else
                            Call WritePosUpdate(CharIndex)
                            Exit Sub
                        End If
                    Case E_Heading.SOUTH
                        If MoveToLegalPos(.Pos.X, .Pos.Y + 1) Then
                            nPos.X = .Pos.X
                            nPos.Y = .Pos.Y + 1
                        Else
                            Call WritePosUpdate(CharIndex)
                            Exit Sub
                        End If
                    Case E_Heading.WEST
                        If MoveToLegalPos(.Pos.X - 1, .Pos.Y) Then
                            nPos.X = .Pos.X - 1
                            nPos.Y = .Pos.Y
                        Else
                            Call WritePosUpdate(CharIndex)
                            Exit Sub
                        End If
                End Select
                X = .Pos.X
                Y = .Pos.Y
                
                addX = nPos.X - X
                addY = nPos.Y - Y
        
                MapData(.Pos.X, .Pos.Y).CharIndex = 0
                MapData(nPos.X, nPos.Y).CharIndex = CharIndex
                .Pos.X = nPos.X
                .Pos.Y = nPos.Y
                .heading = heading

                .MoveOffsetX = -1 * (TilePixelWidth * addX)
                .MoveOffsetY = -1 * (TilePixelHeight * addY)
                
                .Moving = 1
                
                .scrollDirectionX = Sgn(addX)
                .scrollDirectionY = Sgn(addY)
            Else
                nPos.X = .Pos.X
                nPos.Y = .Pos.Y
            End If
            Call ServerSendData(SendTarget.ToAllButIndexAndHost, CharIndex, PrepareMessageCharacterMove(CharIndex, nPos.X, nPos.Y))
        End If
        
    End With
End Sub

Private Sub HandleRequestPositionUpdate(ByVal CharIndex As Integer)

    'Remove packet ID
    charlist(CharIndex).incomingData.ReadByte
    
    Call WritePosUpdate(CharIndex)
End Sub

Private Sub HandleAttack(ByVal CharIndex As Integer)

    With charlist(CharIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If dead, can't attack
        If .MinHP = 0 Then
            Call WriteConsoleMsg(CharIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If equiped weapon is ranged, can't attack this way
        If .TipoArma = 2 Then
            Call WriteConsoleMsg(CharIndex, "No puedes usar así este arma.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Attack!
        Call UsuarioAtaca(CharIndex)
    End With
End Sub

Private Sub HandleCastSpell(ByVal CharIndex As Integer)

    If charlist(CharIndex).incomingData.length < 4 Then
        Err.Raise charlist(CharIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With charlist(CharIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Spell As Byte
        Dim X As Byte
        Dim Y As Byte
        
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        Spell = .incomingData.ReadByte()
        
        If .MinHP = 0 Then
            Call WriteConsoleMsg(CharIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Spell < LBound(Hechizo) Then
            Exit Sub
        ElseIf Spell > UBound(Hechizo) Then
            Exit Sub
        End If
        
        If Not CharTimer(CharIndex).Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
            If charlist(CharIndex).LastCombo = 1 And CharTimer(CharIndex).Check(TimersIndex.AttackCast, False) Then
                Call CharTimer(CharIndex).Restart(TimersIndex.AttackCast)
                Call LanzarSpell(CharIndex, Spell, X, Y)
                charlist(CharIndex).LastCombo = 2
            Else
                Call WriteConsoleMsg(CharIndex, "No puedes lanzar hechizos tan rápido.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            If CharTimer(CharIndex).Check(TimersIndex.CastSpell) Then
                Call LanzarSpell(CharIndex, Spell, X, Y)
                charlist(CharIndex).LastCombo = 2
            Else
                Call WriteConsoleMsg(CharIndex, "No puedes lanzar hechizos tan rápido.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
End Sub

Private Sub HandleLeftClick(ByVal CharIndex As Integer)

    If charlist(CharIndex).incomingData.length < 3 Then
        Err.Raise charlist(CharIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With charlist(CharIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        If X < 10 Or X > 90 Or Y < 10 Or Y > 90 Then Exit Sub
        
        Call VerTile(CharIndex, X, Y)
    End With
End Sub

Private Sub HandleUseItem(ByVal CharIndex As Integer)

    If charlist(CharIndex).incomingData.length < 3 Then
        Err.Raise charlist(CharIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With charlist(CharIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Click As Byte
        
        Slot = .incomingData.ReadByte()
        Click = .incomingData.ReadByte()
        
        If Slot < 1 Or Slot > 3 Then Exit Sub

        If CharTimer(CharIndex).Check(TimersIndex.GolpeU, False) Then
            If Click = 0 Then
                If CharTimer(CharIndex).Check(TimersIndex.UseItemWithU, False) Then
                    If CharTimer(CharIndex).Check(TimersIndex.UseItem, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItem)
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItemWithU)
                        Call UsarInvItem(CharIndex, Slot)
                    End If
                End If
            Else
                If CharTimer(CharIndex).Check(TimersIndex.UseItem, False) Then
                    Call CharTimer(CharIndex).Restart(TimersIndex.UseItem)
                    Call UsarInvItem(CharIndex, Slot)
                End If
            End If
        Else
            Call WriteConsoleMsg(CharIndex, "Debes esperar unos momentos para tomar una pocion.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleEquipItem(ByVal CharIndex As Integer)

    If charlist(CharIndex).incomingData.length < 2 Then
        Err.Raise charlist(CharIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With charlist(CharIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim itemSlot As Byte
        
        itemSlot = .incomingData.ReadByte()
        
        'Dead Chars can't equip items
        If .MinHP = 0 Then Exit Sub
        
        'Validate item slot
        If itemSlot > 5 Or itemSlot < 1 Then Exit Sub

        Call EquiparInvItem(CharIndex, itemSlot)
    End With
End Sub

Private Sub HandleChangeHeading(ByVal CharIndex As Integer)

    If charlist(CharIndex).incomingData.length < 2 Then
        Err.Raise charlist(CharIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With charlist(CharIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim heading As E_Heading
        Dim posX As Integer
        Dim posY As Integer
                
        heading = .incomingData.ReadByte()
        
        If .Inmo = 0 Then
            Select Case heading
                Case E_Heading.NORTH
                    posY = -1
                Case E_Heading.EAST
                    posX = 1
                Case E_Heading.SOUTH
                    posY = 1
                Case E_Heading.WEST
                    posX = -1
            End Select
            
                If LegalPos(.Pos.X + posX, .Pos.Y + posY) Then
                    Exit Sub
                End If
        End If
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If heading > 0 And heading < 5 Then
            .heading = heading
            Call ServerSendData(SendTarget.ToAllButIndex, CharIndex, PrepareMessageCharacterChange(.iBody, .iHead, heading, CharIndex, .iArma, .iEscudo, .FxIndex, .FX.Loops, .iCasco))
        End If
    End With
End Sub

Private Sub HandleOnline(ByVal CharIndex As Integer)

    Dim i As Long
    Dim Count As Long
    
    With charlist(CharIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        For i = 1 To MaxUsers
            If LenB(charlist(i).Nombre) <> 0 And .Bot = 0 Then
                Count = Count + 1
            End If
        Next i
        
        Call WriteConsoleMsg(CharIndex, "Número de usuarios: " & CStr(Count), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Private Sub HandleQuit(ByVal CharIndex As Integer)

    With charlist(CharIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call CloseSocket(CharIndex)
    End With
End Sub

Public Sub HandlePing(ByVal CharIndex As Integer)

    With charlist(CharIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Call WritePong(CharIndex)
    End With
End Sub

Public Sub ServerFlushBuffer(ByVal CharIndex As Integer)

    Dim sndData As String
    
    With charlist(CharIndex).outgoingData
        If .length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call EnviarDatosASlot(CharIndex, sndData)
    End With
End Sub

Public Sub WriteErrorMsg(ByVal CharIndex As Integer, ByVal Message As String)

On Error GoTo ErrHandler
    Call charlist(CharIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(Message))
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

Public Function PrepareMessageChatOverHead(ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(Chat)
        Call .WriteInteger(CharIndex)
        
        ' Write rgb channels and save one byte from long :D
        Call .WriteByte(color And &HFF)
        Call .WriteByte((color And &HFF00&) \ &H100&)
        Call .WriteByte((color And &HFF0000) \ &H10000)
        
        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub WriteConsoleMsg(ByVal CharIndex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames)

On Error GoTo ErrHandler
    Call charlist(CharIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(Chat, FontIndex))
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

Public Sub WritePong(ByVal CharIndex As Integer)

On Error GoTo ErrHandler
    Call charlist(CharIndex).outgoingData.WriteByte(ServerPacketID.Pong)
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

Public Function PrepareMessageConsoleMsg(ByVal Chat As String, ByVal FontIndex As FontTypeNames) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        
        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub WriteLoggedMessage(ByVal CharIndex As Integer)

On Error GoTo ErrHandler
    Call charlist(CharIndex).outgoingData.WriteByte(ServerPacketID.Logged)
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

Public Sub WriteCreateFX(ByVal CharIndex As Integer, ByVal TargetIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)

On Error GoTo ErrHandler
    Call charlist(CharIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(TargetIndex, FX, FXLoops))
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCreateFX = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As E_Heading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal helmet As Integer, ByVal name As String, ByVal MinHP As Integer, ByVal Team As Byte)

On Error GoTo ErrHandler
    Call charlist(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(body, Head, heading, CharIndex, X, Y, weapon, shield, _
                                                            helmet, name, MinHP, Team))
Exit Sub

ErrHandler:
    If Err.Number = charlist(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, ByVal Head As Integer, ByVal heading As E_Heading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal helmet As Integer, ByVal name As String, ByVal MinHP As Integer, ByVal Team As Byte) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteASCIIString(name)
        Call .WriteInteger(MinHP)
        Call .WriteByte(Team)
        
        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub WriteUpdateCharStats(ByVal CharIndex As Integer)

On Error GoTo ErrHandler
    With charlist(CharIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateCharStats)
        Call .WriteInteger(charlist(CharIndex).MaxHP)
        Call .WriteInteger(charlist(CharIndex).MinHP)
        Call .WriteInteger(charlist(CharIndex).MaxMAN)
        Call .WriteInteger(charlist(CharIndex).MinMAN)
        Call .WriteInteger(charlist(CharIndex).MaxSTA)
        Call .WriteInteger(charlist(CharIndex).MinSTA)
        Call .WriteByte(charlist(CharIndex).Lvl)
    End With
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

Public Sub WriteCharIndexInServer(ByVal CharIndex As Integer)

On Error GoTo ErrHandler
    With charlist(CharIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharIndexInServer)
        Call .WriteInteger(CharIndex)
    End With
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

Public Sub WritePosUpdate(ByVal CharIndex As Integer)

On Error GoTo ErrHandler
    With charlist(CharIndex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(charlist(CharIndex).Pos.X)
        Call .WriteByte(charlist(CharIndex).Pos.Y)
    End With
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub WritePlayWave(ByVal CharIndex As Integer, ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte)

On Error GoTo ErrHandler
    Call charlist(CharIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

Public Function PrepareMessagePlayWave(ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteByte(wave)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessagePlayWave = .ReadASCIIStringFixed(.length)
    End With
End Function

Private Sub HandleLanzaFlecha(ByVal CharIndex As Integer)

    If charlist(CharIndex).incomingData.length < 3 Then
        Err.Raise charlist(CharIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim X As Byte
    Dim Y As Byte
    
    With charlist(CharIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        'If dead, can't attack
        If .MinHP = 0 Then
            Call WriteConsoleMsg(CharIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If equiped weapon is melee, can't attack this way
        If .TipoArma = 1 Then
            Call WriteConsoleMsg(CharIndex, "No puedes usar así este arma.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not CharTimer(CharIndex).Check(TimersIndex.Attack, False) Then Exit Sub 'Check if arrows interval has finished.
        If Not CharTimer(CharIndex).Check(TimersIndex.Arrows) Then Exit Sub
            
        If MapData(X, Y).CharIndex > 0 Then
            Call UsuarioAtacaUsuario(CharIndex, MapData(X, Y).CharIndex)
        ElseIf MapData(X, Y + 1).CharIndex > 0 Then
            Call UsuarioAtacaUsuario(CharIndex, MapData(X, Y + 1).CharIndex)
        End If
    End With
End Sub

Public Sub WriteLanzaFlecha(ByVal X As Byte, ByVal Y As Byte)

    Call outgoingData.WriteByte(ClientPacketID.LanzaFlecha)
    
    Call outgoingData.WriteByte(X)
    Call outgoingData.WriteByte(Y)
End Sub

Public Sub WriteCharacterRemove(ByVal CharIndex As Integer, ByVal TargetIndex As Integer)

On Error GoTo ErrHandler
    Call charlist(CharIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(TargetIndex))
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub WritePauseToggle(ByVal CharIndex As Integer, ByVal Pausado As Byte)

On Error GoTo ErrHandler
    Call charlist(CharIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle(Pausado))
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

Public Function PrepareMessagePauseToggle(ByVal Pausado As Byte) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        Call .WriteByte(Pausado)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub WriteCuentaToggle(ByVal CharIndex As Integer, ByVal Contando As Byte)

On Error GoTo ErrHandler
    Call charlist(CharIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCuentaToggle(Contando))
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

Public Function PrepareMessageCuentaToggle(ByVal Contando As Byte) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CuentaToggle)
        Call .WriteByte(Contando)
        PrepareMessageCuentaToggle = .ReadASCIIStringFixed(.length)
    End With
End Function

Private Sub HandleCuentaToggle()
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    EnCuenta = incomingData.ReadByte()
End Sub

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As E_Heading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer)

On Error GoTo ErrHandler
    Call charlist(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(body, Head, heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
Exit Sub

ErrHandler:
    If Err.Number = charlist(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Function PrepareMessageCharacterChange(ByVal body As Integer, ByVal Head As Integer, ByVal heading As E_Heading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub WriteParalizeOK(ByVal CharIndex As Integer, ByVal Inmo As Byte)

On Error GoTo ErrHandler
    Call charlist(CharIndex).outgoingData.WriteByte(ServerPacketID.ParalizeOK)
    Call charlist(CharIndex).outgoingData.WriteByte(Inmo)
    'If CharIndex <> UserCharIndex Then Call WritePosUpdate(CharIndex)
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

Private Sub HandleChangePj(ByVal CharIndex As Integer)

    If charlist(CharIndex).incomingData.length < 6 Then
        Err.Raise charlist(CharIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim Clase As Byte
    Dim Raza As Byte
    Dim Genero As Byte
    Dim Nivel As Byte
    Dim Team As Byte
    
    With charlist(CharIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Clase = .incomingData.ReadByte()
        Raza = .incomingData.ReadByte()
        Genero = .incomingData.ReadByte()
        Nivel = .incomingData.ReadByte()
        Team = .incomingData.ReadByte()

        If Clase < 0 Or Clase > 6 Then Exit Sub
        If Raza < 0 Or Raza > 4 Then Exit Sub
        If Genero < 0 Or Genero > 1 Then Exit Sub
        If Nivel < 35 Or Nivel > 50 Then Exit Sub
        If Team < 1 Or Team > 2 Then Exit Sub
        
        If Team = 1 Then
            If UBound(Team1) = 5 Then
                Call WriteConsoleMsg(CharIndex, "El equipo está lleno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            If UBound(Team2) = 5 Then
                Call WriteConsoleMsg(CharIndex, "El equipo está lleno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        Call CambiarPj(CharIndex, Clase, Raza, Genero, Nivel, Team)
    End With
End Sub

Public Sub WriteChangePj()

    With outgoingData
        Call .WriteByte(ClientPacketID.ChangePj)
        
        Call .WriteByte(UserClase)
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserGenero)
        Call .WriteByte(UserLvl)
        Call .WriteByte(UserTeam)
    End With
End Sub

Public Sub WriteDisconnect(ByVal CharIndex As Integer)

On Error GoTo ErrHandler
    Call charlist(CharIndex).outgoingData.WriteByte(ServerPacketID.Disconnect)
Exit Sub

ErrHandler:
    If Err.Number = charlist(CharIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call ServerFlushBuffer(CharIndex)
        Resume
    End If
End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)
    
    'No enviamos nada si no estamos conectados
#If UsarWrench = 1 Then
    If Not frmMain.Socket1.IsWritable Then
        'Put data back in the bytequeue
        Call outgoingData.WriteASCIIStringFixed(sdData)
        
        Exit Sub
    End If
    
    If Not frmMain.Socket1.Connected Then Exit Sub
#Else
    If frmMain.Winsock1.State <> sckConnected Then Exit Sub
#End If

#If SeguridadAlkon Then
    Dim data() As Byte
    
    data = StrConv(sdData, vbFromUnicode)
    
    Call DataSent(data)
    
    sdData = StrConv(data, vbUnicode)
#End If
    
    'Send data!
#If UsarWrench = 1 Then
    Call frmMain.Socket1.Write(sdData, Len(sdData))
#Else
    Call frmMain.Winsock1.SendData(sdData)
#End If

End Sub
