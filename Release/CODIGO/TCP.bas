Attribute VB_Name = "TCP"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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

Public Const ULTIMAVERSION As String = "0.2.4"
Public Hosting As Byte

Sub Login()

    Call WriteLoginChar
    
    DoEvents
    
    Call FlushBuffer
End Sub

Sub ConnectChar(ByVal CharIndex As Integer, ByRef name As String, ByVal Clase As Byte, ByVal Raza As Byte, _
                ByVal Genero As Byte, ByVal Lvl As Byte, ByVal Team As Byte)

Dim N As Integer
Dim tStr As String
Dim i As Byte

With charlist(CharIndex)

    'Controlamos no pasar el maximo de usuarios
    If NumChars >= MaxUsers Then
        Call WriteErrorMsg(CharIndex, "El servidor ha alcanzado el máximo de usuarios soportado, por favor vuelva a intertarlo más tarde.")
        Call ServerFlushBuffer(CharIndex)
        Call CloseSocket(CharIndex)
        Exit Sub
    End If
    
    For i = 1 To MaxUsers
        If charlist(i).ConnID <> -1 Then
            If name = charlist(i).Nombre Then
                Call WriteErrorMsg(CharIndex, "Ya hay un usuario con ese nombre, por favor elija otro.")
                Call ServerFlushBuffer(CharIndex)
                Call CloseSocket(CharIndex)
                Exit Sub
            End If
        End If
    Next

    'Nombre de sistema
    .Nombre = name
    .Ai = Clase
    .Raza = Raza
    .Genero = Genero
    .Lvl = Lvl
    
    If Team = 1 Then
        If UserCharIndex = 0 Then
            Team1(UBound(Team1)) = CharIndex
            TeamData1(UBound(Team1)).Clase = Clase
            TeamData1(UBound(Team1)).Raza = Raza
            TeamData1(UBound(Team1)).Genero = Genero
            TeamData1(UBound(Team1)).Nivel = Lvl
            TeamData1(UBound(Team1)).index = CharIndex
            TeamData1(UBound(Team1)).Nombre = name
            TeamData2(UBound(Team1)).Bot = 0
        Else
            If UBound(Team1) > 0 Then
                ReDim Preserve Team1(1 To UBound(Team1) + 1) As Integer
            Else
                ReDim Team1(1 To UBound(Team1) + 1) As Integer
            End If
            Team1(UBound(Team1)) = CharIndex
            TeamData1(UBound(Team1)).Clase = Clase
            TeamData1(UBound(Team1)).Raza = Raza
            TeamData1(UBound(Team1)).Genero = Genero
            TeamData1(UBound(Team1)).Nivel = Lvl
            TeamData1(UBound(Team1)).index = CharIndex
            TeamData1(UBound(Team1)).Nombre = name
            TeamData2(UBound(Team1)).Bot = 0
        End If
    Else
        If UBound(Team2) > 0 Then
            ReDim Preserve Team2(1 To UBound(Team2) + 1) As Integer
        Else
            ReDim Team2(1 To UBound(Team2) + 1) As Integer
        End If
        Team2(UBound(Team2)) = CharIndex
        TeamData2(UBound(Team2)).Clase = Clase
        TeamData2(UBound(Team2)).Raza = Raza
        TeamData2(UBound(Team2)).Genero = Genero
        TeamData2(UBound(Team2)).Nivel = Lvl
        TeamData2(UBound(Team2)).index = CharIndex
        TeamData2(UBound(Team2)).Nombre = name
        TeamData2(UBound(Team2)).Bot = 0
    End If
    
    .Logged = True
    
    If UserCharIndex = 0 Then
        Call CrearChar(CharIndex)
        Call WriteCharIndexInServer(CharIndex)
        Call WriteUpdateCharStats(CharIndex)
    Else
        Call CrearChar(CharIndex)
        'For i = 1 To MaxUsers
        '    If charlist(i).ConnID <> -1 And charlist(i).Bot = 0 Then
                'If UserCharIndex <> i Then
        '            Call WriteCharacterCreate(CharIndex, charlist(i).iBody, charlist(i).iHead, charlist(i).heading, i, _
                            charlist(i).Pos.X, charlist(i).Pos.Y, charlist(i).iArma, charlist(i).iEscudo, charlist(i).iCasco, charlist(i).Nombre, charlist(i).MinHP, EnTeam(i))
                'End If
        '    End If
        'Next
        Call WriteCharIndexInServer(CharIndex)
        Call WriteUpdateCharStats(CharIndex)
        Call ResetDuelo(-1, UBound(Team1), UBound(Team2))
    End If
    
    Call IniTimers(CharIndex)
    Call WriteLoggedMessage(CharIndex)
    
    Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessageConsoleMsg("El usuario " & charlist(CharIndex).Nombre & " ha ingresado al equipo " & Team & ".", FontTypeNames.FONTTYPE_SERVER))
End With
End Sub

Function VersionOK(ByVal Ver As String) As Boolean
    VersionOK = (Ver = ULTIMAVERSION)
End Function

Function NextOpenUser() As Integer
    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers + 1
        If LoopC > MaxUsers Then Exit For
        If (charlist(LoopC).ConnID = -1 And charlist(LoopC).Bot = 0) Then Exit For
    Next LoopC
    
    NextOpenUser = LoopC
End Function

Function Numeric(ByVal cad As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function

Sub CloseSocket(ByVal CharIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler
    
    'Call SecurityIp.IpRestarConexion(GetLongIp(CharList(CharIndex).ip))
    
    If charlist(CharIndex).ConnID <> -1 Then
        If charlist(CharIndex).Logged > 0 Then
            Call WriteDisconnect(CharIndex)
            Call ServerFlushBuffer(CharIndex)
            Call CloseSocketSL(CharIndex)
            If EnTeam(CharIndex) = 1 Then
                If GetTeamIndex(CharIndex, 1) < 5 Then _
                    Call PushTeamData(1, GetTeamIndex(CharIndex, 1))
                Call ResetDuelo(-1, UBound(Team1) - 1, UBound(Team2))
            Else
                If GetTeamIndex(CharIndex, 2) < 5 Then _
                Call PushTeamData(2, GetTeamIndex(CharIndex, 2))
                Call ResetDuelo(-1, UBound(Team1), UBound(Team2) - 1)
            End If
            Call ServerSendData(SendTarget.ToAllButIndex, UserCharIndex, PrepareMessageCharacterRemove(CharIndex))
            Call EraseChar(CharIndex)
            NumChars = NumChars - 1
        Else
            Call CloseSocketSL(CharIndex)
        End If
    End If

    'Empty buffer for reuse
    Call charlist(CharIndex).incomingData.ReadASCIIStringFixed(charlist(CharIndex).incomingData.length)

    charlist(CharIndex).ConnID = -1
    charlist(CharIndex).ConnIDValida = False
Exit Sub

ErrHandler:
    charlist(CharIndex).ConnID = -1
    charlist(CharIndex).ConnIDValida = False
End Sub

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal CharIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

If charlist(CharIndex).ConnID <> -1 And charlist(CharIndex).ConnIDValida Then
    Call BorraSlotSock(charlist(CharIndex).ConnID)
    Call WSApiCloseSocket(charlist(CharIndex).ConnID)
    charlist(CharIndex).ConnID = -1
    charlist(CharIndex).ConnIDValida = False
End If

End Sub

''
' Send an string to a Slot
'
' @param userIndex The index of the User
' @param Datos The string that will be send
' @remarks If UsarQueSocket is 3 it won`t use the clsByteQueue

Public Function EnviarDatosASlot(ByVal UserIndex As Integer, ByRef Datos As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: 01/10/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Now it uses the clsByteQueue class and don`t make a FIFO Queue of String
'***************************************************

    Dim Ret As Long
    
    Ret = WsApiEnviar(UserIndex, Datos)
    
    If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
        ' Close the socket avoiding any critical error
        Call CloseSocketSL(UserIndex)
    End If
Exit Function

End Function

Sub CloseChar(ByVal CharIndex As Integer)

charlist(CharIndex).FX.GrhIndex = 0
charlist(CharIndex).FX.Loops = 0
Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessageCreateFX(CharIndex, 0, 0))


charlist(CharIndex).Logged = False

'Call EraseChar(CharIndex)

End Sub

