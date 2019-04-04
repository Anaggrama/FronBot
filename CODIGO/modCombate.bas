Attribute VB_Name = "modCombate"
'---------------------------------------------------------------------------------------
' Module    : modCombate
' Author    : Anagrama
' Date      : ???
' Purpose   : Adaptación del modulo de combate del servidor al cliente.
'---------------------------------------------------------------------------------------

Option Explicit

Public Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MinimoInt = b
    Else
        MinimoInt = a
    End If
End Function

Public Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MaximoInt = a
    Else
        MaximoInt = b
    End If
End Function

Private Function PoderEvasionEscudo(ByVal CharIndex As Integer) As Long
    PoderEvasionEscudo = (100 * charlist(CharIndex).ModEscudo) / 2
End Function

Private Function PoderEvasion(ByVal CharIndex As Integer) As Long
    Dim lTemp As Long
    With charlist(CharIndex)
        lTemp = (100 + _
          100 / 33 * .Agilidad) * .ModEvasion
       
        PoderEvasion = (lTemp + (2.5 * MaximoInt(40 - 12, 0)))
    End With
End Function

Private Function PoderAtaqueArma(ByVal CharIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With charlist(CharIndex)
        If .TipoArma = 1 Then
           PoderAtaqueTemp = (100 + 3 * .Agilidad) * .ModAtaqueArma
        Else
            PoderAtaqueTemp = (100 + 3 * .Agilidad) * .ModAtaqueProyectil
        End If
        
        PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt(40 - 12, 0)))
    End With
End Function

Public Function CalcularDaño(ByVal CharIndex As Integer) As Long
    Dim DañoArma As Long
    Dim DañoUsuario As Long
    Dim ModifClase As Single
    Dim DañoMaxArma As Long
    Dim DañoMinArma As Long
    
    With charlist(CharIndex)
        
        DañoArma = RandomNumber(.ArmaMinHit, .ArmaMaxHit)
        DañoMaxArma = .ArmaMaxHit

        DañoUsuario = RandomNumber(.MinHit, .MaxHit)
        
        If .TipoArma = 1 Then
            CalcularDaño = (3 * DañoArma + ((DañoMaxArma / 5) * MaximoInt(0, .Fuerza - 15)) + DañoUsuario) * .ModDañoArma
        Else
            CalcularDaño = (3 * DañoArma + ((DañoMaxArma / 5) * MaximoInt(0, .Fuerza - 15)) + DañoUsuario) * .ModDañoProyectil
        End If
    End With
    
End Function

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean

    Dim ProbRechazo As Long
    Dim Rechazo As Boolean
    Dim ProbExito As Long

    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaqueArma(AtacanteIndex) - (PoderEvasion(VictimaIndex) + PoderEvasionEscudo(VictimaIndex))) * 0.4))
    
    UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)

    If Not UsuarioImpacto Then
        If charlist(VictimaIndex).ModEscudo > 0 Then
            ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * 100 / (100 + 100)))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If Rechazo Then
                Call ServerSendData(SendTarget.ToAll, AtacanteIndex, PrepareMessagePlayWave(37, charlist(VictimaIndex).Pos.X, charlist(VictimaIndex).Pos.Y))
            Else
                Call ServerSendData(SendTarget.ToAll, AtacanteIndex, PrepareMessagePlayWave(2, charlist(AtacanteIndex).Pos.X, charlist(AtacanteIndex).Pos.Y))
            End If
        Else: Call ServerSendData(SendTarget.ToAll, AtacanteIndex, PrepareMessagePlayWave(2, charlist(AtacanteIndex).Pos.X, charlist(AtacanteIndex).Pos.Y))
        End If
    End If
    
    If charlist(VictimaIndex).Bot = 1 Then
        If UsuarioImpacto = True Then
            If (charlist(VictimaIndex).ComportamientoPotas = 2 Or charlist(VictimaIndex).MinHP < charlist(VictimaIndex).MaxHP / 2) Then
                charlist(VictimaIndex).ComportamientoPotas = 3
                If charlist(VictimaIndex).Lanzando = 1 Then
                    charlist(VictimaIndex).Lanzando = 0
                    If CharTimer(VictimaIndex).Check(TimersIndex.WaitP, False) Then
                        Call CharTimer(VictimaIndex).Restart(TimersIndex.WaitP)
                    End If
                End If
            End If
            If charlist(VictimaIndex).Ai = eBotAi.Paladin Or charlist(VictimaIndex).Ai = eBotAi.Asesino Or charlist(VictimaIndex).Ai = eBotAi.Clerigo Or charlist(VictimaIndex).Ai = eBotAi.Bardo Then
                If charlist(charlist(VictimaIndex).TargetIndex).Inmo = 1 And charlist(VictimaIndex).Inmo = 0 Then
                    charlist(VictimaIndex).ComportamientoPotas = 4
                   If charlist(VictimaIndex).Lanzando = 1 Then
                        charlist(VictimaIndex).Lanzando = 0
                        If CharTimer(VictimaIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(VictimaIndex).Restart(TimersIndex.WaitP)
                        End If
                    End If
                ElseIf charlist(charlist(VictimaIndex).TargetIndex).Inmo = 1 And charlist(VictimaIndex).Inmo = 1 Then
                    If charlist(VictimaIndex).ComportamientoPotas <> 4 And charlist(VictimaIndex).MinHP < charlist(VictimaIndex).MaxHP Then
                        charlist(VictimaIndex).ComportamientoPotas = 4
                        If charlist(VictimaIndex).Lanzando = 1 Then
                            charlist(VictimaIndex).Lanzando = 0
                            If CharTimer(VictimaIndex).Check(TimersIndex.WaitP, False) Then
                                Call CharTimer(VictimaIndex).Restart(TimersIndex.WaitP)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    If charlist(AtacanteIndex).Bot = 1 Then
        If charlist(AtacanteIndex).Ai = eBotAi.Paladin Then
            If charlist(AtacanteIndex).MinMAN >= 250 And UsuarioImpacto = True Then
                charlist(AtacanteIndex).ComportamientoCombo = IIf(RandomNumber(1, 3) <= 2, 2, 1)
            ElseIf charlist(AtacanteIndex).MinMAN < 250 And (charlist(VictimaIndex).Ai = eBotAi.Mago Or charlist(VictimaIndex).Ai = eBotAi.Druida) Then
                If RandomNumber(1, 2) = 1 Then
                    charlist(AtacanteIndex).ComportamientoPotas = 5
                    charlist(AtacanteIndex).ComportamientoCombo = 1
                Else
                    charlist(AtacanteIndex).ComportamientoCombo = IIf(RandomNumber(1, 3) <= 2, 2, 1)
                End If
            ElseIf charlist(AtacanteIndex).MinMAN < 250 Then
                charlist(AtacanteIndex).ComportamientoPotas = 5
                charlist(AtacanteIndex).ComportamientoCombo = 1
            End If
        ElseIf charlist(AtacanteIndex).Ai = eBotAi.Clerigo Or charlist(AtacanteIndex).Ai = eBotAi.Bardo Then
            If charlist(AtacanteIndex).MinMAN >= 460 And UsuarioImpacto = True Then
                charlist(AtacanteIndex).ComportamientoCombo = IIf(RandomNumber(1, 3) <= 2, 3, 1)
            ElseIf charlist(AtacanteIndex).MinMAN < 460 And (charlist(VictimaIndex).Ai = eBotAi.Mago Or charlist(VictimaIndex).Ai = eBotAi.Druida) Then
                If RandomNumber(1, 2) = 1 Then
                    charlist(AtacanteIndex).ComportamientoPotas = 5
                    charlist(AtacanteIndex).ComportamientoCombo = 1
                Else
                    charlist(AtacanteIndex).ComportamientoCombo = IIf(RandomNumber(1, 3) <= 2, 3, 1)
                End If
            ElseIf charlist(AtacanteIndex).MinMAN < 460 Then
                charlist(AtacanteIndex).ComportamientoPotas = 5
                charlist(AtacanteIndex).ComportamientoCombo = 1
            End If
        End If
    End If
    
End Function

Public Function UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
    
    With charlist(AtacanteIndex)
        If EnTeam(AtacanteIndex) = EnTeam(VictimaIndex) Then
            Call WriteConsoleMsg(AtacanteIndex, "No puedes atacar a tus compañeros.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If
        If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
            Call UserDañoUser(AtacanteIndex, VictimaIndex)
        Else
            If charlist(AtacanteIndex).Bot = 0 Then
                Call WriteConsoleMsg(AtacanteIndex, "¡¡Has fallado!!.", FontTypeNames.FONTTYPE_FIGHT)
            End If
            If charlist(VictimaIndex).Bot = 0 Then
                Call WriteConsoleMsg(VictimaIndex, "¡¡" & charlist(AtacanteIndex).Nombre & " ha fallado!!.", FontTypeNames.FONTTYPE_FIGHT)
            End If
        End If
        
    End With
    
    UsuarioAtacaUsuario = True
End Function

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    
    Dim daño As Long
    Dim Lugar As Byte
    Dim absorbido As Long
    Dim Suerte As Integer
            
    daño = CalcularDaño(AtacanteIndex)

    With charlist(AtacanteIndex)
        
        Lugar = RandomNum(6)
        
        If Lugar <> 6 Then
            absorbido = RandomNumber(charlist(VictimaIndex).MinDef, charlist(VictimaIndex).MaxDef)
            absorbido = absorbido - .Refuerzo
            daño = daño - absorbido
        Else
            absorbido = RandomNumber(charlist(VictimaIndex).MinDefH, charlist(VictimaIndex).MaxDefH)
            absorbido = absorbido - .Refuerzo
            daño = daño - absorbido
        End If
        
        If daño <= 0 Then daño = 1
        
        If .Ai = eBotAi.Asesino Or .Ai = eBotAi.Bardo Then
            If .Ai = eBotAi.Asesino Then
                Suerte = Int(((0.00003 * 100 - 0.002) * 100 + 0.098) * 100 + 4.25)
            ElseIf .Ai = eBotAi.Bardo Then
                Suerte = Int(((0.000002 * 100 + 0.0002) * 100 + 0.032) * 100 + 4.81)
            End If

            If RandomNum(101) - 1 < Suerte Then
                If .Ai = eBotAi.Asesino Then
                    daño = daño + Round(daño * 1.4, 0)
                    charlist(AtacanteIndex).ComportamientoCombo = 2
                ElseIf .Ai = eBotAi.Bardo Then
                    daño = daño + Round(daño * 1.5, 0)
                    charlist(AtacanteIndex).ComportamientoCombo = 2
                End If
                
                If .Bot = 0 Then
                    Call WriteConsoleMsg(AtacanteIndex, "¡¡Has apuñalado a " & charlist(VictimaIndex).Nombre & "!!.", FontTypeNames.FONTTYPE_FIGHT)
                End If
                If charlist(VictimaIndex).Bot = 0 Then
                    Call WriteConsoleMsg(VictimaIndex, "¡¡" & charlist(VictimaIndex).Nombre & " te ha apuñalado!!.", FontTypeNames.FONTTYPE_FIGHT)
                End If
            Else
                If .Bot = 0 Then _
                    Call WriteConsoleMsg(AtacanteIndex, "¡No has logrado apuñalar a tu enemigo!.", FontTypeNames.FONTTYPE_FIGHT)
            End If
        End If
        
        Call ServerSendData(SendTarget.ToAll, AtacanteIndex, PrepareMessagePlayWave(10, charlist(VictimaIndex).Pos.X, charlist(VictimaIndex).Pos.Y))
        
        charlist(VictimaIndex).MinHP = charlist(VictimaIndex).MinHP - daño
        If charlist(VictimaIndex).MinHP < 0 Then charlist(VictimaIndex).MinHP = 0
        
        If .Bot = 0 Then
            Call WriteConsoleMsg(AtacanteIndex, "Le has pegado a " & charlist(VictimaIndex).Nombre & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteUpdateCharStats(AtacanteIndex)
            If charlist(VictimaIndex).Bot = 0 Then
                Call WriteConsoleMsg(VictimaIndex, charlist(AtacanteIndex).Nombre & " Te ha pegado por " & daño, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteUpdateCharStats(VictimaIndex)
                If charlist(VictimaIndex).MinHP <= 0 Then
                    Call WriteConsoleMsg(VictimaIndex, charlist(AtacanteIndex).Nombre & " te ha matado.", FontTypeNames.FONTTYPE_FIGHT)
                End If
            End If
            If charlist(VictimaIndex).MinHP <= 0 Then
                Call WriteConsoleMsg(AtacanteIndex, "Has matado a " & charlist(VictimaIndex).Nombre, FontTypeNames.FONTTYPE_FIGHT)
                Call ServerSendData(SendTarget.ToAllButIndex, AtacanteIndex, PrepareMessageConsoleMsg(charlist(AtacanteIndex).Nombre & " ha matado a " & charlist(VictimaIndex).Nombre & ".", FontTypeNames.FONTTYPE_FIGHT))
                
                Call CharDie(VictimaIndex)
                Call ResetDuelo(AtacanteIndex)
                Call GetTargetBotTeam(EnTeam(AtacanteIndex))
            End If
        ElseIf charlist(VictimaIndex).Bot = 0 Then
            Call WriteConsoleMsg(VictimaIndex, charlist(AtacanteIndex).Nombre & " Te ha pegado por " & daño, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteUpdateCharStats(VictimaIndex)
            If charlist(VictimaIndex).MinHP <= 0 Then
                Call WriteConsoleMsg(VictimaIndex, charlist(AtacanteIndex).Nombre & " te ha matado.", FontTypeNames.FONTTYPE_FIGHT)
                Call ServerSendData(SendTarget.ToAllButIndex, VictimaIndex, PrepareMessageConsoleMsg(charlist(AtacanteIndex).Nombre & " ha matado a " & charlist(VictimaIndex).Nombre & ".", FontTypeNames.FONTTYPE_FIGHT))
                Call CharDie(VictimaIndex)
                Call ResetDuelo(AtacanteIndex)
                Call GetTargetBotTeam(EnTeam(AtacanteIndex))
            End If
        Else
            If charlist(VictimaIndex).MinHP <= 0 Then
                Call ServerSendData(SendTarget.ToAllButIndex, VictimaIndex, PrepareMessageConsoleMsg(charlist(AtacanteIndex).Nombre & " ha matado a " & charlist(VictimaIndex).Nombre & ".", FontTypeNames.FONTTYPE_FIGHT))
                Call CharDie(VictimaIndex)
                Call ResetDuelo(AtacanteIndex)
                Call GetTargetBotTeam(EnTeam(AtacanteIndex))
            End If
        End If
        
    End With
    
End Sub

Public Sub UsuarioAtaca(ByVal CharIndex As Integer)

    Dim index As Integer
    Dim AttackPos As WorldPos
    
    If Not CharTimer(CharIndex).Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
    If Not CharTimer(CharIndex).Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
        If charlist(CharIndex).LastCombo = 2 Then
            If Not CharTimer(CharIndex).Check(TimersIndex.CastAttack) Then Exit Sub
        Else
            If Not CharTimer(CharIndex).Check(TimersIndex.Attack) Then Exit Sub
        End If
    Else
        If Not CharTimer(CharIndex).Check(TimersIndex.Attack) Then Exit Sub
    End If
    
    With charlist(CharIndex)
        'Quitamos stamina
        If .MinSTA >= 10 Then
            .MinSTA = .MinSTA - RandomNum(10)
        Else
            If .Genero = 0 Then
                Call WriteConsoleMsg(CharIndex, "Estás muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(CharIndex, "Estás muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If

        Select Case charlist(CharIndex).heading
            Case 1
                If MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y - 1).CharIndex > 0 Then
                    Call UsuarioAtacaUsuario(CharIndex, MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y - 1).CharIndex)
                Else
                    Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(2, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
                End If
            Case 2
                If MapData(charlist(CharIndex).Pos.X + 1, charlist(CharIndex).Pos.Y).CharIndex > 0 Then
                    Call UsuarioAtacaUsuario(CharIndex, MapData(charlist(CharIndex).Pos.X + 1, charlist(CharIndex).Pos.Y).CharIndex)
                Else
                    Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(2, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
                End If
            Case 3
                If MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y + 1).CharIndex > 0 Then
                    Call UsuarioAtacaUsuario(CharIndex, MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y + 1).CharIndex)
                Else
                    Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(2, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
                End If
            Case 4
                If MapData(charlist(CharIndex).Pos.X - 1, charlist(CharIndex).Pos.Y).CharIndex > 0 Then
                    Call UsuarioAtacaUsuario(CharIndex, MapData(charlist(CharIndex).Pos.X - 1, charlist(CharIndex).Pos.Y).CharIndex)
                Else
                    Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(2, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
                End If
        End Select
        Call CharTimer(CharIndex).Restart(TimersIndex.GolpeU)
        charlist(CharIndex).LastCombo = 1

    End With
End Sub

