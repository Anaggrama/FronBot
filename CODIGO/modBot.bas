Attribute VB_Name = "modBot"
'---------------------------------------------------------------------------------------
' Module    : modBot
' Author    : Anagrama
' Date      : ???
' Purpose   : Contiene casi todas las funciones que se utilizan en la inteligencia artificial de los bots.
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : BotMoveTo
' Author    : Anagrama
' Date      : ???
' Purpose   : Mueve al bot, busca una dirección posible.
'---------------------------------------------------------------------------------------
'
Public Sub BotMoveTo(ByVal CharIndex As Integer, ByVal Direccion As E_Heading, Optional ByVal Atacando As Byte)
    Dim LegalOk As Boolean
    Dim i As Byte
    
    With charlist(CharIndex)
        
        Select Case Direccion
            Case E_Heading.NORTH
                LegalOk = MoveToLegalPos(.Pos.X, .Pos.Y - 1)
            Case E_Heading.EAST
                LegalOk = MoveToLegalPos(.Pos.X + 1, .Pos.Y)
            Case E_Heading.SOUTH
                LegalOk = MoveToLegalPos(.Pos.X, .Pos.Y + 1)
            Case E_Heading.WEST
                LegalOk = MoveToLegalPos(.Pos.X - 1, .Pos.Y)
        End Select
        
        If Atacando = 0 Then
            If LegalOk = False Then
                If Direccion = E_Heading.NORTH Then
                    If MoveToLegalPos(.Pos.X + 1, .Pos.Y) Then
                        Direccion = 2
                        LegalOk = True
                    ElseIf MoveToLegalPos(.Pos.X, .Pos.Y + 1) Then
                        Direccion = 3
                        LegalOk = True
                    ElseIf MoveToLegalPos(.Pos.X - 1, .Pos.Y) Then
                        Direccion = 4
                        LegalOk = True
                    End If
                 ElseIf Direccion = E_Heading.EAST Then
                    If MoveToLegalPos(.Pos.X, .Pos.Y - 1) Then
                        Direccion = 1
                        LegalOk = True
                    ElseIf MoveToLegalPos(.Pos.X, .Pos.Y + 1) Then
                        Direccion = 3
                        LegalOk = True
                    ElseIf MoveToLegalPos(.Pos.X - 1, .Pos.Y) Then
                        Direccion = 4
                        LegalOk = True
                    End If
                  ElseIf Direccion = E_Heading.SOUTH Then
                    If MoveToLegalPos(.Pos.X + 1, .Pos.Y) Then
                        Direccion = 2
                        LegalOk = True
                    ElseIf MoveToLegalPos(.Pos.X, .Pos.Y - 1) Then
                        Direccion = 1
                        LegalOk = True
                    ElseIf MoveToLegalPos(.Pos.X - 1, .Pos.Y) Then
                        Direccion = 4
                        LegalOk = True
                    End If
                 ElseIf Direccion = E_Heading.WEST Then
                    If MoveToLegalPos(.Pos.X + 1, .Pos.Y) Then
                        Direccion = 2
                        LegalOk = True
                    ElseIf MoveToLegalPos(.Pos.X, .Pos.Y - 1) Then
                        Direccion = 1
                        LegalOk = True
                    ElseIf MoveToLegalPos(.Pos.X, .Pos.Y + 1) Then
                        Direccion = 3
                        LegalOk = True
                    End If
                 End If
            End If
        Else
            If LegalOk = False Then
                If Direccion = E_Heading.NORTH Then
                    If MoveToLegalPos(.Pos.X + 1, .Pos.Y) Then
                        Direccion = 2
                        LegalOk = True
                    ElseIf MoveToLegalPos(.Pos.X - 1, .Pos.Y) Then
                        Direccion = 4
                        LegalOk = True
                    End If
                 ElseIf Direccion = E_Heading.EAST Then
                    If MoveToLegalPos(.Pos.X, .Pos.Y - 1) Then
                        Direccion = 1
                        LegalOk = True
                    ElseIf MoveToLegalPos(.Pos.X, .Pos.Y + 1) Then
                        Direccion = 3
                        LegalOk = True
                    End If
                  ElseIf Direccion = E_Heading.SOUTH Then
                    If MoveToLegalPos(.Pos.X + 1, .Pos.Y) Then
                        Direccion = 2
                        LegalOk = True
                    ElseIf MoveToLegalPos(.Pos.X - 1, .Pos.Y) Then
                        Direccion = 4
                        LegalOk = True
                    End If
                 ElseIf Direccion = E_Heading.WEST Then
                    If MoveToLegalPos(.Pos.X, .Pos.Y + 1) Then
                        Direccion = 3
                        LegalOk = True
                    ElseIf MoveToLegalPos(.Pos.X, .Pos.Y - 1) Then
                        Direccion = 1
                        LegalOk = True
                    End If
                 End If
            End If
        End If
        If LegalOk Then
            MoveCharbyHead CharIndex, Direccion
        End If

    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : MoveBot
' Author    : Anagrama
' Date      : ???
' Purpose   : Busca una dirección para moverse, revisa si puede moverse y se usa para manejar los ataques cuerpo a cuerpo.
'---------------------------------------------------------------------------------------
'
Public Sub MoveBot(ByVal CharIndex As Integer)
    Dim i As Byte
    Dim deadTarget As Integer
    
    If charlist(CharIndex).MinHP = 0 Then
        If EnTeam(CharIndex) = 1 Then
            For i = 1 To UBound(Team1)
                If charlist(Team1(i)).MinHP > 0 And charlist(Team1(i)).Ai <> eBotAi.Guerrero Then
                    deadTarget = Team1(i)
                    Exit For
                End If
            Next
        Else
            For i = 1 To UBound(Team2)
                If charlist(Team2(i)).MinHP > 0 And charlist(Team2(i)).Ai <> eBotAi.Guerrero Then
                    deadTarget = Team2(i)
                    Exit For
                End If
            Next
        End If
        If deadTarget = 0 Then Exit Sub
        If Distancia(CharIndex, deadTarget) > 8 Then
            Select Case RandomNum(4)
                Case 1
                    If charlist(deadTarget).Pos.X > charlist(CharIndex).Pos.X Then
                        Call BotMoveTo(CharIndex, 2, 1)
                    ElseIf charlist(deadTarget).Pos.X < charlist(CharIndex).Pos.X Then
                        Call BotMoveTo(CharIndex, 4, 1)
                    ElseIf charlist(deadTarget).Pos.Y > charlist(CharIndex).Pos.Y Then
                        Call BotMoveTo(CharIndex, 3, 1)
                    ElseIf charlist(deadTarget).Pos.Y < charlist(CharIndex).Pos.Y Then
                        Call BotMoveTo(CharIndex, 1, 1)
                    End If
                Case 2
                    If charlist(deadTarget).Pos.Y > charlist(CharIndex).Pos.Y Then
                        Call BotMoveTo(CharIndex, 3, 1)
                    ElseIf charlist(deadTarget).Pos.Y < charlist(CharIndex).Pos.Y Then
                        Call BotMoveTo(CharIndex, 1, 1)
                    ElseIf charlist(deadTarget).Pos.X > charlist(CharIndex).Pos.X Then
                        Call BotMoveTo(CharIndex, 2, 1)
                    ElseIf charlist(deadTarget).Pos.X < charlist(CharIndex).Pos.X Then
                        Call BotMoveTo(CharIndex, 4, 1)
                    End If
                Case 3
                    If charlist(deadTarget).Pos.X > charlist(CharIndex).Pos.X Then
                        Call BotMoveTo(CharIndex, 2, 1)
                    ElseIf charlist(deadTarget).Pos.Y > charlist(CharIndex).Pos.Y Then
                        Call BotMoveTo(CharIndex, 3, 1)
                    ElseIf charlist(deadTarget).Pos.Y < charlist(CharIndex).Pos.Y Then
                        Call BotMoveTo(CharIndex, 1, 1)
                    ElseIf charlist(deadTarget).Pos.X < charlist(CharIndex).Pos.X Then
                        Call BotMoveTo(CharIndex, 4, 1)
                    End If
                Case 4
                    If charlist(deadTarget).Pos.Y > charlist(CharIndex).Pos.Y Then
                        Call BotMoveTo(CharIndex, 3, 1)
                    ElseIf charlist(deadTarget).Pos.X > charlist(CharIndex).Pos.X Then
                        Call BotMoveTo(CharIndex, 2, 1)
                    ElseIf charlist(deadTarget).Pos.X < charlist(CharIndex).Pos.X Then
                        Call BotMoveTo(CharIndex, 4, 1)
                    ElseIf charlist(deadTarget).Pos.Y < charlist(CharIndex).Pos.Y Then
                        Call BotMoveTo(CharIndex, 1, 1)
                    End If
            End Select
        End If
        Exit Sub
    End If
    
    If charlist(CharIndex).TargetIndex = 0 Then Exit Sub
    
    If EnArea(CharIndex, charlist(CharIndex).TargetIndex, True) = 0 Then
        Select Case RandomNum(4)
            Case 1
                If charlist(charlist(CharIndex).TargetIndex).Pos.X > charlist(CharIndex).Pos.X Then
                    Call BotMoveTo(CharIndex, 2)
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.X < charlist(CharIndex).Pos.X Then
                    Call BotMoveTo(CharIndex, 4)
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y > charlist(CharIndex).Pos.Y Then
                    Call BotMoveTo(CharIndex, 3)
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y < charlist(CharIndex).Pos.Y Then
                    Call BotMoveTo(CharIndex, 1)
                End If
            Case 2
                If charlist(charlist(CharIndex).TargetIndex).Pos.Y > charlist(CharIndex).Pos.Y Then
                    Call BotMoveTo(CharIndex, 3)
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y < charlist(CharIndex).Pos.Y Then
                    Call BotMoveTo(CharIndex, 1)
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.X > charlist(CharIndex).Pos.X Then
                    Call BotMoveTo(CharIndex, 2)
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.X < charlist(CharIndex).Pos.X Then
                    Call BotMoveTo(CharIndex, 4)
                End If
            Case 3
                If charlist(charlist(CharIndex).TargetIndex).Pos.X > charlist(CharIndex).Pos.X Then
                    Call BotMoveTo(CharIndex, 2)
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y > charlist(CharIndex).Pos.Y Then
                    Call BotMoveTo(CharIndex, 3)
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y < charlist(CharIndex).Pos.Y Then
                    Call BotMoveTo(CharIndex, 1)
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.X < charlist(CharIndex).Pos.X Then
                    Call BotMoveTo(CharIndex, 4)
                End If
            Case 4
                If charlist(charlist(CharIndex).TargetIndex).Pos.Y > charlist(CharIndex).Pos.Y Then
                    Call BotMoveTo(CharIndex, 3)
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.X > charlist(CharIndex).Pos.X Then
                    Call BotMoveTo(CharIndex, 2)
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.X < charlist(CharIndex).Pos.X Then
                    Call BotMoveTo(CharIndex, 4)
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y < charlist(CharIndex).Pos.Y Then
                    Call BotMoveTo(CharIndex, 1)
                End If
        End Select
        Exit Sub
    End If
    
    If charlist(CharIndex).Ai = eBotAi.Mago Or charlist(CharIndex).Ai = eBotAi.Druida Then
        Call BotMoveTo(CharIndex, RandomNum(4)) 'No tiran inmo, siempre es al azar el movimiento.
    Else
        If charlist(charlist(CharIndex).TargetIndex).Inmo And Not (charlist(CharIndex).Ai = eBotAi.Bardo And (charlist(CharIndex).ComportamientoCombo = 2 Or charlist(CharIndex).ComportamientoCombo = 3)) Then
            If charlist(CharIndex).ComportamientoPotas = 4 Or charlist(CharIndex).ComportamientoPotas = 5 Then
                Call BotMoveTo(CharIndex, RandomNum(4))
                Exit Sub
            End If
            If charlist(CharIndex).Ai = eBotAi.Guerrero Then
                If Distancia(CharIndex, charlist(CharIndex).TargetIndex) <= 5 Then
                    If charlist(CharIndex).TipoArma = 2 Then
                        Call BotEquiparItem(CharIndex, 3)
                    End If
                End If
            End If
            If Distancia(CharIndex, charlist(CharIndex).TargetIndex) > 1 Then 'Tiraron inmo y son clases de golpe, asi que van hacia el usuario para atacar en un movimiento "al azar" pero manteniendo siempre la dirección que los acerque.
                Select Case RandomNum(4)
                    Case 1
                        If charlist(charlist(CharIndex).TargetIndex).Pos.X > charlist(CharIndex).Pos.X Then
                            Call BotMoveTo(CharIndex, 2, 1)
                        ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.X < charlist(CharIndex).Pos.X Then
                            Call BotMoveTo(CharIndex, 4, 1)
                        ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y > charlist(CharIndex).Pos.Y Then
                            Call BotMoveTo(CharIndex, 3, 1)
                        ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y < charlist(CharIndex).Pos.Y Then
                            Call BotMoveTo(CharIndex, 1, 1)
                        End If
                    Case 2
                        If charlist(charlist(CharIndex).TargetIndex).Pos.Y > charlist(CharIndex).Pos.Y Then
                            Call BotMoveTo(CharIndex, 3, 1)
                        ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y < charlist(CharIndex).Pos.Y Then
                            Call BotMoveTo(CharIndex, 1, 1)
                        ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.X > charlist(CharIndex).Pos.X Then
                            Call BotMoveTo(CharIndex, 2, 1)
                        ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.X < charlist(CharIndex).Pos.X Then
                            Call BotMoveTo(CharIndex, 4, 1)
                        End If
                    Case 3
                        If charlist(charlist(CharIndex).TargetIndex).Pos.X > charlist(CharIndex).Pos.X Then
                            Call BotMoveTo(CharIndex, 2, 1)
                        ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y > charlist(CharIndex).Pos.Y Then
                            Call BotMoveTo(CharIndex, 3, 1)
                        ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y < charlist(CharIndex).Pos.Y Then
                            Call BotMoveTo(CharIndex, 1, 1)
                        ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.X < charlist(CharIndex).Pos.X Then
                            Call BotMoveTo(CharIndex, 4, 1)
                        End If
                    Case 4
                        If charlist(charlist(CharIndex).TargetIndex).Pos.Y > charlist(CharIndex).Pos.Y Then
                            Call BotMoveTo(CharIndex, 3, 1)
                        ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.X > charlist(CharIndex).Pos.X Then
                            Call BotMoveTo(CharIndex, 2, 1)
                        ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.X < charlist(CharIndex).Pos.X Then
                            Call BotMoveTo(CharIndex, 4, 1)
                        ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y < charlist(CharIndex).Pos.Y Then
                            Call BotMoveTo(CharIndex, 1, 1)
                        End If
                End Select
                Exit Sub
            Else 'Tire inmo, estoy al lado, miro hacia el usuario e intento atacar.
                If charlist(charlist(CharIndex).TargetIndex).Pos.X > charlist(CharIndex).Pos.X Then
                    charlist(CharIndex).heading = 2
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.X < charlist(CharIndex).Pos.X Then
                    charlist(CharIndex).heading = 4
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y > charlist(CharIndex).Pos.Y Then
                    charlist(CharIndex).heading = 3
                ElseIf charlist(charlist(CharIndex).TargetIndex).Pos.Y < charlist(CharIndex).Pos.Y Then
                    charlist(CharIndex).heading = 1
                End If
                
                If charlist(CharIndex).ComportamientoCombo = 2 Or charlist(CharIndex).ComportamientoPotas = 4 Or charlist(CharIndex).ComportamientoPotas = 5 Then Exit Sub
                
                If Not CharTimer(CharIndex).Check(TimersIndex.Arrows, False) Then Exit Sub
                If Not CharTimer(CharIndex).Check(TimersIndex.CastSpell, False) Then
                    If charlist(CharIndex).LastCombo = 2 Then
                        If Not CharTimer(CharIndex).Check(TimersIndex.CastAttack) Then Exit Sub
                    Else
                        If Not CharTimer(CharIndex).Check(TimersIndex.Attack) Then Exit Sub
                    End If
                Else
                    If Not CharTimer(CharIndex).Check(TimersIndex.Attack) Then Exit Sub
                End If
                    Call CharTimer(CharIndex).Restart(TimersIndex.GolpeU)
                    Call UsuarioAtacaUsuario(CharIndex, charlist(CharIndex).TargetIndex)
                    charlist(CharIndex).LastCombo = 1
                Exit Sub
            End If
        Else
            If charlist(CharIndex).Ai = eBotAi.Guerrero And charlist(CharIndex).TipoArma = 1 Then
                Call BotEquiparItem(CharIndex, 4)
            End If
            Call BotMoveTo(CharIndex, RandomNum(4))
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Distancia
' Author    : Anagrama
' Date      : ???
' Purpose   : Distancia...
'---------------------------------------------------------------------------------------
'
Public Function Distancia(ByVal CharIndex As Integer, ByVal TargetIndex As Integer) As Byte
    Distancia = Abs(charlist(TargetIndex).Pos.X - charlist(CharIndex).Pos.X)
    Distancia = Distancia + Abs(charlist(TargetIndex).Pos.Y - charlist(CharIndex).Pos.Y)
End Function

'---------------------------------------------------------------------------------------
' Procedure : EnArea
' Author    : Anagrama
' Date      : ???
' Purpose   : Revisa si el usuario está en el area de visión, cuando se mueve se agrega 1 tile más
' porque si no se mueve sin atacar justo en el borde de la pantalla.
'---------------------------------------------------------------------------------------
'
Public Function EnArea(ByVal CharIndex As Integer, ByVal TargetIndex As Integer, Optional Moviendo As Boolean = False) As Byte
    If Not Moviendo Then
        If Abs(charlist(TargetIndex).Pos.X - charlist(CharIndex).Pos.X) < 9 And Abs(charlist(TargetIndex).Pos.Y - charlist(CharIndex).Pos.Y) < 7 Then EnArea = 1
    Else
        If Abs(charlist(TargetIndex).Pos.X - charlist(CharIndex).Pos.X) < 8 And Abs(charlist(TargetIndex).Pos.Y - charlist(CharIndex).Pos.Y) < 6 Then EnArea = 1
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : BotLanzaSpell
' Author    : Anagrama
' Date      : ???
' Purpose   : Reforme todo esto para darle más coherencia, ahora cada clase se maneja en su propio sub.
' El lado malo de esto es que me pueden haber quedado algunas cosas mal por el pasaje.
'---------------------------------------------------------------------------------------
'
Public Sub BotLanzaSpell(ByVal CharIndex As Integer)
    
    If charlist(CharIndex).TargetIndex = 0 Then Exit Sub
    If charlist(CharIndex).Ai = eBotAi.Mago Then
        Call MagoLanzaSpell(CharIndex)
    ElseIf charlist(CharIndex).Ai = eBotAi.Paladin Or charlist(CharIndex).Ai = eBotAi.Asesino Then
        Call PaladinLanzaSpell(CharIndex) 'El pala y el ase usan la misma lógica.
    ElseIf charlist(CharIndex).Ai = eBotAi.Clerigo Then
        Call ClerigoLanzaSpell(CharIndex)
    ElseIf charlist(CharIndex).Ai = eBotAi.Bardo Then
        Call BardoLanzaSpell(CharIndex)
    ElseIf charlist(CharIndex).Ai = eBotAi.Druida Then
        Call DruidaLanzaSpell(CharIndex)
    End If
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : MagoLanzaSpell
' Author    : Anagrama
' Date      : ???
' Purpose   : Selección de hechizo a lanzar según la situación para el mago.
'---------------------------------------------------------------------------------------
'
Public Sub MagoLanzaSpell(ByVal CharIndex As Integer)
    Dim daño As Integer
    Dim h As Byte
    Dim i As Byte
    
    If charlist(CharIndex).Inmo = 1 Then
        If CharTimer(CharIndex).Check(TimersIndex.Remo, False) Then
            Call BotCasteoSpell(CharIndex, CharIndex, 4)
        End If
        Exit Sub
    End If
    
    charlist(CharIndex).Lanzando = 1
    
    If RandomNum(Dificultad) <> 1 And charlist(charlist(CharIndex).TargetIndex).Quieto < 30 And charlist(charlist(CharIndex).TargetIndex).Inmo = 0 Then  'El quieto sirve para que el bot se entere si te estas moviendo o no, si no lo estás no falla.
        If RandomNum(2) = 1 Then
            If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
            End If
            charlist(CharIndex).Lanzando = 0
            charlist(CharIndex).ComportamientoHechizos = IIf(RandomNum(10) <= 3, 1, 2)  'El mago usa ComportamientoHechizos en vez de ComportamientoCombo porque es el primero que use, despues cambie a Combo, tengo que convertir todo en una sola variable.
        End If
        Exit Sub
    End If
    
    If EnTeam(CharIndex) = 1 Then
        For i = 1 To UBound(Team1)
            If charlist(Team1(i)).Inmo Then
                If RandomNum(3) < 3 Then
                    If CharTimer(CharIndex).Check(TimersIndex.RemoOtro, False) Then
                        Call BotCasteoSpell(CharIndex, Team1(i), 4)
                    End If
                    Exit Sub
                End If
            ElseIf charlist(Team1(i)).MinHP = 0 And SinResu = 0 Then
                If RandomNum(3) < 3 Then
                    If CharTimer(CharIndex).Check(TimersIndex.Resu, False) Then
                        Call BotCasteoSpell(CharIndex, Team1(i), 6)
                    End If
                    Exit Sub
                End If
            End If
        Next i
    Else
        For i = 1 To UBound(Team2)
            If charlist(Team2(i)).Inmo Then
                If RandomNum(3) = 1 Then
                    If CharTimer(CharIndex).Check(TimersIndex.RemoOtro, False) Then
                        Call BotCasteoSpell(CharIndex, Team2(i), 4)
                    End If
                    Exit Sub
                End If
            ElseIf charlist(Team2(i)).MinHP = 0 And SinResu = 0 Then
                If RandomNum(3) = 1 Then
                    If CharTimer(CharIndex).Check(TimersIndex.Resu, False) Then
                        Call BotCasteoSpell(CharIndex, Team2(i), 6)
                    End If
                    Exit Sub
                End If
            End If
        Next i
    End If
    Call GetTargetBot(CharIndex)

    If EnArea(CharIndex, charlist(CharIndex).TargetIndex) Then
        h = 1
        If charlist(CharIndex).ComportamientoHechizos = 1 Then
            If charlist(CharIndex).MinMAN < Hechizo(1).Mana Then
                h = 2
                If charlist(CharIndex).MinMAN < Hechizo(2).Mana Then
                    If charlist(CharIndex).Lanzando = 1 Then
                        charlist(CharIndex).Lanzando = 0
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                    End If
                    charlist(CharIndex).ComportamientoHechizos = IIf(RandomNum(10) <= 3, 1, 2)
                    Exit Sub
                End If
            End If
        ElseIf charlist(CharIndex).ComportamientoHechizos = 2 Then
            charlist(CharIndex).ComportamientoHechizos = 3
        ElseIf charlist(CharIndex).ComportamientoHechizos = 3 Then
            If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
            End If
            charlist(CharIndex).Lanzando = 0
            charlist(CharIndex).ComportamientoHechizos = IIf(RandomNum(10) <= 3, 1, 2)
        End If
    Else: Exit Sub
    End If

    Call MainTimer.Restart(TimersIndex.Lan)
    Call BotCasteoSpell(CharIndex, charlist(CharIndex).TargetIndex, h)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PaladinLanzaSpell
' Author    : Anagrama
' Date      : ???
' Purpose   : Selección de hechizo a lanzar según la situación para el paladín.
'---------------------------------------------------------------------------------------
'
Public Sub PaladinLanzaSpell(ByVal CharIndex As Integer)
    Dim daño As Integer
    Dim h As Byte
    Dim i As Byte
    
    If charlist(CharIndex).Inmo = 1 Then
        charlist(CharIndex).Lanzando = 0
        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
        End If
        If CharTimer(CharIndex).Check(TimersIndex.Remo, False) Then
            Call BotCasteoSpell(CharIndex, CharIndex, 4)
        End If
        Exit Sub
    End If
    
    charlist(CharIndex).Lanzando = 1
    
    If EnTeam(CharIndex) = 1 Then
        For i = 1 To UBound(Team1)
            If charlist(Team1(i)).Inmo Then
                If charlist(CharIndex).ComportamientoCombo = 1 Then
                    If CharTimer(CharIndex).Check(TimersIndex.RemoOtro, False) Then
                        Call BotCasteoSpell(CharIndex, Team1(i), 4)
                    End If
                    Exit Sub
                End If
            ElseIf charlist(Team1(i)).MinHP = 0 And SinResu = 0 Then
                If charlist(CharIndex).ComportamientoCombo = 1 Then
                    If CharTimer(CharIndex).Check(TimersIndex.Resu, False) Then
                        Call BotCasteoSpell(CharIndex, Team1(i), 6)
                    End If
                    Exit Sub
                End If
            End If
        Next i
    Else
        For i = 1 To UBound(Team2)
            If charlist(Team2(i)).Inmo Then
                If charlist(CharIndex).ComportamientoCombo = 1 Then
                    If CharTimer(CharIndex).Check(TimersIndex.RemoOtro, False) Then
                        Call BotCasteoSpell(CharIndex, Team2(i), 4)
                    End If
                    Exit Sub
                End If
            ElseIf charlist(Team2(i)).MinHP = 0 And SinResu = 0 Then
                If charlist(CharIndex).ComportamientoCombo = 1 Then
                    If CharTimer(CharIndex).Check(TimersIndex.Resu, False) Then
                        Call BotCasteoSpell(CharIndex, Team2(i), 6)
                    End If
                    Exit Sub
                End If
            End If
        Next i
    End If
    
    
    If Not MainTimer.Check(TimersIndex.Lan) Then Exit Sub
    
    Call GetTargetBot(CharIndex)
    
    If EnArea(CharIndex, charlist(CharIndex).TargetIndex) Then
        If charlist(charlist(CharIndex).TargetIndex).Inmo Then
            h = 2
            If charlist(CharIndex).ComportamientoCombo = 2 Then
                If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                    h = 5
                    If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                        charlist(CharIndex).Lanzando = 0
                        charlist(CharIndex).ComportamientoCombo = 1
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        Exit Sub
                    End If
                End If
                charlist(CharIndex).ComportamientoCombo = 1
            Else: Exit Sub
            End If
        Else
            h = 2
            If charlist(CharIndex).ComportamientoCombo = 1 Then
                h = 3
                If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                    charlist(CharIndex).Lanzando = 0
                    If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                    End If
                    Exit Sub
                End If
                If RandomNum(Dificultad) <> 1 And charlist(charlist(CharIndex).TargetIndex).Quieto < 30 And charlist(charlist(CharIndex).TargetIndex).Inmo = 0 Then
                    If charlist(CharIndex).Lanzando = 1 Then
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        charlist(CharIndex).Lanzando = 0
                    End If
                    Exit Sub
                End If
                Call BotCasteoSpell(CharIndex, charlist(CharIndex).TargetIndex, h)
                Exit Sub
            ElseIf charlist(CharIndex).ComportamientoCombo = 2 Then
                charlist(CharIndex).ComportamientoCombo = 1
                If RandomNum(Dificultad) <> 1 And charlist(charlist(CharIndex).TargetIndex).Quieto < 30 And charlist(charlist(CharIndex).TargetIndex).Inmo = 0 Then
                    If charlist(CharIndex).Lanzando = 1 Then
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        charlist(CharIndex).Lanzando = 0
                    End If
                    Exit Sub
                End If
                h = 2
                If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                    h = 5
                    If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                        charlist(CharIndex).Lanzando = 0
                        charlist(CharIndex).ComportamientoCombo = 1
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        Exit Sub
                    End If
                End If
            Else
                charlist(CharIndex).ComportamientoCombo = 1
                Exit Sub
            End If
        End If
    Else: Exit Sub
    End If
    
    Call MainTimer.Restart(TimersIndex.Lan)
    Call BotCasteoSpell(CharIndex, charlist(CharIndex).TargetIndex, h)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ClerigoLanzaSpell
' Author    : Anagrama
' Date      : ???
' Purpose   : Selección de hechizo a lanzar según la situación para el clérigo.
'---------------------------------------------------------------------------------------
'
Public Sub ClerigoLanzaSpell(ByVal CharIndex As Integer)
    Dim daño As Integer
    Dim h As Byte
    Dim i As Byte
    
    If charlist(CharIndex).Inmo = 1 Then
        charlist(CharIndex).Lanzando = 0
        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
        End If
        If CharTimer(CharIndex).Check(TimersIndex.Remo, False) Then
            Call BotCasteoSpell(CharIndex, CharIndex, 4)
        End If
        Exit Sub
    End If
    
    If EnTeam(CharIndex) = 1 Then
        For i = 1 To UBound(Team1)
            If charlist(Team1(i)).Inmo Then
                If charlist(CharIndex).ComportamientoCombo <> 3 Then
                    If CharTimer(CharIndex).Check(TimersIndex.RemoOtro, False) Then
                        Call BotCasteoSpell(CharIndex, Team1(i), 4)
                    End If
                    Exit Sub
                End If
            ElseIf charlist(Team1(i)).MinHP = 0 And SinResu = 0 Then
                If charlist(CharIndex).ComportamientoCombo <> 3 Then
                    If CharTimer(CharIndex).Check(TimersIndex.Resu, False) Then
                        Call BotCasteoSpell(CharIndex, Team1(i), 6)
                    End If
                    Exit Sub
                End If
            End If
        Next i
    Else
        For i = 1 To UBound(Team2)
            If charlist(Team2(i)).Inmo Then
                If charlist(CharIndex).ComportamientoCombo <> 3 Then
                    If CharTimer(CharIndex).Check(TimersIndex.RemoOtro, False) Then
                        Call BotCasteoSpell(CharIndex, Team2(i), 4)
                    End If
                    Exit Sub
                End If
            ElseIf charlist(Team2(i)).MinHP = 0 And SinResu = 0 Then
                If charlist(CharIndex).ComportamientoCombo <> 3 Then
                    If CharTimer(CharIndex).Check(TimersIndex.Resu, False) Then
                        Call BotCasteoSpell(CharIndex, Team2(i), 6)
                    End If
                    Exit Sub
                End If
            End If
        Next i
    End If
    
    Call GetTargetBot(CharIndex)
    
    If EnArea(CharIndex, charlist(CharIndex).TargetIndex) Then
        If charlist(charlist(CharIndex).TargetIndex).Inmo Then
            h = 1
            If charlist(CharIndex).ComportamientoCombo = 3 Then
                If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                    h = 2
                    If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                        charlist(CharIndex).Lanzando = 0
                        charlist(CharIndex).ComportamientoCombo = 1
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        Exit Sub
                    End If
                End If
                charlist(CharIndex).ComportamientoCombo = 1
                If charlist(CharIndex).ComportamientoPotas = 4 Then
                    charlist(CharIndex).Lanzando = 0
                    If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                    End If
                End If
            Else: Exit Sub
            End If
        Else
            If charlist(CharIndex).ComportamientoCombo = 1 Then
                h = 3
                If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                    charlist(CharIndex).Lanzando = 0
                    If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                    End If
                    Exit Sub
                End If
                If RandomNum(Dificultad) <> 1 And charlist(charlist(CharIndex).TargetIndex).Quieto < 30 And charlist(charlist(CharIndex).TargetIndex).Inmo = 0 Then
                    If charlist(CharIndex).Lanzando = 1 Then
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        charlist(CharIndex).Lanzando = 0
                    End If
                    charlist(CharIndex).ComportamientoCombo = IIf(RandomNum(4) = 1, 2, 1)
                    Exit Sub
                End If
                Call BotCasteoSpell(CharIndex, charlist(CharIndex).TargetIndex, h)
                Exit Sub
            ElseIf charlist(CharIndex).ComportamientoCombo = 2 Then
                If RandomNum(Dificultad) <> 1 And charlist(charlist(CharIndex).TargetIndex).Quieto < 30 And charlist(charlist(CharIndex).TargetIndex).Inmo = 0 Then
                    If charlist(CharIndex).Lanzando = 1 Then
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        charlist(CharIndex).Lanzando = 0
                    End If
                    charlist(CharIndex).ComportamientoCombo = 1
                    Exit Sub
                End If
                h = 1
                If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                    h = 2
                    If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                        charlist(CharIndex).Lanzando = 0
                        charlist(CharIndex).ComportamientoCombo = 1
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        Exit Sub
                    End If
                End If
                charlist(CharIndex).ComportamientoCombo = 3
            ElseIf charlist(CharIndex).ComportamientoCombo = 3 Then
                If RandomNum(Dificultad) <> 1 And charlist(charlist(CharIndex).TargetIndex).Quieto < 30 And charlist(charlist(CharIndex).TargetIndex).Inmo = 0 Then
                    If charlist(CharIndex).Lanzando = 1 Then
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        charlist(CharIndex).Lanzando = 0
                    End If
                    charlist(CharIndex).ComportamientoCombo = 1
                    Exit Sub
                End If
                h = 1
                If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                    h = 2
                    If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                        charlist(CharIndex).Lanzando = 0
                        charlist(CharIndex).ComportamientoCombo = 1
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        Exit Sub
                    End If
                End If
            Else
                charlist(CharIndex).ComportamientoCombo = 1
                Exit Sub
            End If
        End If
    Else: Exit Sub
    End If

    Call MainTimer.Restart(TimersIndex.Lan)
    Call BotCasteoSpell(CharIndex, charlist(CharIndex).TargetIndex, h)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : BardoLanzaSpell
' Author    : Anagrama
' Date      : ???
' Purpose   : Selección de hechizo a lanzar según la situación para el bardo.
'---------------------------------------------------------------------------------------
'
Public Sub BardoLanzaSpell(ByVal CharIndex As Integer)
    Dim daño As Integer
    Dim h As Byte
    Dim i As Byte
    
    If charlist(CharIndex).Inmo = 1 Then
        charlist(CharIndex).Lanzando = 0
        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
        End If
        If CharTimer(CharIndex).Check(TimersIndex.Remo, False) Then
            Call BotCasteoSpell(CharIndex, CharIndex, 4)
        End If
        Exit Sub
    End If
    
    charlist(CharIndex).Lanzando = 1
    
    If EnTeam(CharIndex) = 1 Then
        For i = 1 To UBound(Team1)
            If charlist(Team1(i)).Inmo Then
                If charlist(CharIndex).ComportamientoCombo <> 3 Then
                    If CharTimer(CharIndex).Check(TimersIndex.RemoOtro, False) Then
                        Call BotCasteoSpell(CharIndex, Team1(i), 4)
                    End If
                    Exit Sub
                End If
            ElseIf charlist(Team1(i)).MinHP = 0 And SinResu = 0 Then
                If charlist(CharIndex).ComportamientoCombo <> 3 Then
                    If CharTimer(CharIndex).Check(TimersIndex.Resu, False) Then
                        Call BotCasteoSpell(CharIndex, Team1(i), 6)
                    End If
                    Exit Sub
                End If
            End If
        Next i
    Else
        For i = 1 To UBound(Team2)
            If charlist(Team2(i)).Inmo Then
                If charlist(CharIndex).ComportamientoCombo <> 3 Then
                    If CharTimer(CharIndex).Check(TimersIndex.RemoOtro, False) Then
                        Call BotCasteoSpell(CharIndex, Team2(i), 4)
                    End If
                    Exit Sub
                End If
            ElseIf charlist(Team2(i)).MinHP = 0 And SinResu = 0 Then
                If charlist(CharIndex).ComportamientoCombo <> 3 Then
                    If CharTimer(CharIndex).Check(TimersIndex.Resu, False) Then
                        Call BotCasteoSpell(CharIndex, Team2(i), 6)
                    End If
                    Exit Sub
                End If
            End If
        Next i
    End If
    
    Call GetTargetBot(CharIndex)
    
    If EnArea(CharIndex, charlist(CharIndex).TargetIndex) Then
        If charlist(charlist(CharIndex).TargetIndex).Inmo Then
            h = 1
            If charlist(CharIndex).ComportamientoCombo <> 1 Then
                If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                    h = 2
                    If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                        charlist(CharIndex).Lanzando = 0
                        charlist(CharIndex).ComportamientoCombo = IIf(RandomNum(10) = 1, 1, 2)
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        Exit Sub
                    End If
                End If
                charlist(CharIndex).ComportamientoCombo = 3
            Else: Exit Sub
            End If
        Else
            If charlist(CharIndex).ComportamientoCombo = 1 Then
                h = 3
                If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                    charlist(CharIndex).Lanzando = 0
                    If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                    End If
                    Exit Sub
                End If
                If RandomNum(Dificultad) <> 1 And charlist(charlist(CharIndex).TargetIndex).Quieto < 30 And charlist(charlist(CharIndex).TargetIndex).Inmo = 0 Then
                    If charlist(CharIndex).Lanzando = 1 Then
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        charlist(CharIndex).Lanzando = 0
                    End If
                    charlist(CharIndex).ComportamientoCombo = IIf(RandomNum(10) = 1, 1, 2)
                    Exit Sub
                End If
                Call BotCasteoSpell(CharIndex, charlist(CharIndex).TargetIndex, h)
                Exit Sub
            ElseIf charlist(CharIndex).ComportamientoCombo = 2 Then
                If RandomNum(Dificultad) <> 1 And charlist(charlist(CharIndex).TargetIndex).Quieto < 30 And charlist(charlist(CharIndex).TargetIndex).Inmo = 0 Then
                    If charlist(CharIndex).Lanzando = 1 Then
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        charlist(CharIndex).Lanzando = 0
                    End If
                    charlist(CharIndex).ComportamientoCombo = IIf(RandomNum(10) = 1, 1, 2)
                    Exit Sub
                End If
                h = 1
                If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                    h = 2
                    If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                        charlist(CharIndex).Lanzando = 0
                        charlist(CharIndex).ComportamientoCombo = IIf(RandomNum(10) = 1, 1, 2)
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        Exit Sub
                    End If
                End If
                charlist(CharIndex).ComportamientoCombo = 3
            ElseIf charlist(CharIndex).ComportamientoCombo = 3 Then
                If RandomNum(Dificultad) <> 1 And charlist(charlist(CharIndex).TargetIndex).Quieto < 30 And charlist(charlist(CharIndex).TargetIndex).Inmo = 0 Then
                    If charlist(CharIndex).Lanzando = 1 Then
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        charlist(CharIndex).Lanzando = 0
                    End If
                    charlist(CharIndex).ComportamientoCombo = IIf(RandomNum(10) = 1, 1, 2)
                    Exit Sub
                End If
                h = 1
                If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                    h = 2
                    If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                        charlist(CharIndex).Lanzando = 0
                        charlist(CharIndex).ComportamientoCombo = IIf(RandomNum(10) = 1, 1, 2)
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                        Exit Sub
                    End If
                End If
            Else
                charlist(CharIndex).ComportamientoCombo = IIf(RandomNum(10) = 1, 1, 2)
                Exit Sub
            End If
        End If
    Else: Exit Sub
    End If

    Call MainTimer.Restart(TimersIndex.Lan)
    Call BotCasteoSpell(CharIndex, charlist(CharIndex).TargetIndex, h)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DruidaLanzaSpell
' Author    : Anagrama
' Date      : ???
' Purpose   : Selección de hechizo a lanzar según la situación para el druida.
'---------------------------------------------------------------------------------------
'
Public Sub DruidaLanzaSpell(ByVal CharIndex As Integer)
    Dim daño As Integer
    Dim h As Byte
    Dim i As Byte
    
    If charlist(CharIndex).Inmo = 1 Then
        charlist(CharIndex).Lanzando = 0
        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
        End If
        If CharTimer(CharIndex).Check(TimersIndex.Remo, False) Then
            Call BotCasteoSpell(CharIndex, CharIndex, 4)
        End If
        Exit Sub
    End If
    
    charlist(CharIndex).Lanzando = 1
    
    If EnTeam(CharIndex) = 1 Then
        For i = 1 To UBound(Team1)
            If charlist(Team1(i)).Inmo Then
                If charlist(CharIndex).ComportamientoCombo = 1 Then
                    If CharTimer(CharIndex).Check(TimersIndex.RemoOtro, False) Then
                        Call BotCasteoSpell(CharIndex, Team1(i), 4)
                    End If
                    Exit Sub
                End If
            ElseIf charlist(Team1(i)).MinHP = 0 And SinResu = 0 Then
                If charlist(CharIndex).ComportamientoCombo = 1 Then
                    If CharTimer(CharIndex).Check(TimersIndex.Resu, False) Then
                        Call BotCasteoSpell(CharIndex, Team1(i), 6)
                    End If
                    Exit Sub
                End If
            End If
        Next i
    Else
        For i = 1 To UBound(Team2)
            If charlist(Team2(i)).Inmo Then
                If charlist(CharIndex).ComportamientoCombo = 1 Then
                    If CharTimer(CharIndex).Check(TimersIndex.RemoOtro, False) Then
                        Call BotCasteoSpell(CharIndex, Team2(i), 4)
                    End If
                    Exit Sub
                End If
            ElseIf charlist(Team2(i)).MinHP = 0 And SinResu = 0 Then
                If charlist(CharIndex).ComportamientoCombo = 1 Then
                    If CharTimer(CharIndex).Check(TimersIndex.Resu, False) Then
                        Call BotCasteoSpell(CharIndex, Team2(i), 6)
                    End If
                    Exit Sub
                End If
            End If
        Next i
    End If
    
    Call GetTargetBot(CharIndex)
    
    If EnArea(CharIndex, charlist(CharIndex).TargetIndex) Then
        If charlist(CharIndex).ComportamientoCombo = 1 Then
            If RandomNum(Dificultad) <> 1 And charlist(charlist(CharIndex).TargetIndex).Quieto < 30 And charlist(charlist(CharIndex).TargetIndex).Inmo = 0 Then
                If charlist(CharIndex).Lanzando = 1 Then
                    If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                    End If
                    charlist(CharIndex).Lanzando = 0
                End If
                Exit Sub
            End If
            h = 1
            If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                h = 2
                If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                    charlist(CharIndex).Lanzando = 0
                    If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                    End If
                    Exit Sub
                End If
            End If
            charlist(CharIndex).ComportamientoCombo = 2
        ElseIf charlist(CharIndex).ComportamientoCombo = 2 Then
            If RandomNum(Dificultad) <> 1 And charlist(charlist(CharIndex).TargetIndex).Quieto < 30 And charlist(charlist(CharIndex).TargetIndex).Inmo = 0 Then
                If charlist(CharIndex).Lanzando = 1 Then
                    If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                    End If
                    charlist(CharIndex).Lanzando = 0
                End If
                charlist(CharIndex).ComportamientoCombo = 1
                Exit Sub
            End If
            h = 1
            If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                h = 2
                If charlist(CharIndex).MinMAN < Hechizo(h).Mana Then
                    charlist(CharIndex).Lanzando = 0
                    charlist(CharIndex).ComportamientoCombo = 1
                    If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                    End If
                    Exit Sub
                End If
            End If
        Else
            charlist(CharIndex).ComportamientoCombo = IIf(RandomNum(10) = 1, 1, 2)
            Exit Sub
        End If
    Else: Exit Sub
    End If
    
    Call MainTimer.Restart(TimersIndex.Lan)
    Call BotCasteoSpell(CharIndex, charlist(CharIndex).TargetIndex, h)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : BotTomaPot
' Author    : Anagrama
' Date      : ???
' Purpose   : Revisa las condiciones necesarias para tomar una poción, segun cual necesita, si está en inventario,
' si está en hechizos, si está pasando de H a I, si está usando nada más U, si usa U CLICK, etc.
'---------------------------------------------------------------------------------------
'
Public Sub BotTomaPot(ByVal CharIndex As Integer)
    'Rojas.
    If charlist(CharIndex).MinHP < charlist(CharIndex).MaxHP And (charlist(CharIndex).TipoPocion = 1 Or (charlist(CharIndex).Lanzando = 0 And CharTimer(CharIndex).Check(TimersIndex.WaitP, False))) Then
        charlist(CharIndex).TipoPocion = 1
        charlist(CharIndex).MinHP = charlist(CharIndex).MinHP + 30
        If charlist(CharIndex).MinHP > charlist(CharIndex).MaxHP Then charlist(CharIndex).MinHP = charlist(CharIndex).MaxHP
        Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(46, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
        If charlist(CharIndex).MinHP >= charlist(CharIndex).MaxHP * 0.7 And UserParalizado = False Then
            If charlist(CharIndex).ComportamientoPotas = 3 Or charlist(CharIndex).ComportamientoPotas = 4 Then
                charlist(CharIndex).ComportamientoPotas = IIf(RandomNum(10) <= 3, 1, 2)
            End If
        ElseIf charlist(CharIndex).MinHP <= charlist(CharIndex).MaxHP * 0.7 And charlist(CharIndex).Inmo = 1 And charlist(CharIndex).Ai <> eBotAi.Guerrero Then
            If RandomNum(10) > Dificultad Then
                If charlist(CharIndex).ComportamientoPotas = 3 Or charlist(CharIndex).ComportamientoPotas = 4 Then
                    charlist(CharIndex).ComportamientoPotas = IIf(RandomNum(10) <= 3, 1, 2)
                End If
            End If
        End If
        If charlist(CharIndex).MinHP = charlist(CharIndex).MaxHP Then
            If charlist(CharIndex).ComportamientoPotas = 3 Or charlist(CharIndex).ComportamientoPotas = 4 Then
                charlist(CharIndex).ComportamientoPotas = IIf(RandomNum(10) <= 3, 1, 2)
            End If
            If charlist(CharIndex).Ai <> eBotAi.Guerrero Then
                charlist(CharIndex).TipoPocion = 2
            End If
            If charlist(CharIndex).Lanzando = 0 Then
                If CharTimer(CharIndex).Check(TimersIndex.WaitH, False) Then
                    Call CharTimer(CharIndex).Restart(TimersIndex.WaitH)
                End If
                charlist(CharIndex).Lanzando = 1
            End If
        End If
    'Azules.
    ElseIf charlist(CharIndex).MinMAN < charlist(CharIndex).MaxMAN And (charlist(CharIndex).TipoPocion = 2 Or (charlist(CharIndex).Lanzando = 0 And CharTimer(CharIndex).Check(TimersIndex.WaitP, False))) Then
        charlist(CharIndex).TipoPocion = 2
        If charlist(CharIndex).MinHP < charlist(CharIndex).MaxHP / 2 Then
            If charlist(CharIndex).Lanzando = 1 Then
                If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                    Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                End If
                charlist(CharIndex).Lanzando = 0
            End If
        End If
        charlist(CharIndex).MinMAN = charlist(CharIndex).MinMAN + (charlist(CharIndex).MaxMAN * 4) / 100 + UserLvl \ 2 + 40 / UserLvl
        If charlist(CharIndex).MinMAN > charlist(CharIndex).MaxMAN Then charlist(CharIndex).MinMAN = charlist(CharIndex).MaxMAN
        Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(46, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
        If charlist(CharIndex).MinMAN = charlist(CharIndex).MaxMAN Then
            If charlist(CharIndex).Lanzando = 0 Then
                If CharTimer(CharIndex).Check(TimersIndex.WaitH, False) Then
                    Call CharTimer(CharIndex).Restart(TimersIndex.WaitH)
                End If
                If charlist(CharIndex).MinHP = charlist(CharIndex).MaxHP Then charlist(CharIndex).Lanzando = 1
            End If
            If charlist(CharIndex).ComportamientoPotas = 5 Then charlist(CharIndex).ComportamientoPotas = IIf(RandomNum(10) <= 3, 1, 2)
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : BotCasteoSpell
' Author    : Anagrama
' Date      : ???
' Purpose   : En un intento de hacer un poco mas legible el sub BotLanzaSpell, pase la parte de efectos y la verdadera
' mecánica de los hechizos aca, de algo sirvió.
'---------------------------------------------------------------------------------------
'
Public Sub BotCasteoSpell(ByVal CharIndex As Integer, ByVal TargetIndex As Integer, ByVal HIndex As Byte)
    Dim daño As Integer
    Dim i As Byte
    
    If HIndex = 0 Then Exit Sub 'Imposible que pase pero nunca se sabe.
    
    If charlist(CharIndex).MinMAN < Hechizo(HIndex).Mana Then Exit Sub 'Revisa el costo de mana.
    If charlist(CharIndex).MinSTA < Hechizo(HIndex).Sta Then Exit Sub 'Revisa el costo de energia.
    If charlist(CharIndex).MinHP < charlist(CharIndex).MaxHP Then charlist(CharIndex).Lanzando = 1 'Si pasa a inventario ahora.
    
    If HIndex = 4 Then 'Remo
        If EnArea(CharIndex, TargetIndex) = 0 Then Exit Sub
        charlist(TargetIndex).Inmo = 0
        If charlist(TargetIndex).Bot = 0 Then
            Call WriteParalizeOK(TargetIndex, charlist(TargetIndex).Inmo)
            Call WriteConsoleMsg(TargetIndex, charlist(CharIndex).Nombre & " te ha lanzado " & Hechizo(HIndex).name & ".", FontTypeNames.FONTTYPE_FIGHT)
        End If
    ElseIf HIndex = 3 Then 'Inmo
        charlist(TargetIndex).Inmo = 1
        If charlist(TargetIndex).Bot = 0 Then
            Call WriteParalizeOK(TargetIndex, charlist(TargetIndex).Inmo)
            Call WriteConsoleMsg(TargetIndex, charlist(CharIndex).Nombre & " te ha lanzado " & Hechizo(HIndex).name & ".", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call CharTimer(TargetIndex).SetInterval(TimersIndex.Remo, INT_REMO + RandomNumber(100 * Dificultad, 200 * Dificultad))
            Call CharTimer(TargetIndex).Restart(TimersIndex.Remo)
        End If
        Call GetTargetBotTeam(EnTeam(CharIndex))
        If EnTeam(CharIndex) = 1 Then
            For i = 1 To UBound(Team2)
                Call CharTimer(Team2(i)).SetInterval(TimersIndex.RemoOtro, INT_REMOOTRO + RandomNumber(100 * Dificultad, 200 * Dificultad))
                Call CharTimer(Team2(i)).Restart(TimersIndex.RemoOtro)
            Next i
        Else
            If UBound(Team1) > 1 Then
                For i = 2 To UBound(Team1)
                    Call CharTimer(Team1(i)).SetInterval(TimersIndex.RemoOtro, INT_REMOOTRO + RandomNumber(100 * Dificultad, 200 * Dificultad))
                    Call CharTimer(Team1(i)).Restart(TimersIndex.RemoOtro)
                Next i
            End If
        End If
    ElseIf HIndex = 6 Then 'Resu
        If EnArea(CharIndex, TargetIndex) = 0 Then Exit Sub
        daño = charlist(CharIndex).MaxHP - charlist(CharIndex).MinHP * (1 - charlist(TargetIndex).Lvl * 0.015)
        If ResuNoVida = 0 Then
            If charlist(CharIndex).MinHP <= daño Then Exit Sub
            charlist(CharIndex).MinHP = charlist(CharIndex).MinHP - daño
        End If
        charlist(TargetIndex).MinHP = 1
        charlist(TargetIndex).MinMAN = 0
        charlist(TargetIndex).body = BodyData(BodyClaseRaza(charlist(TargetIndex).Ai, charlist(TargetIndex).Raza))
        charlist(TargetIndex).Head = HeadData(HeadRazaGenero(charlist(TargetIndex).Raza, charlist(TargetIndex).Genero))
        charlist(TargetIndex).Arma = WeaponAnimData(ArmaClase(charlist(TargetIndex).Ai))
        charlist(TargetIndex).Escudo = ShieldAnimData(EscudoClase(charlist(TargetIndex).Ai))
        charlist(TargetIndex).Casco = CascoAnimData(CascoClase(charlist(TargetIndex).Ai))
        If charlist(CharIndex).MinHP <= 0 Then
            charlist(CharIndex).MinHP = 0
            Call CharDie(CharIndex)
        End If
        If charlist(TargetIndex).Bot = 1 Then
            charlist(TargetIndex).ComportamientoPotas = 4
            charlist(TargetIndex).Lanzando = 0
        Else
            Call WriteUpdateCharStats(TargetIndex)
        End If
        Call ServerSendData(SendTarget.ToAllButIndex, UserCharIndex, PrepareMessageCharacterChange(charlist(TargetIndex).iBody, charlist(TargetIndex).iHead, charlist(TargetIndex).heading, TargetIndex _
                            , charlist(TargetIndex).iArma, charlist(TargetIndex).iEscudo, charlist(TargetIndex).FxIndex, charlist(TargetIndex).FX.Loops, charlist(TargetIndex).iCasco))
        Call DarPrioridadTarget(EnTeam(TargetIndex))
    Else 'Apoca, desca, tormenta
        daño = RandomNumber(Hechizo(HIndex).MinHP, Hechizo(HIndex).MaxHP)
        daño = daño + (daño * 3 * charlist(CharIndex).Lvl) / 100
        daño = daño * charlist(CharIndex).DM
        daño = daño - RandomNumber(charlist(TargetIndex).MinRM, charlist(TargetIndex).MaxRM)
        charlist(TargetIndex).MinHP = charlist(TargetIndex).MinHP - daño
        If charlist(TargetIndex).MinHP < 0 Then charlist(TargetIndex).MinHP = 0
        If charlist(TargetIndex).Bot = 0 Then
            Call WriteConsoleMsg(TargetIndex, charlist(CharIndex).Nombre & " te ha lanzado " & Hechizo(HIndex).name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(TargetIndex, charlist(CharIndex).Nombre & " te ha pegado por " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteUpdateCharStats(TargetIndex)
        End If
        If charlist(TargetIndex).MinHP <= 0 Then
            charlist(TargetIndex).MinHP = 0
            If charlist(TargetIndex).Bot = 0 Then
                Call WriteConsoleMsg(TargetIndex, charlist(CharIndex).Nombre & " te ha matado.", FontTypeNames.FONTTYPE_FIGHT)
            End If
            
            Call ServerSendData(SendTarget.ToAllButIndex, TargetIndex, PrepareMessageConsoleMsg(charlist(CharIndex).Nombre & " ha matado a " & charlist(TargetIndex).Nombre & ".", FontTypeNames.FONTTYPE_FIGHT))
            
            Call CharDie(TargetIndex)
            Call ResetDuelo(CharIndex)
            Call GetTargetBotTeam(EnTeam(CharIndex))
        End If
    End If

    Call ServerSendData(SendTarget.ToAll, TargetIndex, PrepareMessageChatOverHead(Hechizo(HIndex).Palabras, CharIndex, vbCyan)) 'Palabras mágicas.
    Call ServerSendData(SendTarget.ToAll, TargetIndex, PrepareMessagePlayWave(Hechizo(HIndex).WAV, charlist(TargetIndex).Pos.X, charlist(TargetIndex).Pos.Y)) 'Sonido.
    Call ServerSendData(SendTarget.ToAll, TargetIndex, PrepareMessageCreateFX(TargetIndex, Hechizo(HIndex).FX, 0)) 'Gráfico al objetivo.
    charlist(CharIndex).MinMAN = charlist(CharIndex).MinMAN - Hechizo(HIndex).Mana 'Le saca la mana.
    charlist(CharIndex).MinSTA = charlist(CharIndex).MinSTA - Hechizo(HIndex).Sta 'Le saca la energia.
End Sub

Public Sub GetTargetBot(ByVal CharIndex As Integer)
'---------------------------------------------------------------------------------------
' Procedure : GetTargetBot
' Author    : Anagrama
' Date      : ???
' Purpose   : Verifica equipos y decide el objetivo del bot.
'---------------------------------------------------------------------------------------
'
    Dim i As Byte
    Dim a As Byte
    Dim b As Byte
    Dim MaxPrioridad As Byte
    Dim MinPrioridad As Byte
    Dim tmpChar() As Byte
    Dim num As Byte
    
    If EnTeam(CharIndex) = 1 Then
        If UBound(Team2) = 0 Then
            charlist(CharIndex).TargetIndex = 0
            Exit Sub
        End If
        For i = 1 To UBound(Team2)
            If Team2(i) > 0 Then
                If charlist(Team2(i)).MinHP > 0 Then
                    If charlist(Team2(i)).Inmo = 1 Then
                        charlist(CharIndex).TargetIndex = Team2(i)
                        Exit Sub
                    End If
                End If
            End If
        Next i
        For a = 1 To UBound(Team2)
            If Prioridad2(a).Char > 0 Then
                If MaxPrioridad = 0 Then
                    If charlist(Prioridad2(a).Char).MinHP > 0 Then MaxPrioridad = a
                Else
                    If charlist(Prioridad2(a).Char).MinHP > 0 Then MinPrioridad = a
                    Exit For
                End If
            End If
        Next
        If MinPrioridad = 0 Then
            charlist(CharIndex).TargetIndex = Prioridad2(MaxPrioridad).Char
        ElseIf UBound(Team2) = 2 Then
            If num > Prioridad2(MaxPrioridad).Probabilidad Then
                charlist(CharIndex).TargetIndex = Prioridad2(MinPrioridad).Char
            Else
                charlist(CharIndex).TargetIndex = Prioridad2(MaxPrioridad).Char
            End If
        Else
            num = RandomNum(100)
            If num > Prioridad2(MinPrioridad).Probabilidad + Prioridad2(MaxPrioridad).Probabilidad Then
                For b = 1 To UBound(Team2)
                    ReDim tmpChar(1 To 1) As Byte
                    If b <> MaxPrioridad And b <> MinPrioridad And charlist(Prioridad2(b).Char).MinHP > 0 Then
                        If tmpChar(1) <> 0 Then
                            ReDim Preserve tmpChar(1 To UBound(tmpChar) + 1) As Byte
                            tmpChar(UBound(tmpChar)) = b
                        Else
                            tmpChar(1) = b
                        End If
                    End If
                Next
                If tmpChar(1) > 0 Then
                    charlist(CharIndex).TargetIndex = Prioridad2(tmpChar(RandomNum(UBound(tmpChar)))).Char
                Else: charlist(CharIndex).TargetIndex = MinPrioridad
                End If
            ElseIf num < Prioridad2(MinPrioridad).Probabilidad Then
                charlist(CharIndex).TargetIndex = Prioridad2(MinPrioridad).Char
            Else
                charlist(CharIndex).TargetIndex = Prioridad2(MaxPrioridad).Char
            End If
        End If
    Else
        If UBound(Team1) = 0 Then
            charlist(CharIndex).TargetIndex = 0
            Exit Sub
        End If
        For i = 1 To UBound(Team1)
            If Team1(i) > 0 Then
                If charlist(Team1(i)).MinHP > 0 Then
                    If charlist(Team1(i)).Inmo = 1 Then
                        charlist(CharIndex).TargetIndex = Team1(i)
                        Exit Sub
                    End If
                End If
            End If
        Next i
        For a = 1 To UBound(Team1)
            If Prioridad1(a).Char > 0 Then
                If MaxPrioridad = 0 Then
                    If charlist(Prioridad1(a).Char).MinHP > 0 Then MaxPrioridad = a
                Else
                    If charlist(Prioridad1(a).Char).MinHP > 0 Then MinPrioridad = a
                    Exit For
                End If
            End If
        Next
        If MinPrioridad = 0 Then
            charlist(CharIndex).TargetIndex = Prioridad1(MaxPrioridad).Char
        ElseIf UBound(Team1) = 2 Then
            If num > Prioridad1(MaxPrioridad).Probabilidad Then
                charlist(CharIndex).TargetIndex = Prioridad1(MinPrioridad).Char
            Else
                charlist(CharIndex).TargetIndex = Prioridad1(MaxPrioridad).Char
            End If
        Else
            num = RandomNum(100)
            If num > Prioridad1(MinPrioridad).Probabilidad + Prioridad1(MaxPrioridad).Probabilidad Then
                For b = 1 To UBound(Team1)
                    ReDim tmpChar(1 To 1) As Byte
                    If b <> MaxPrioridad And b <> MinPrioridad And charlist(Prioridad1(b).Char).MinHP > 0 Then
                        If tmpChar(1) <> 0 Then
                            ReDim Preserve tmpChar(1 To UBound(tmpChar) + 1) As Byte
                            tmpChar(UBound(tmpChar)) = b
                        Else
                            tmpChar(1) = b
                        End If
                    End If
                Next
                If tmpChar(1) > 0 Then
                    charlist(CharIndex).TargetIndex = Prioridad1(tmpChar(RandomNum(UBound(tmpChar)))).Char
                Else: charlist(CharIndex).TargetIndex = MinPrioridad
                End If
            ElseIf num < Prioridad1(MinPrioridad).Probabilidad Then
                charlist(CharIndex).TargetIndex = Prioridad1(MinPrioridad).Char
            Else
                charlist(CharIndex).TargetIndex = Prioridad1(MaxPrioridad).Char
            End If
        End If
    End If
End Sub

Public Sub GetTargetBotTeam(ByVal Team As Byte)
'---------------------------------------------------------------------------------------
' Procedure : GetTargetBotTeam
' Author    : Anagrama
' Date      : ???
' Purpose   : Cuando alguien muere se revisa el siguiente objetivo para los bots.
'---------------------------------------------------------------------------------------
'

    Dim i As Byte
    
    If Team = 1 Then
        For i = 1 To UBound(Team1)
            If Team1(i) > 0 Then
                If charlist(Team1(i)).MinHP > 0 And charlist(Team1(i)).Bot = 1 Then Call GetTargetBot(Team1(i))
            End If
        Next i
    Else
        For i = 1 To UBound(Team2)
            If Team2(i) > 0 Then
                If charlist(Team2(i)).MinHP > 0 And charlist(Team2(i)).Bot = 1 Then Call GetTargetBot(Team2(i))
            End If
        Next i
    End If
End Sub

Public Sub BotLanzaFlecha(ByVal CharIndex As Integer)
'---------------------------------------------------------------------------------------
' Procedure : BotLanzaFlecha
' Author    : Anagrama
' Date      : ???
' Purpose   : Lanza la flecha si puede hacerlo y si no falla.
'---------------------------------------------------------------------------------------
'
    
    If Not CharTimer(CharIndex).Check(TimersIndex.Attack, False) Then Exit Sub  'Check if arrows interval has finished.
    If Not CharTimer(CharIndex).Check(TimersIndex.Arrows) Then Exit Sub

    Call GetTargetBot(CharIndex)
    
    If charlist(CharIndex).ComportamientoPotas < 3 Then
        If EnArea(CharIndex, charlist(CharIndex).TargetIndex) Then
            If RandomNum(Dificultad) = 1 Or charlist(charlist(CharIndex).TargetIndex).Quieto >= 30 Or charlist(charlist(CharIndex).TargetIndex).Inmo = 1 Then
                Call UsuarioAtacaUsuario(CharIndex, charlist(CharIndex).TargetIndex)
                charlist(CharIndex).Lanzando = 1
            End If
        End If
    End If
End Sub

Public Sub BotEquiparItem(ByVal CharIndex As Integer, ByVal Item As Byte)
'---------------------------------------------------------------------------------------
' Procedure : BotEquiparItem
' Author    : Anagrama
' Date      : ???
' Purpose   : Equipa el item que se le solicita.
'---------------------------------------------------------------------------------------
'

    Select Case Item
        Case 3
            Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(25, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
            charlist(CharIndex).Arma = WeaponAnimData(3)
            charlist(CharIndex).ArmaMinHit = ItemData(1).MinHit
            charlist(CharIndex).ArmaMaxHit = ItemData(1).MaxHit
            charlist(CharIndex).Refuerzo = 0
            charlist(CharIndex).TipoArma = 1
        Case 4
            Call ServerSendData(SendTarget.ToAll, CharIndex, PrepareMessagePlayWave(25, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y))
            charlist(CharIndex).Arma = WeaponAnimData(5)
            charlist(CharIndex).ArmaMinHit = ItemData(2).MinHit
            charlist(CharIndex).ArmaMaxHit = ItemData(2).MaxHit
            charlist(CharIndex).Refuerzo = ItemData(2).Refuerzo
            charlist(CharIndex).TipoArma = 2
        Case 5
            charlist(CharIndex).MinRM = ItemData(3).MinRM
            charlist(CharIndex).MaxRM = ItemData(3).MaxRM
            charlist(CharIndex).DM = ItemData(3).DM
        Case 6
            charlist(CharIndex).MinRM = ItemData(4).MinRM
            charlist(CharIndex).MaxRM = ItemData(4).MaxRM
            charlist(CharIndex).DM = ItemData(4).DM
    End Select
End Sub

Public Sub GuerreroAI(ByVal CharIndex As Integer)
'---------------------------------------------------------------------------------------
' Procedure : GuerreroAI
' Author    : Anagrama
' Date      : ???
' Purpose   : Nucleo de la IA del guerrero, decide si moverse, atacar o potear.
'---------------------------------------------------------------------------------------
'

    If charlist(CharIndex).Moving = 0 Then
        If charlist(CharIndex).Inmo = 0 Then
            Call MoveBot(CharIndex)
        End If
    End If
    If charlist(CharIndex).MinHP > 0 Then
        If charlist(CharIndex).TipoArma = 2 Then
            If MainTimer.Check(TimersIndex.Lan, False) = True Then
                If charlist(CharIndex).ComportamientoPotas < 3 Then
                    Call BotLanzaFlecha(CharIndex)
                End If
            End If
        End If
        If CharTimer(CharIndex).Check(TimersIndex.GolpeU, False) Then
            If CharTimer(CharIndex).Check(TimersIndex.UseItem, False) Then
                If charlist(CharIndex).Lanzando = 0 And CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                    If CharTimer(CharIndex).Check(TimersIndex.UseItemWithDblClick, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItem)
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItemWithDblClick)
                        Call BotTomaPot(CharIndex)
                    End If
                Else
                    If CharTimer(CharIndex).Check(TimersIndex.UseItemWithU, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItem)
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItemWithU)
                        Call BotTomaPot(CharIndex)
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub MagoAI(ByVal CharIndex As Integer)
'---------------------------------------------------------------------------------------
' Procedure : MagoAI
' Author    : Anagrama
' Date      : ???
' Purpose   : Nucleo de la IA del mago, decide si moverse, atacar o potear.
'---------------------------------------------------------------------------------------
'
    If charlist(CharIndex).Moving = 0 Then
        If charlist(CharIndex).Inmo = 0 Then
            Call MoveBot(CharIndex)
        End If
    End If
    If charlist(CharIndex).MinHP > 0 Then
        If charlist(CharIndex).Inmo = 1 Or MainTimer.Check(TimersIndex.Lan, False) = True Then
            If charlist(CharIndex).ComportamientoPotas <> 3 Or charlist(CharIndex).MinHP > charlist(CharIndex).MaxHP * 0.7 Then
                If CharTimer(CharIndex).Check(TimersIndex.WaitH, False) Then
                    If charlist(CharIndex).ComportamientoHechizos = 2 Then
                        If charlist(CharIndex).MaxMAN >= 1790 Then
                            If charlist(CharIndex).MinMAN >= 1790 Then
                                If CharTimer(CharIndex).Check(TimersIndex.CastSpell) Then
                                    Call BotLanzaSpell(CharIndex)
                                End If
                            Else
                                If charlist(CharIndex).Lanzando = 1 Then
                                    If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                                        Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                                        charlist(CharIndex).Lanzando = 0
                                    End If
                                End If
                            End If
                        Else
                            If charlist(CharIndex).MinMAN >= 1200 Then
                                If CharTimer(CharIndex).Check(TimersIndex.CastSpell) Then
                                    Call BotLanzaSpell(CharIndex)
                                End If
                            Else
                                If charlist(CharIndex).Lanzando = 1 Then
                                    If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                                        Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                                        charlist(CharIndex).Lanzando = 0
                                    End If
                                End If
                            End If
                        End If
                    ElseIf charlist(CharIndex).ComportamientoHechizos = 3 Then
                        If charlist(CharIndex).MinMAN >= 1000 Then
                            If CharTimer(CharIndex).Check(TimersIndex.CastSpell) Then
                                Call BotLanzaSpell(CharIndex)
                            End If
                        End If
                    Else
                        If CharTimer(CharIndex).Check(TimersIndex.CastSpell) Then
                            Call BotLanzaSpell(CharIndex)
                        End If
                    End If
                    Call CharTimer(CharIndex).Restart(TimersIndex.WaitH)
                End If
            End If
        End If
        If CharTimer(CharIndex).Check(TimersIndex.GolpeU, False) Then
            If CharTimer(CharIndex).Check(TimersIndex.UseItem, False) Then
                If charlist(CharIndex).Lanzando = 0 And CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                    If CharTimer(CharIndex).Check(TimersIndex.UseItemWithDblClick, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItem)
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItemWithDblClick)
                        Call BotTomaPot(CharIndex)
                    End If
                Else
                    If CharTimer(CharIndex).Check(TimersIndex.UseItemWithU, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItem)
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItemWithU)
                        Call BotTomaPot(CharIndex)
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub PaladinAI(ByVal CharIndex As Integer)
'---------------------------------------------------------------------------------------
' Procedure : PaladinAI
' Author    : Anagrama
' Date      : ???
' Purpose   : Nucleo de la IA del paladin, decide si moverse, atacar o potear.
'---------------------------------------------------------------------------------------
'
    If charlist(CharIndex).Moving = 0 Then
        If charlist(CharIndex).Inmo = 0 Then
            Call MoveBot(CharIndex)
        End If
    End If
    If charlist(CharIndex).MinHP > 0 Then
        If charlist(CharIndex).Inmo = 1 Or MainTimer.Check(TimersIndex.Lan, False) = True Then
            If CharTimer(CharIndex).Check(TimersIndex.WaitH, False) Then
                If (charlist(CharIndex).ComportamientoPotas = 4 And charlist(CharIndex).ComportamientoCombo = 2) Or (charlist(CharIndex).ComportamientoPotas <> 4 And charlist(CharIndex).ComportamientoPotas <> 5) Then
                    If (charlist(CharIndex).ComportamientoCombo <> 2 And charlist(CharIndex).MinMAN >= 400) Or charlist(CharIndex).ComportamientoCombo = 2 Then
                        If Not CharTimer(CharIndex).Check(TimersIndex.Attack, False) Then  'Check if attack interval has finished.
                            If charlist(CharIndex).LastCombo = 1 Then
                                If CharTimer(CharIndex).Check(TimersIndex.AttackCast) Then
                                    Call BotLanzaSpell(CharIndex)
                                    charlist(CharIndex).LastCombo = 2
                                End If
                            End If
                        Else
                            If CharTimer(CharIndex).Check(TimersIndex.CastSpell) Then
                                Call BotLanzaSpell(CharIndex)
                                charlist(CharIndex).LastCombo = 2
                            End If
                        End If
                    Else
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
        If CharTimer(CharIndex).Check(TimersIndex.GolpeU, False) Then
            If CharTimer(CharIndex).Check(TimersIndex.UseItem, False) Then
                If charlist(CharIndex).Lanzando = 0 And CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                    If CharTimer(CharIndex).Check(TimersIndex.UseItemWithDblClick, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItem)
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItemWithDblClick)
                        Call BotTomaPot(CharIndex)
                    End If
                Else
                    If CharTimer(CharIndex).Check(TimersIndex.UseItemWithU, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItem)
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItemWithU)
                        Call BotTomaPot(CharIndex)
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub ClerigoAI(ByVal CharIndex As Integer)
'---------------------------------------------------------------------------------------
' Procedure : ClerigoAI
' Author    : Anagrama
' Date      : ???
' Purpose   : Nucleo de la IA del clerigo, decide si moverse, atacar o potear.
'---------------------------------------------------------------------------------------
'
    If charlist(CharIndex).Moving = 0 Then
        If charlist(CharIndex).Inmo = 0 Then
            Call MoveBot(CharIndex)
        End If
    End If
    If charlist(CharIndex).MinHP > 0 Then
        If charlist(CharIndex).Inmo = 1 Or MainTimer.Check(TimersIndex.Lan, False) = True Then
            If CharTimer(CharIndex).Check(TimersIndex.WaitH, False) Then
                If (charlist(CharIndex).ComportamientoPotas = 4 And charlist(CharIndex).ComportamientoCombo = 3) Or (charlist(CharIndex).ComportamientoPotas <> 4 And charlist(CharIndex).ComportamientoPotas <> 5) Then
                    If charlist(CharIndex).ComportamientoCombo = 3 Or (charlist(CharIndex).MaxMAN > 1200 And charlist(CharIndex).MinMAN >= 1200) Or (charlist(CharIndex).MaxMAN < 1200 And charlist(CharIndex).MinMAN >= 1000) Then
                        If Not CharTimer(CharIndex).Check(TimersIndex.Attack, False) Then  'Check if attack interval has finished.
                            If charlist(CharIndex).LastCombo = 1 Then
                                If CharTimer(CharIndex).Check(TimersIndex.AttackCast) Then
                                    Call BotLanzaSpell(CharIndex)
                                    charlist(CharIndex).LastCombo = 2
                                End If
                            End If
                        Else
                            If CharTimer(CharIndex).Check(TimersIndex.CastSpell) Then
                                Call BotLanzaSpell(CharIndex)
                                charlist(CharIndex).LastCombo = 2
                            End If
                        End If
                    Else
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
        If CharTimer(CharIndex).Check(TimersIndex.GolpeU, False) Then
            If CharTimer(CharIndex).Check(TimersIndex.UseItem, False) Then
                If charlist(CharIndex).Lanzando = 0 And CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                    If CharTimer(CharIndex).Check(TimersIndex.UseItemWithDblClick, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItem)
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItemWithDblClick)
                        Call BotTomaPot(CharIndex)
                    End If
                Else
                    If CharTimer(CharIndex).Check(TimersIndex.UseItemWithU, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItem)
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItemWithU)
                        Call BotTomaPot(CharIndex)
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub DruidaAI(ByVal CharIndex As Integer)
'---------------------------------------------------------------------------------------
' Procedure : DruidaAI
' Author    : Anagrama
' Date      : ???
' Purpose   : Nucleo de la IA del druida, decide si moverse, atacar o potear.
'---------------------------------------------------------------------------------------
'
    If charlist(CharIndex).Moving = 0 Then
        If charlist(CharIndex).Inmo = 0 Then
            Call MoveBot(CharIndex)
        End If
    End If
    If charlist(CharIndex).MinHP > 0 Then
        If charlist(CharIndex).Inmo = 1 Or MainTimer.Check(TimersIndex.Lan, False) = True Then
            If CharTimer(CharIndex).Check(TimersIndex.WaitH, False) Then
                If charlist(CharIndex).ComportamientoCombo = 2 Or charlist(CharIndex).MinMAN >= 1200 Then
                    If CharTimer(CharIndex).Check(TimersIndex.CastSpell) Then
                        Call BotLanzaSpell(CharIndex)
                        charlist(CharIndex).LastCombo = 2
                    End If
                Else
                    If charlist(CharIndex).Lanzando = 1 Then
                        charlist(CharIndex).Lanzando = 0
                        If CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                            Call CharTimer(CharIndex).Restart(TimersIndex.WaitP)
                        End If
                    End If
                End If
            End If
        End If
        If CharTimer(CharIndex).Check(TimersIndex.GolpeU, False) Then
            If CharTimer(CharIndex).Check(TimersIndex.UseItem, False) Then
                If charlist(CharIndex).Lanzando = 0 And CharTimer(CharIndex).Check(TimersIndex.WaitP, False) Then
                    If CharTimer(CharIndex).Check(TimersIndex.UseItemWithDblClick, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItem)
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItemWithDblClick)
                        Call BotTomaPot(CharIndex)
                    End If
                Else
                    If CharTimer(CharIndex).Check(TimersIndex.UseItemWithU, False) Then
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItem)
                        Call CharTimer(CharIndex).Restart(TimersIndex.UseItemWithU)
                        Call BotTomaPot(CharIndex)
                    End If
                End If
            End If
        End If
    End If
End Sub
