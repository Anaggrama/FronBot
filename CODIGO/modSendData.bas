Attribute VB_Name = "modSendData"

Option Explicit

Public Enum SendTarget
    ToAll = 1
    ToAllButIndex
    ToAllButIndexAndHost
    ToPCArea
    ToDeadArea
End Enum

Public Sub ServerSendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndData As String)

On Error Resume Next
    Dim LoopC As Long
    Dim Map As Integer
    
    Select Case sndRoute
        Case SendTarget.ToPCArea
            Call SendToCharArea(sndIndex, sndData)
            Exit Sub

        Case SendTarget.ToAll
            For LoopC = 1 To LastUser
                If charlist(LoopC).ConnID <> -1 Then
                    If charlist(LoopC).Logged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub

        Case SendTarget.ToAllButIndex
            For LoopC = 1 To LastUser
                If charlist(LoopC).ConnID <> -1 Then
                    If LoopC <> sndIndex Then
                        If charlist(LoopC).Logged Then 'Esta logeado como usuario?
                            Call EnviarDatosASlot(LoopC, sndData)
                        End If
                    End If
                End If
            Next LoopC
            Exit Sub
            
        Case SendTarget.ToAllButIndexAndHost
            For LoopC = 1 To LastUser
                If charlist(LoopC).ConnID <> -1 Then
                    If LoopC <> UserCharIndex Then
                        If LoopC <> sndIndex Then
                            If charlist(LoopC).Logged Then 'Esta logeado como usuario?
                                Call EnviarDatosASlot(LoopC, sndData)
                            End If
                        End If
                    End If
                End If
            Next LoopC
            Exit Sub
            
        Case SendTarget.ToDeadArea
            Call SendToDeadCharArea(sndIndex, sndData)
            Exit Sub
    End Select
End Sub

Private Sub SendToCharArea(ByVal CharIndex As Integer, ByVal sdData As String)

    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers
        If charlist(LoopC).Bot = 0 Then
            If charlist(LoopC).ConnID <> -1 Then
                If EnArea(CharIndex, LoopC) Then
                    Call EnviarDatosASlot(LoopC, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Private Sub SendToDeadCharArea(ByVal CharIndex As Integer, ByVal sdData As String)
    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers
        If charlist(LoopC).Bot = 0 Then
            If charlist(LoopC).MinHP = 0 Then
                If charlist(LoopC).ConnID <> -1 Then
                    If EnArea(CharIndex, LoopC) Then
                        Call EnviarDatosASlot(LoopC, sdData)
                    End If
                End If
            End If
        End If
    Next LoopC
End Sub
