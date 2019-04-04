Attribute VB_Name = "ProtocolCmdParse"
'Argentum Online
'
'Copyright (C) 2006 Juan Martín Sotuyo Dodero (Maraxus)
'Copyright (C) 2006 Alejandro Santos (AlejoLp)

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

Option Explicit

Public Enum eNumber_Types
    ent_Byte
    ent_Integer
    ent_Long
    ent_Trigger
End Enum

''
' Interpreta, valida y ejecuta el comando ingresado .
'
' @param    RawCommand El comando en version String
' @remarks  None Known.

Public Sub ParseUserCommand(ByVal RawCommand As String)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modification: 16/11/2009
'Interpreta, valida y ejecuta el comando ingresado
'26/03/2009: ZaMa - Flexibilizo la cantidad de parametros de /nene,  /onlinemap y /telep
'16/11/2009: ZaMa - Ahora el /ct admite radio
'***************************************************
    Dim TmpArgos() As String
    
    Dim Comando As String
    Dim ArgumentosAll() As String
    Dim ArgumentosRaw As String
    Dim Argumentos2() As String
    Dim Argumentos3() As String
    Dim Argumentos4() As String
    Dim CantidadArgumentos As Long
    Dim notNullArguments As Boolean
    
    Dim tmpArr() As String
    Dim tmpInt As Integer
    
    ' TmpArgs: Un array de a lo sumo dos elementos,
    ' el primero es el comando (hasta el primer espacio)
    ' y el segundo elemento es el resto. Si no hay argumentos
    ' devuelve un array de un solo elemento
    TmpArgos = Split(RawCommand, " ", 2)
    
    Comando = Trim$(UCase$(TmpArgos(0)))
    
    If UBound(TmpArgos) > 0 Then
        ' El string en crudo que este despues del primer espacio
        ArgumentosRaw = TmpArgos(1)
        
        'veo que los argumentos no sean nulos
        notNullArguments = LenB(Trim$(ArgumentosRaw))
        
        ' Un array separado por blancos, con tantos elementos como
        ' se pueda
        ArgumentosAll = Split(TmpArgos(1), " ")
        
        ' Cantidad de argumentos. En ESTE PUNTO el minimo es 1
        CantidadArgumentos = UBound(ArgumentosAll) + 1
        
        ' Los siguientes arrays tienen A LO SUMO, COMO MAXIMO
        ' 2, 3 y 4 elementos respectivamente. Eso significa
        ' que pueden tener menos, por lo que es imperativo
        ' preguntar por CantidadArgumentos.
        
        Argumentos2 = Split(TmpArgos(1), " ", 2)
        Argumentos3 = Split(TmpArgos(1), " ", 3)
        Argumentos4 = Split(TmpArgos(1), " ", 4)
    Else
        CantidadArgumentos = 0
    End If
    
    ' Sacar cartel APESTA!! (y es ilógico, estás diciendo una pausa/espacio  :rolleyes: )
    If Comando = "" Then Comando = " "
    
    If Left$(Comando, 1) = "/" Then
        ' Comando normal
        
        Select Case Comando
            Case "/ONLINE"
                Call WriteOnline
                
            Case "/SALIR"
                Call WriteQuit

            Case "/PING"
                Call WritePing

        End Select

    Else
        ' Hablar
        Call WriteTalk(RawCommand)
    End If
End Sub

''
' Show a console message.
'
' @param    Message The message to be written.
' @param    red Sets the font red color.
' @param    green Sets the font green color.
' @param    blue Sets the font blue color.
' @param    bold Sets the font bold style.
' @param    italic Sets the font italic style.

Public Sub ShowConsoleMsg(ByVal Message As String, Optional ByVal red As Integer = 255, Optional ByVal green As Integer = 255, Optional ByVal blue As Integer = 255, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/03/07
'
'***************************************************
    Call AddtoRichTextBox(frmMain.RecTxt, Message, red, green, blue, bold, italic)
End Sub

''
' Returns whether the number is correct.
'
' @param    Numero The number to be checked.
' @param    Tipo The acceptable type of number.

Public Function ValidNumber(ByVal Numero As String, ByVal TIPO As eNumber_Types) As Boolean
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/06/07
'
'***************************************************
    Dim Minimo As Long
    Dim Maximo As Long
    
    If Not IsNumeric(Numero) Then _
        Exit Function
    
    Select Case TIPO
        Case eNumber_Types.ent_Byte
            Minimo = 0
            Maximo = 255

        Case eNumber_Types.ent_Integer
            Minimo = -32768
            Maximo = 32767

        Case eNumber_Types.ent_Long
            Minimo = -2147483648#
            Maximo = 2147483647
        
        Case eNumber_Types.ent_Trigger
            Minimo = 0
            Maximo = 6
    End Select
    
    If Val(Numero) >= Minimo And Val(Numero) <= Maximo Then _
        ValidNumber = True
End Function

''
' Returns whether the ip format is correct.
'
' @param    IP The ip to be checked.

Private Function validipv4str(ByVal ip As String) As Boolean
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/06/07
'
'***************************************************
    Dim tmpArr() As String
    
    tmpArr = Split(ip, ".")
    
    If UBound(tmpArr) <> 3 Then _
        Exit Function

    If Not ValidNumber(tmpArr(0), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(3), eNumber_Types.ent_Byte) Then _
        Exit Function
    
    validipv4str = True
End Function

''
' Converts a string into the correct ip format.
'
' @param    IP The ip to be converted.

Private Function str2ipv4l(ByVal ip As String) As Byte()
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/26/07
'Last Modified By: Rapsodius
'Specify Return Type as Array of Bytes
'Otherwise, the default is a Variant or Array of Variants, that slows down
'the function
'***************************************************
    Dim tmpArr() As String
    Dim bArr(3) As Byte
    
    tmpArr = Split(ip, ".")
    
    bArr(0) = CByte(tmpArr(0))
    bArr(1) = CByte(tmpArr(1))
    bArr(2) = CByte(tmpArr(2))
    bArr(3) = CByte(tmpArr(3))

    str2ipv4l = bArr
End Function

''
' Do an Split() in the /AEMAIL in onother way
'
' @param text All the comand without the /aemail
' @return An bidimensional array with user and mail

Private Function AEMAILSplit(ByRef Text As String) As String()
'***************************************************
'Author: Lucas Tavolaro Ortuz (Tavo)
'Useful for AEMAIL BUG FIX
'Last Modification: 07/26/07
'Last Modified By: Rapsodius
'Specify Return Type as Array of Strings
'Otherwise, the default is a Variant or Array of Variants, that slows down
'the function
'***************************************************
    Dim tmpArr(0 To 1) As String
    Dim Pos As Byte
    
    Pos = InStr(1, Text, "-")
    
    If Pos <> 0 Then
        tmpArr(0) = mid$(Text, 1, Pos - 1)
        tmpArr(1) = mid$(Text, Pos + 1)
    Else
        tmpArr(0) = vbNullString
    End If
    
    AEMAILSplit = tmpArr
End Function
