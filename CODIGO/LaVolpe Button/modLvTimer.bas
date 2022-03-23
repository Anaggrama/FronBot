Attribute VB_Name = "modLvTimer"

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (pDest As Any, _
                                       pSource As Any, _
                                       ByVal ByteLen As Long)

Private Declare Function GetProp _
                Lib "user32" _
                Alias "GetPropA" (ByVal hwnd As Long, _
                                  ByVal lpString As String) As Long

Public Function lv_TimerCallBack(ByVal hwnd As Long, _
                                 ByVal Message As Long, _
                                 ByVal wParam As Long, _
                                 ByVal lParam As Long) As Long
    '*************************************************
    'Author: LaVolpe
    'Last modified: 20/05/06
    '*************************************************

    Dim tgtButton As lvButtons_H

    ' when timer was intialized, the button control's hWnd
    ' had property set to the handle of the control itself
    ' and the timer ID was also set as a window property
    CopyMemory tgtButton, GetProp(hwnd, "lv_ClassID"), &H4
    Call tgtButton.TimerUpdate(GetProp(hwnd, "lv_TimerID"))  ' fire the button's event
    CopyMemory tgtButton, 0&, &H4                                    ' erase this instance

End Function

