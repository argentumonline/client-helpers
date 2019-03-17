Attribute VB_Name = "Global"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function MyPicture_GetOrSetSingleton(ByVal IsAction As Boolean, ByVal Control As MyPicture) As MyPicture
    Static s_Control As MyPicture
    
    Set MyPicture_GetOrSetSingleton = IIf(IsAction, Control, s_Control)
    
    If (IsAction) Then
        Set s_Control = Control
    End If
End Function


