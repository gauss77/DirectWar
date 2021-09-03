Attribute VB_Name = "Module5"
Option Explicit

Public Sub InitializeDI()
Set DInput = DirectX.DirectInputCreate()
        
If Err.Number <> 0 Then
MsgBox "Error starting Direct Input, please make sure you have DirectX installed", vbApplicationModal
EndGame
End If

Set DIDevice = DInput.CreateDevice("GUID_SysKeyboard")
    
DIDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
DIDevice.SetCooperativeLevel Form1.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE

DIDevice.Acquire

Exit Sub

End Sub

Public Function GetKey()
DIDevice.GetDeviceStateKeyboard DIState
End Function
