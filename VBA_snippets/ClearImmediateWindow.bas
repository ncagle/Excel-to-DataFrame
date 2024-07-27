Attribute VB_Name = "ClearDebug"

' Declare key states and key events
#If VBA7 Then
    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
    Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
#Else
    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#End If

Private Const VK_NUMLOCK As Byte = &H90
Private Const KEYEVENTF_EXTENDEDKEY As Long = &H1
Private Const KEYEVENTF_KEYUP As Long = &H2

' Fill the Immediate debugging window with junk for testing
Sub PrintPeriods()
    Dim i As Integer
    For i = 1 To 50
        Debug.Print (".")
    Next i

End Sub

' Programmatically toggle the Num Lock key
Sub ToggleNumLock()
    keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0

End Sub

' Clear the Immediate debugging window of all outputs
Sub ClearImmediateWindow()
    Dim numLockState As Boolean

    ' Save the current Num Lock state
    numLockState = CBool(GetKeyState(VK_NUMLOCK) And 1)

    ' Activate the Immediate debugging window
    Application.VBE.Windows("Immediate").SetFocus

    ' Select all text in the Immediate debugging window
    SendKeys "^a", True

    ' Delete the selected text
    SendKeys "{DEL}", True

    ' Restore the Num Lock state if necessary
    If CBool(GetKeyState(VK_NUMLOCK) And 1) <> numLockState Then
        ToggleNumLock
    End If

End Sub
