'Handle 64-bit and 32-bit Office
#If VBA7 Then
  Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, _
    ByVal dwBytes As LongPtr) As Long
  Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
  Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As Long
  Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
  Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
    ByVal lpString2 As Any) As Long
  Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat _
    As LongPtr, ByVal hMem As LongPtr) As Long
#Else
  Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
  Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
  Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
    ByVal dwBytes As Long) As Long
  Declare Function CloseClipboard Lib "User32" () As Long
  Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
  Declare Function EmptyClipboard Lib "User32" () As Long
  Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
    ByVal lpString2 As Any) As Long
  Declare Function SetClipboardData Lib "User32" (ByVal wFormat _
    As Long, ByVal hMem As Long) As Long
#End If

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096