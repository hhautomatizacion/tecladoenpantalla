Attribute VB_Name = "modSample"
'******************************************************************
' Sample for using the "dsmouse.dll"   (c) 2001 by Delphin Software
'******************************************************************
Option Explicit
'******************************************************************
Public hNotepad&, hEdit&
'******************************************************************
Public Const WM_SETTEXT = &HC
Public Const SWP_HIDEWINDOW& = &H80
Public Const SWP_SHOWWINDOW& = &H40
'------------------------------------------------------------------
Declare Function SetMouseHook& Lib "dsmouse" (ByVal hTarget&, ByVal Address&)
Declare Function SetMoveCallback& Lib "dsmouse" (ByVal Callback&)
Declare Function SetDiscard& Lib "dsmouse" (ByVal Discard&)
Declare Function AddMouseWindow& Lib "dsmouse" (ByVal hwnd&, ByVal Discard&, ByVal Thread&)
Declare Function RemoveMouseWindow& Lib "dsmouse" (ByVal hwnd&)
'------------------------------------------------------------------
Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" (ByVal hWndParent&, ByVal hWndChildAfter&, ByVal lpClassName$, ByVal lpWindowName$)
Declare Function IsWindow& Lib "user32" (ByVal hwnd&)
Declare Function SetWindowPos& Lib "user32" (ByVal hwnd&, ByVal hWndInsertAfter&, ByVal X&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal wFlags&)
Declare Function SendMessageByLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam$)
'******************************************************************
Public Function Callback&(ByVal msg&, ByVal hwnd&, ByVal X&, ByVal Y&, ByVal HTI&)
'
End Function

'******************************************************************
Sub Main()
hNotepad = FindWindow("notepad", vbNullString)
If IsWindow(hNotepad) Then
  MsgBox "Notepad is running!" & vbCrLf & "Please close Notepad first!", 64, " Mouse hook sample:"
  Exit Sub
End If
Call Shell("notepad", 0)
Do Until IsWindow(hNotepad)
  hNotepad = FindWindow("notepad", vbNullString)
  DoEvents
Loop
hEdit = FindWindowEx(hNotepad, 0, "edit", vbNullString)
Call SetWindowPos(hNotepad, 0, 50, 50, 300, 200, SWP_SHOWWINDOW)
Load frmSample
Call SetWindowPos(frmSample.hwnd, 0, 50, 250, 300, 100, SWP_HIDEWINDOW)
Call frmSample.PosiCmd
frmSample.Show
End Sub

'******************************************************************

