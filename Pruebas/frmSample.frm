VERSION 5.00
Begin VB.Form frmSample 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   " MOUSE HOOK SAMPLE"
   ClientHeight    =   1545
   ClientLeft      =   1005
   ClientTop       =   2760
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3105
   Begin VB.CommandButton Command1 
      Caption         =   "Set Mouse Hook"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   360
      TabIndex        =   0
      Top             =   540
      Width           =   2175
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
' Sample for using the "dsmouse.dll"   (c) 2001 by Delphin Software
'******************************************************************
Option Explicit
'******************************************************************
Dim dummy$, IDE As Boolean
'******************************************************************
Private Sub Command1_Click()
Command1.Enabled = False

Call SetMouseHook(hNotepad, 0)       '# Thread Hook to notepad
Call AddMouseWindow(hNotepad, 64, 0) '# Suppress moving
Call AddMouseWindow(hEdit, 12, 0)    '# Suppress right mouse button

dummy = "After setting the hook, you" & vbCrLf & _
        "cannot move Notepad and don't" & vbCrLf & "get a context menu."
Call SendMessageByString(hEdit, WM_SETTEXT, 0, dummy)
End Sub

'******************************************************************
Private Sub Form_Load()
IDE = IsIDE
If IDE Then '# Let VB find the DLL in every case
  ChDrive Left$(App.Path, 1)
  ChDir App.Path
End If
End Sub

'******************************************************************
Private Sub Form_Unload(Cancel As Integer)
Call SetMouseHook(0, 0)
End Sub

'******************************************************************
Private Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function

'******************************************************************
Sub PosiCmd()
With Command1
  .Left = (ScaleWidth - .Width) \ 2
  .Top = (ScaleHeight - .Height) \ 2
End With
End Sub

'******************************************************************

