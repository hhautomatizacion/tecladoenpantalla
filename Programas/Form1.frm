VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Keyboard"
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   147
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   550
      Index           =   11
      Left            =   1002
      TabIndex        =   15
      Tag             =   "-"
      Top             =   1650
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "0"
      Height          =   550
      Index           =   10
      Left            =   501
      TabIndex        =   14
      Tag             =   "0"
      Top             =   1650
      Width           =   500
   End
   Begin VB.CommandButton Command3 
      Height          =   550
      Index           =   3
      Left            =   1500
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "(Enter)"
      Top             =   1650
      Width           =   700
   End
   Begin VB.CommandButton Command3 
      Height          =   550
      Index           =   2
      Left            =   1500
      Picture         =   "Form1.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "(Teclas)"
      Top             =   1100
      Width           =   700
   End
   Begin VB.CommandButton Command3 
      Height          =   550
      Index           =   1
      Left            =   1500
      Picture         =   "Form1.frx":15D6
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "(BackSpace)"
      Top             =   550
      Width           =   700
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   550
      Index           =   9
      Left            =   0
      TabIndex        =   10
      Tag             =   "+"
      Top             =   1650
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "9"
      Height          =   550
      Index           =   8
      Left            =   1002
      TabIndex        =   9
      Tag             =   "9"
      Top             =   1100
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "8"
      Height          =   550
      Index           =   7
      Left            =   501
      TabIndex        =   8
      Tag             =   "8"
      Top             =   1100
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "7"
      Height          =   550
      Index           =   6
      Left            =   0
      TabIndex        =   7
      Tag             =   "7"
      Top             =   1100
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "6"
      Height          =   550
      Index           =   5
      Left            =   1002
      TabIndex        =   6
      Tag             =   "6"
      Top             =   550
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "5"
      Height          =   550
      Index           =   4
      Left            =   501
      TabIndex        =   5
      Tag             =   "5"
      Top             =   550
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "4"
      Height          =   550
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Tag             =   "4"
      Top             =   550
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "3"
      Height          =   550
      Index           =   2
      Left            =   1002
      TabIndex        =   3
      Tag             =   "3"
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   550
      Index           =   1
      Left            =   501
      TabIndex        =   2
      Tag             =   "2"
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton Command3 
      Height          =   550
      Index           =   0
      Left            =   1500
      Picture         =   "Form1.frx":1EA0
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "(Esc)"
      Top             =   0
      Width           =   700
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   720
      Top             =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      Height          =   550
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Tag             =   "1"
      Top             =   0
      Width           =   500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ResetTimer()
    Timer1.Enabled = False
    Timer1.Interval = lTiempoAutoOcultar * 1000
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click(Index As Integer)
    Dim lResultado As Long
    ResetTimer
    Select Case Index
        Case 10
            lResultado = SendMessage(lHwndDestino, &H102, vbKey0, 0)
        Case 0
            lResultado = SendMessage(lHwndDestino, &H102, vbKey1, 0)
        Case 1
            lResultado = SendMessage(lHwndDestino, &H102, vbKey2, 0)
        Case 2
            lResultado = SendMessage(lHwndDestino, &H102, vbKey3, 0)
        Case 3
            lResultado = SendMessage(lHwndDestino, &H102, vbKey4, 0)
        Case 4
            lResultado = SendMessage(lHwndDestino, &H102, vbKey5, 0)
        Case 5
            lResultado = SendMessage(lHwndDestino, &H102, vbKey6, 0)
        Case 6
            lResultado = SendMessage(lHwndDestino, &H102, vbKey7, 0)
        Case 7
            lResultado = SendMessage(lHwndDestino, &H102, vbKey8, 0)
        Case 8
            lResultado = SendMessage(lHwndDestino, &H102, vbKey9, 0)
        Case 9
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyAdd, 0)
        Case 11
            lResultado = SendMessage(lHwndDestino, &H102, vbKeySubtract, 0)
    
    End Select
    dbg "Envia '" & Command2(Index).Tag & "' a " & lHwndDestino & " (" & sResultado(lResultado) & ")"
End Sub

Sub AbrirLetras()
    'Dim rPos As RECT
    'GetWindowRect lHwndDestino, rPos
    ColocarVentana fTecladoLetras, lHwndDestino, lAnchoBotonNormal * 7 + lAnchoBotonIcono, lAltoBoton * 4
End Sub



Private Sub Command3_Click(Index As Integer)
    Dim lResultado As Long

    ResetTimer
    Select Case Index
    Case 0
        lResultado = SendMessage(lHwndDestino, &H102, &H1B, &H0)

        If lResultado = 0 Then
            lResultado = SendMessage(lHwndDestino, &H100, 27, &H11C0001)
            lResultado = SendMessage(lHwndDestino, &H2111, &H40001D4, &H401D4)
            
        End If
        Me.Visible = False
    Case 1
        lResultado = SendMessage(lHwndDestino, &H102, &H8, 0)
        If lResultado = 0 Then

            lResultado = SendMessage(lHwndDestino, &H100, 8, &H11C0001)
            lResultado = SendMessage(lHwndDestino, &H2111, &H40001D4, &H401D4)
        End If
    Case 2
        Me.Visible = False
        AbrirLetras

    Case 3
        lResultado = SendMessage(lHwndDestino, &H102, &HD, 0)
        If lResultado = 0 Then
            lResultado = SendMessage(lHwndDestino, &H100, &HD, &H11C0001)
        End If
        Me.Visible = False
    End Select
    dbg "Envia " & Command3(Index).Tag & " a " & lHwndDestino & " (" & sResultado(lResultado) & ")"
End Sub




Private Sub Form_Load()
    Dim lIter As Long
    Me.WindowState = vbMaximized
    For lIter = 0 To 11
        Command2(lIter).Font = sNombreFuente
        Command2(lIter).FontSize = lTamanioFuente
        Command2(lIter).FontBold = bFuenteNegrita
        Command2(lIter).Height = lAltoBotonNormal
        Command2(lIter).Width = lAnchoBotonNormal
        Command2(lIter).Top = (lIter \ 3) * lAltoBotonNormal
        Command2(lIter).Left = (lIter Mod 3) * lAnchoBotonNormal
    Next lIter
    For lIter = 0 To 3
        Command3(lIter).Height = lAltoBotonIcono
        Command3(lIter).Width = lAnchoBotonIcono
        Command3(lIter).Top = lIter * lAltoBotonIcono
        Command3(lIter).Left = 3 * lAnchoBotonNormal
    Next lIter
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Me.Visible = False
    'CerrarVentanasLetras
    Select Case UnloadMode
    Case vbAppWindows
        bApagandoWindows = True
        
    Case Else
        dbg "Form1 UnloadMode: " & Format$(UnloadMode)
    End Select
    If bPermitirSalir Or bApagandoWindows Or bSalirTonello Then
        bSalir = True
    '    Call SetMouseHook(0, 0)
    End If
End Sub

Private Sub Timer1_Timer()
    Me.Visible = False
    Timer1.Enabled = False
End Sub
