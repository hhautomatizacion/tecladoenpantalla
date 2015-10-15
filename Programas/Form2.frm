VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Keyboard"
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   4335
   LinkTopic       =   "Form2"
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Height          =   550
      Index           =   0
      Left            =   3600
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Tag             =   "(Esc)"
      Top             =   0
      Width           =   722
   End
   Begin VB.CommandButton Command2 
      Height          =   550
      Index           =   1
      Left            =   3600
      Picture         =   "Form2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   28
      Tag             =   "(BackSpace)"
      Top             =   550
      Width           =   722
   End
   Begin VB.CommandButton Command2 
      Height          =   550
      Index           =   2
      Left            =   3600
      Picture         =   "Form2.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   29
      Tag             =   "(Numeros)"
      Top             =   1100
      Width           =   722
   End
   Begin VB.CommandButton Command2 
      Height          =   550
      Index           =   3
      Left            =   3600
      Picture         =   "Form2.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   30
      Tag             =   "(Enter)"
      Top             =   1650
      Width           =   722
   End
   Begin VB.CommandButton Command1 
      Caption         =   "U"
      Height          =   550
      Index           =   20
      Left            =   3095
      TabIndex        =   20
      Tag             =   "U"
      Top             =   1100
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O"
      Height          =   550
      Index           =   14
      Left            =   0
      TabIndex        =   14
      Tag             =   "O"
      Top             =   1100
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "P"
      Height          =   550
      Index           =   15
      Left            =   516
      TabIndex        =   15
      Tag             =   "P"
      Top             =   1100
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Q"
      Height          =   550
      Index           =   16
      Left            =   1032
      TabIndex        =   16
      Tag             =   "Q"
      Top             =   1100
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "R"
      Height          =   550
      Index           =   17
      Left            =   1548
      TabIndex        =   17
      Tag             =   "R"
      Top             =   1100
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S"
      Height          =   550
      Index           =   18
      Left            =   2064
      TabIndex        =   18
      Tag             =   "S"
      Top             =   1100
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "T"
      Height          =   550
      Index           =   19
      Left            =   2580
      TabIndex        =   19
      Tag             =   "T"
      Top             =   1100
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A"
      Height          =   550
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Tag             =   "A"
      Top             =   0
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "B"
      Height          =   550
      Index           =   1
      Left            =   516
      TabIndex        =   1
      Tag             =   "B"
      Top             =   0
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      Height          =   550
      Index           =   2
      Left            =   1032
      TabIndex        =   2
      Tag             =   "C"
      Top             =   0
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "D"
      Height          =   550
      Index           =   3
      Left            =   1548
      TabIndex        =   3
      Tag             =   "D"
      Top             =   0
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E"
      Height          =   550
      Index           =   4
      Left            =   2064
      TabIndex        =   4
      Tag             =   "E"
      Top             =   0
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F"
      Height          =   550
      Index           =   5
      Left            =   2580
      TabIndex        =   5
      Tag             =   "F"
      Top             =   0
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "G"
      Height          =   550
      Index           =   6
      Left            =   3095
      TabIndex        =   6
      Tag             =   "G"
      Top             =   0
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "H"
      Height          =   550
      Index           =   7
      Left            =   0
      TabIndex        =   7
      Tag             =   "H"
      Top             =   550
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I"
      Height          =   550
      Index           =   8
      Left            =   516
      TabIndex        =   8
      Tag             =   "I"
      Top             =   550
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "J"
      Height          =   550
      Index           =   9
      Left            =   1032
      TabIndex        =   9
      Tag             =   "J"
      Top             =   550
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "K"
      Height          =   550
      Index           =   10
      Left            =   1548
      TabIndex        =   10
      Tag             =   "K"
      Top             =   550
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "L"
      Height          =   550
      Index           =   11
      Left            =   2064
      TabIndex        =   11
      Tag             =   "L"
      Top             =   550
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "M"
      Height          =   550
      Index           =   12
      Left            =   2580
      TabIndex        =   12
      Tag             =   "M"
      Top             =   550
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "N"
      Height          =   550
      Index           =   13
      Left            =   3095
      TabIndex        =   13
      Tag             =   "N"
      Top             =   550
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "V"
      Height          =   550
      Index           =   21
      Left            =   0
      TabIndex        =   21
      Tag             =   "V"
      Top             =   1650
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "W"
      Height          =   550
      Index           =   22
      Left            =   516
      TabIndex        =   22
      Tag             =   "W"
      Top             =   1650
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   550
      Index           =   23
      Left            =   1032
      TabIndex        =   23
      Tag             =   "X"
      Top             =   1650
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Y"
      Height          =   550
      Index           =   24
      Left            =   1548
      TabIndex        =   24
      Tag             =   "Y"
      Top             =   1650
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Z"
      Height          =   550
      Index           =   25
      Left            =   2064
      TabIndex        =   25
      Tag             =   "Z"
      Top             =   1650
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Caption         =   "."
      Height          =   550
      Index           =   26
      Left            =   2580
      TabIndex        =   26
      Tag             =   "(Punto)"
      Top             =   1650
      Width           =   516
   End
   Begin VB.CommandButton Command1 
      Height          =   550
      Index           =   27
      Left            =   3095
      TabIndex        =   31
      Tag             =   "(Espacio)"
      Top             =   1650
      Width           =   516
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1440
      Top             =   960
   End
End
Attribute VB_Name = "Form2"
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
Private Sub Command1_Click(Index As Integer)
    Dim lResultado As Long
    ResetTimer
    Select Case Index
        Case 0
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyA, 0)
        Case 1
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyB, 0)
        Case 2
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyC, 0)
        Case 3
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyD, 0)
        Case 4
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyE, 0)
        Case 5
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyF, 0)
        Case 6
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyG, 0)
        Case 7
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyH, 0)
        Case 8
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyI, 0)
        Case 9
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyJ, 0)
        Case 10
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyK, 0)
        Case 11
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyL, 0)
        Case 12
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyM, 0)
        Case 13
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyN, 0)
        Case 14
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyO, 0)
        Case 15
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyP, 0)
        Case 16
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyQ, 0)
        Case 17
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyR, 0)
        Case 18
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyS, 0)
        Case 19
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyT, 0)
        Case 20
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyU, 0)
        Case 21
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyV, 0)
        Case 22
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyW, 0)
        Case 23
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyX, 0)
        Case 24
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyY, 0)
        Case 25
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyZ, 0)
        Case 26
            lResultado = SendMessage(lHwndDestino, &H102, vbKeyDecimal, 0)
        Case 27
            lResultado = SendMessage(lHwndDestino, &H102, vbKeySpace, 0)
    End Select
    dbg "Envia '" & Command1(Index).Tag & "' a " & lHwndDestino & " (" & sResultado(lResultado) & ")"
    
End Sub

Sub AbrirNumeros()
    'Dim rPos As RECT
    'GetWindowRect lHwndDestino, rPos
    ColocarVentana fTecladoNumeros, lHwndDestino, lAnchoBotonNormal * 3 + lAnchoBotonIcono, lAltoBoton * 4
End Sub


Private Sub Command2_Click(Index As Integer)
    Dim lResultado As Long
    Dim rPos As RECT

    'lTiempoBoton = Timer
    ResetTimer

    Select Case Index
    Case 0
        lResultado = SendMessage(lHwndDestino, &H102, &H1B, 0)
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
        AbrirNumeros
    Case 3
        lResultado = SendMessage(lHwndDestino, &H102, &HD, 0)
        If lResultado = 0 Then
            lResultado = SendMessage(lHwndDestino, &H100, &HD, &H11C0001)
        End If
        Me.Visible = False
    End Select
    dbg "Envia '" & Command2(Index).Tag & "' a " & lHwndDestino & " (" & sResultado(lResultado) & ")"
End Sub



Private Sub Form_Load()
    Dim lIter As Long
    Me.WindowState = vbMaximized
    For lIter = 0 To 27
        Command1(lIter).Font = sNombreFuente
        Command1(lIter).FontSize = lTamanioFuente
        Command1(lIter).FontBold = bFuenteNegrita
        Command1(lIter).Height = lAltoBotonNormal
        Command1(lIter).Width = lAnchoBotonNormal
        Command1(lIter).Top = (lIter \ 7) * lAltoBotonNormal
        Command1(lIter).Left = (lIter Mod 7) * lAnchoBotonNormal
    Next lIter
    For lIter = 0 To 3
        Command2(lIter).Height = lAltoBotonIcono
        Command2(lIter).Width = lAnchoBotonIcono
        Command2(lIter).Top = lIter * lAltoBotonIcono
        Command2(lIter).Left = 7 * lAnchoBotonNormal
    Next lIter

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
    Case vbAppWindows
        bApagandoWindows = True
    Case Else
        dbg "From2 UnloadMode: " & Format$(UnloadMode)
    End Select
    If bPermitirSalir Or bApagandoWindows Or bSalirTonello Then
        bSalir = True
    End If
End Sub
Private Sub Timer1_Timer()
    Me.Visible = False
    Timer1.Enabled = False
End Sub
