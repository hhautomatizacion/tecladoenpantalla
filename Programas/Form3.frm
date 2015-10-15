VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   Caption         =   "Configurar"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form3"
   ScaleHeight     =   5670
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Permitir salir"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Frame Frame5 
      Caption         =   "Presentacion"
      Height          =   1095
      Left            =   0
      TabIndex        =   16
      Top             =   4005
      Width           =   4575
      Begin VB.TextBox Text4 
         Height          =   315
         Index           =   3
         Left            =   3360
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Index           =   2
         Left            =   3360
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Offset Y:"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   25
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Offset X:"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Borde lateral"
         Height          =   255
         Index           =   1
         Left            =   100
         TabIndex        =   23
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Borde superior"
         Height          =   255
         Index           =   0
         Left            =   100
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Timer Timer4 
      Left            =   1560
      Top             =   5640
   End
   Begin VB.Timer Timer3 
      Interval        =   3000
      Left            =   1080
      Top             =   5640
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   3360
      TabIndex        =   14
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   600
      Top             =   5640
   End
   Begin VB.Frame Frame4 
      Caption         =   "Boton con icono"
      Height          =   1095
      Left            =   0
      TabIndex        =   10
      Top             =   2910
      Width           =   4575
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Height          =   495
         Left            =   2760
         TabIndex        =   11
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Alto"
         Height          =   255
         Index           =   1
         Left            =   100
         TabIndex        =   29
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Ancho"
         Height          =   255
         Index           =   0
         Left            =   100
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   120
      Top             =   5640
   End
   Begin VB.Frame Frame3 
      Caption         =   "Boton normal"
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   1815
      Width           =   4575
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Height          =   495
         Left            =   2760
         TabIndex        =   7
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Alto"
         Height          =   255
         Index           =   1
         Left            =   100
         TabIndex        =   27
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Ancho"
         Height          =   255
         Index           =   0
         Left            =   100
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Temporizador"
      Height          =   960
      Left            =   0
      TabIndex        =   3
      Top             =   855
      Width           =   4575
      Begin VB.CommandButton Command5 
         Height          =   495
         Left            =   3480
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar tiempo restante"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "AutoOcultar"
         Height          =   255
         Left            =   100
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de letra"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "Cambiar"
         Height          =   495
         Left            =   3480
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "AUTOMATIZACION H&&H"
         Height          =   480
         Left            =   100
         TabIndex        =   1
         Top             =   200
         Width           =   3135
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    On Error Resume Next
    CommonDialog1.ShowFont
    'MsgBox Err.Number
    If Err.Number = 0 Then
        Label1.FontName = CommonDialog1.FontName
        Label1.FontSize = CommonDialog1.FontSize
        Label1.FontBold = CommonDialog1.FontBold
    Else
        CommonDialog1.FontBold = bFuenteNegrita
        CommonDialog1.FontName = sNombreFuente
        CommonDialog1.FontSize = lTamanioFuente
    End If
    On Error GoTo 0
End Sub

Private Sub Command2_Click()
    MostrarAbout
End Sub

Private Sub Command3_Click()
    MostrarAbout
End Sub

Private Sub Command4_Click()
    sNombreFuente = CommonDialog1.FontName
    lTamanioFuente = CommonDialog1.FontSize
    bFuenteNegrita = CommonDialog1.FontBold
    lTiempoAutoOcultar = Val(Text1.Text)
    bMostrarTiempoRestante = Check1.Value
    sAnchoBotonNormal = Text2(0).Text
    sAltoBotonNormal = Text2(1).Text
    sAnchoBotonIcono = Text3(0).Text
    sAltoBotonIcono = Text3(1).Text
    sAltoBordeVentana = Text4(0).Text
    sAnchoBordeVentana = Text4(1).Text
    lOffsetX = Val(Text4(2).Text)
    lOffsetY = Val(Text4(3).Text)
    bPermitirSalir = Check2.Value
    GuardarOpciones
    CargarOpciones
    Unload Me
End Sub

Private Sub Command5_Click()
    MostrarAbout
End Sub
Private Sub MostrarAbout()
    Dim f As New frmAbout
    f.Show
End Sub
Private Sub Form_Load()
    CargarOpciones
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlCFBoth
    CommonDialog1.FontBold = bFuenteNegrita
    CommonDialog1.FontName = sNombreFuente
    CommonDialog1.FontSize = lTamanioFuente
    Label1.FontName = sNombreFuente
    Label1.FontSize = lTamanioFuente
    Label1.FontBold = bFuenteNegrita
    Text1.Text = Format$(lTiempoAutoOcultar)
    Check1.Value = -bMostrarTiempoRestante
    Text2(0).Text = sAnchoBotonNormal
    Text2(1).Text = sAltoBotonNormal
    Text3(0).Text = sAnchoBotonIcono
    Text3(1).Text = sAltoBotonIcono
    Text4(0).Text = sAltoBordeVentana
    Text4(1).Text = sAnchoBordeVentana
    Text4(2).Text = Format$(lOffsetX)
    Text4(3).Text = Format$(lOffsetY)
    Check2.Value = -bPermitirSalir
    
End Sub

Private Sub Text1_Change()
    Timer3.Enabled = False
    Timer3.Interval = 3000
    Timer3.Enabled = True
End Sub

Private Sub Text2_Change(Index As Integer)
    Timer1.Enabled = False
    Timer1.Interval = 3000
    Timer1.Enabled = True
End Sub

Private Sub Text3_Change(Index As Integer)
    Timer2.Enabled = False
    Timer2.Interval = 3000
    Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Command2.Width = Val(Text2(0).Text) * Screen.TwipsPerPixelX
    Command2.Height = Val(Text2(1).Text) * Screen.TwipsPerPixelY
    Command2.Top = (Frame3.Height / 2) - (Command2.Height / 2)
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    Command3.Width = Val(Text3(0).Text) * Screen.TwipsPerPixelX
    Command3.Height = Val(Text3(1).Text) * Screen.TwipsPerPixelY
    Command3.Top = (Frame4.Height / 2) - (Command3.Height / 2)

End Sub

Private Sub Timer3_Timer()
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer4.Interval = Val(Text1.Text) * 1000
    Timer4.Enabled = True
    
End Sub

Private Sub Timer4_Timer()
    Command5.Visible = Not Command5.Visible
End Sub
