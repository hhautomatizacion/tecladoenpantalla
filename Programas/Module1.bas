Attribute VB_Name = "Module1"
Option Explicit
Public lHwndDestino As Long
Public lHwndVentana As Long
Public lHwndTemporal As Long
Public lHwndAnterior As Long
Public lAnchoPantalla As Long
Public lAltoPantalla As Long
Public sAltoBotonNormal As String
Public sAnchoBotonNormal As String
Public sAltoBotonIcono As String
Public sAnchoBotonIcono As String
Public lAltoBotonNormal As Long
Public lAnchoBotonNormal As Long
Public lAltoBotonIcono As Long
Public lAnchoBotonIcono As Long
Public lAltoBordeVentana As Long
Public lAnchoBordeVentana As Long
Public lAltoBoton As Long
Public bApagandoWindows As Boolean
Public lOffsetX As Long
Public lOffsetY As Long
Public lTiempoAutoOcultar As Long
Public bPresionado As Boolean
Public sNombreFuente As String
Public lTamanioFuente As String
Public bFuenteNegrita As Boolean
Public bPermitirSalir As Boolean
Public bSalirTonello As Boolean
Public bSalir As Boolean
Public bMostrarTiempoRestante As Boolean
'Public lHwndPadre As Long
Public sVersion As String
Public MaxLogSize As Long
Public fTecladoNumeros As Form1
Public fTecladoLetras As Form2
Public sResultado(1) As String
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Declare Function SetMouseHook& Lib "dsmouse.dll" (ByVal hTarget&, ByVal Address&)
Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function PostMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

Public Function Callback&(ByVal msg&, ByVal hWnd&, ByVal X&, ByVal Y&, ByVal HTI&)
    Dim rPos As RECT
    Dim fLetras As Form

    Dim sNombreClase As String
    Dim sTextoBoton As String
    Dim lLong As String
  Select Case msg
  Case 513
    lHwndTemporal = hWnd
    sNombreClase = Space(255)
    If lHwndTemporal <> 0 Then
        lLong = GetClassName(lHwndTemporal, sNombreClase, 255)
        If lLong > 0 Then
            sNombreClase = UCase(Left$(sNombreClase, lLong))
            Select Case sNombreClase
            Case "BUTTON"
                sTextoBoton = Space$(100)
                PostMessage lHwndTemporal, &HD, 101, sTextoBoton
                sTextoBoton = Replace$(sTextoBoton, Chr$(0), "")
                sTextoBoton = Trim$(sTextoBoton)
                dbg "Texto boton: " & sTextoBoton
                Select Case sTextoBoton
                Case "POWER OFF"
                    bSalirTonello = True
                Case Else
                End Select
            Case "THUNDERCOMMANDBUTTON", "COMBOBOX", "BASEBAR", "SYSLISTVIEW32", "THUNDERRT6FORMDC", "THUNDERVWKEY", "THUNDERFILELISTBOX", "THUNDERLISTBOX", "THUNDERVWALARMBOX", "THUNDERVWINDEX", "AFXWND", "THUNDERFRAME", "THUNDERVWVAROUT", "THUNDERVWSTARTUP", "THUNDERVWSHAPE", "THUNDERVWKEYTEXT", "THUNDERVWBAR", "TOOLBARWINDOW32", "COMBOLBOX", "THUNDERFORM"
                sTextoBoton = Space$(100)
                PostMessage lHwndTemporal, &HD, 101, sTextoBoton
                sTextoBoton = Replace$(sTextoBoton, Chr$(0), "")
                sTextoBoton = Trim$(sTextoBoton)
                dbg "Clase: " & sNombreClase & " Texto: " & sTextoBoton
            Case "EDIT", "THUNDERTEXTBOX", "THUNDERVWVARIN", "THUNDERRT6TEXTBOX", "TOVCPICTUREFIELD"
                bSalirTonello = False
                lHwndDestino = lHwndTemporal
                PostMessage lHwndDestino, &HC, 0, Chr$(0)
                If fTecladoLetras.Visible = True Then
                    ColocarVentana fTecladoLetras, lHwndDestino, lAnchoBotonNormal * 7 + lAnchoBotonIcono, lAltoBoton * 4
                Else
                    ColocarVentana fTecladoNumeros, lHwndDestino, lAnchoBotonNormal * 3 + lAnchoBotonIcono, lAltoBoton * 4
                End If
            Case "TTYGRAB", "CONSOLEWINDOWCLASS"
                'todo: msdos window
            Case "THUNDERRT6COMMANDBUTTON"
                
            Case "VBAWINDOW"
                bSalir = True
            Case Else
                dbg "Nombre clase no manejado: " & sNombreClase
            End Select
           
        End If
    End If

  End Select
End Function
Sub ColocarVentana(Ventana As Form, lDest As Long, W As Long, H As Long)
    Dim lRet As Long
    Dim rMy As RECT
    Dim rDest As RECT
    Dim lPosY As Long
    Dim lPosX As Long
    GetWindowRect Ventana.hWnd, rMy
    GetWindowRect lDest, rDest
    
    
    lPosX = rDest.Right
    lPosY = rDest.Top
   
       
        If lPosY + H > lAltoPantalla Then
            lPosY = lAltoPantalla - H
        End If
        If lPosX + W > lAnchoPantalla Then
            lPosY = rDest.Bottom
            lPosX = lAnchoPantalla - W
        End If
    
    Ventana.Timer1.Enabled = False
    Ventana.Timer1.Interval = lTiempoAutoOcultar * 1000
    Ventana.Timer1.Enabled = True


    Ventana.Visible = True
    dbg "Colocar ventana: " & Ventana.Name & " " & W & "x" & H
    lRet = SetWindowPos(Ventana.hWnd, -1, lPosX, lPosY, W, H, &H10 Or &H40)


End Sub
Sub dbg(sMensaje)
    Debug.Print Time$ & vbTab & Len(sMensaje) & vbTab & sMensaje
    WriteLog Format$(Now) & vbTab & Len(sMensaje) & vbTab & sMensaje
End Sub
Sub CargarOpciones()
    Dim lIter As Long
    sAltoBotonNormal = GetSetting("teclado", "opciones", "AltoBotonNormal", "50")
    lAltoBotonNormal = Val(sAltoBotonNormal)
    sAnchoBotonNormal = GetSetting("teclado", "opciones", "AnchoBotonNormal", "50")
    lAnchoBotonNormal = Val(sAnchoBotonNormal)
    sAltoBotonIcono = GetSetting("teclado", "opciones", "AltoBotonIcono", "50")
    lAltoBotonIcono = Val(sAltoBotonIcono)
    sAnchoBotonIcono = GetSetting("teclado", "opciones", "AnchoBotonIcono", "60")
    lAnchoBotonIcono = Val(sAnchoBotonIcono)
    lOffsetX = Val(GetSetting("teclado", "opciones", "OffsetX", "0"))
    lOffsetY = Val(GetSetting("teclado", "opciones", "OffsetY", "35"))
    sNombreFuente = GetSetting("teclado", "opciones", "nombrefuente", "comic sans ms")
    lTamanioFuente = Val(GetSetting("teclado", "opciones", "tamaniofuente", "14"))
    bFuenteNegrita = Val(GetSetting("teclado", "opciones", "fuentenegrita", "1"))
    lTiempoAutoOcultar = Val(GetSetting("teclado", "opciones", "tiempoautoocultar", "8"))
    bPermitirSalir = Val(GetSetting("teclado", "opciones", "permitirsalir", "0"))
    MaxLogSize = Val(GetSetting("teclado", "opciones", "maxlogsize", "2000000"))
    bMostrarTiempoRestante = Val(GetSetting("teclado", "opciones", "mostrartiemporestante", "0"))
    sVersion = GetSetting("teclado", "opciones", "Version", "")
    If lTiempoAutoOcultar > 65 Or lTiempoAutoOcultar < 1 Then lTiempoAutoOcultar = 8
    
    If lAltoBotonIcono > lAltoBotonNormal Then
        lAltoBoton = lAltoBotonIcono
    Else
        lAltoBoton = lAltoBotonNormal
    End If
End Sub
Public Sub WriteLog(sLogEntry As String)
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   Dim sLogFile As String, sLogPath As String, iLogSize As Long
   Dim fso, f
   
On Error GoTo ErrHandler

   'Set the path and filename of the log
   sLogPath = App.Path & "\" & App.EXEName
   sLogFile = sLogPath & ".log"
   
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile(sLogFile, ForAppending, True)
   
   'Get the size of the log to check if it's getting unwieldly
   iLogSize = GetLogSize(sLogFile)
   If iLogSize > MaxLogSize Then
   
        'If too big, back it up to to retain some sort of history
        fso.CopyFile sLogFile, (sLogPath & ".old"), True
        Set f = Nothing
        fso.DeleteFile sLogFile
        'And start with a clean log-file
        Set f = fso.OpenTextFile(sLogFile, ForAppending, True)
        
   End If
    
   'Append the log-entry to the file together with time and date
   f.WriteLine sLogEntry
   
ErrHandler:
    Exit Sub
End Sub

Private Function GetLogSize(filespec As String) As Long
'Returns the size of a file in bytes. If the file does not
'exist, it returns -1.

   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   If (fso.FileExists(filespec)) Then
        Set f = fso.GetFile(filespec)
        GetLogSize = f.Size
   Else
        GetLogSize = -1
   End If
End Function
Sub GuardarOpciones()
    SaveSetting "teclado", "opciones", "NombreFuente", sNombreFuente
    SaveSetting "teclado", "opciones", "TamanioFuente", Format$(lTamanioFuente)
    SaveSetting "teclado", "opciones", "FuenteNegrita", Format$(-bFuenteNegrita)
    SaveSetting "teclado", "opciones", "TiempoAutoOcultar", Format$(lTiempoAutoOcultar)
    SaveSetting "teclado", "opciones", "PermitirSalir", Format$(-bPermitirSalir)
    SaveSetting "teclado", "opciones", "MostrarTiempoRestante", Format$(-bMostrarTiempoRestante)
    SaveSetting "teclado", "opciones", "AltoBotonNormal", sAltoBotonNormal
    SaveSetting "teclado", "opciones", "AnchoBotonNormal", sAnchoBotonNormal
    SaveSetting "teclado", "opciones", "AltoBotonIcono", sAltoBotonIcono
    SaveSetting "teclado", "opciones", "AnchoBotonIcono", sAnchoBotonIcono
    SaveSetting "teclado", "opciones", "maxlogsize", Format$(MaxLogSize)
    SaveSetting "teclado", "opciones", "OffsetX", Format$(lOffsetX)
    SaveSetting "teclado", "opciones", "OffsetY", Format$(lOffsetY)
    SaveSetting "teclado", "opciones", "Version", sVersion
End Sub
Sub Main()
    If App.PrevInstance Then
        dbg "Ya se esta ejecutando el programa"
        bSalir = True
    Else
        lAltoPantalla = GetSystemMetrics(1)
        lAnchoPantalla = GetSystemMetrics(0)
        sResultado(0) = "Fallo"
        sResultado(1) = "Ok"
        CargarOpciones
        If Len(sVersion) = 0 Then
            sVersion = App.Major & "." & App.Minor & "." & App.Revision
            GuardarOpciones
        End If
        bApagandoWindows = False
        dbg "Inicia teclado Version " & sVersion
        Set fTecladoNumeros = New Form1
        Load fTecladoNumeros
        Set fTecladoLetras = New Form2
        Load fTecladoLetras
        dbg "Pantalla: " & lAnchoPantalla & "x" & lAltoPantalla
    End If
    Call SetMouseHook(-1, AddressOf Callback)
    Do
        DoEvents
    Loop Until bSalir
    Call SetMouseHook(0, 0)
    dbg "Termina teclado"
    End
End Sub
