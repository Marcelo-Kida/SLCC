Attribute VB_Name = "basStartScreen"

'Funções comuns para execução do StartScreen

Option Explicit

Public gstrUsuario                          As String

'Variável utilizada para tratamento de erros
Public lngCodigoErroNegocio                 As Long
Private intNumeroSequencialErro             As Integer

'Variaveis de parametrizacao do sistema
Public gblnRegistraTLB                      As Boolean
Public gstrAmbiente                         As String

Public gstrURLWebService                    As String

Public gstrPathA6CC                         As String
Public gstrPathA6                           As String
Public gstrPathA7                           As String
Public gstrPathA8                           As String
Public gstrPathAuditoria                    As String
Public gstrHelpFileA6                       As String
Public gstrHelpFileA7                       As String
Public gstrHelpFileA8                       As String

Public gstrSource                           As String
Public gstrUsuarioRede                      As String
Public gstrEstacaoTrabalho                  As String
Public gstrPrint                            As String

Public glngTimeOut                          As Long

Private Const gstrArquivoINI                As String = "SLCC.INI"

Public Const ERR_USUARIONAOLOGADO           As Long = 18
Public Const ERR_SEMACESSO                  As Long = 35

'PIKA
'Verificar os códigos de transação para acesso
Public Const PER_SISTEMA_A6CC               As String = "A6CCACESS01"
Public Const PER_SISTEMA_A6                 As String = "A6ACESS01"
Public Const PER_SISTEMA_A7                 As String = "A7ACESS01"
Public Const PER_SISTEMA_A8                 As String = "A8ACESS01"

Public Const PER_AUDITORIA_A6CC             As String = "A6CCACESS02"
Public Const PER_AUDITORIA_A6               As String = "A6ACESS02"
Public Const PER_AUDITORIA_A7               As String = "A7ACESS02"
Public Const PER_AUDITORIA_A8               As String = "A8ACESS02"

'--------------------------------------------------------------------

Public Const GWL_USERDATA = (-21&)
Public Const GWL_WNDPROC = (-4&)
Public Const GW_HWNDPREV = 3

Public Const WM_USER = &H400&
Public Const WM_CLOSE = &H10&

Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Const TRAY_CALLBACK = (WM_USER + 101&)
Public Const NIM_ADD = &H0&
Public Const NIM_MODIFY = &H1&
Public Const NIM_DELETE = &H2&
Public Const NIF_MESSAGE = &H1&
Public Const NIF_ICON = &H2&
Public Const NIF_TIP = &H4&

Public Const WM_MOUSEMOVE = &H200&
Public Const WM_LBUTTONDOWN = &H201&
Public Const WM_LBUTTONUP = &H202&
Public Const WM_LBUTTONDBLCLK = &H203&
Public Const WM_RBUTTONDOWN = &H204&
Public Const WM_RBUTTONUP = &H205&
Public Const WM_RBUTTONDBLCLK = &H206&

Public Const BDR_RAISEDOUTER = &H1&
Public Const BDR_RAISEDINNER = &H4&
Public Const BF_LEFT = &H1&             ' Border flags
Public Const BF_TOP = &H2&
Public Const BF_RIGHT = &H4&
Public Const BF_BOTTOM = &H8&
Public Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
Public Const BF_SOFT = &H1000&          ' For softer buttons

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'Adilson V.1.0.5 - 28/10/2003
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'Constantes para o OpenProcess
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16

Private Const MAX_COMPUTERNAME_LENGTH       As Long = 31
Public Const NORMAL_PRIORITY_CLASS = &H20&

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type NOTIFYICONDATA
    cbSize              As Long
    hwnd                As Long
    uID                 As Long
    uFlags              As Long
    uCallbackMessage    As Long
    hIcon               As Long
    szTip               As String * 64
End Type

Public Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Public PrevWndProc                          As Long

Public Function SubWndProc(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim SysTray                                 As ctlSysTray
Dim lngClassAddr                            As Long
    
    Select Case MSG
        Case TRAY_CALLBACK
            
            lngClassAddr = GetWindowLong(hwnd, GWL_USERDATA)
            
            CopyMemory SysTray, lngClassAddr, 4
            
            SysTray.SendEvent lParam, wParam
            
            CopyMemory SysTray, 0&, 4
    End Select
    
    SubWndProc = CallWindowProc(PrevWndProc, hwnd, MSG, wParam, lParam)

End Function

Public Function fgErroLoadXML(ByRef pxmlDocument As MSXML2.DOMDocument40, ByVal psComponente As String, ByVal psClasse As String, ByVal psMetodo As String) As Variant
    

    Err.Raise pxmlDocument.parseError.errorCode, psComponente & " - " & psClasse & " - " & psMetodo, pxmlDocument.parseError.reason
    
End Function


Public Sub fgCenterMe(NameFrm As Form)

Dim intTop                                   As Integer

On Error Resume Next
        
    NameFrm.Left = (Screen.Width - NameFrm.Width) / 2   ' Center form horizontally.
    
    If NameFrm.MDIChild Then
       intTop = (Screen.Height - NameFrm.Height) / 2 - 640
    Else
       intTop = (Screen.Height - NameFrm.Height) / 2 + 200
    End If
    
    If intTop < 0 Then intTop = 0
    NameFrm.Top = intTop ' Center form vertically.
    
End Sub

'Colocar uma janela em evidencia , primero de todaS

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)

Dim lngFlag                                  As Long

On Error GoTo ErrorHandler
    
    If SetOnTop Then
        lngFlag = HWND_TOPMOST
    Else
        lngFlag = HWND_NOTOPMOST
    End If
    
    SetWindowPos myfrm.hwnd, lngFlag, _
                 myfrm.Left / Screen.TwipsPerPixelX, _
                 myfrm.Top / Screen.TwipsPerPixelY, _
                 myfrm.Width / Screen.TwipsPerPixelX, _
                 myfrm.Height / Screen.TwipsPerPixelY, _
                 SWP_NOACTIVATE Or SWP_SHOWWINDOW
    
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError("A6A7A8StartScreen", "basStartScreen", "AlwaysOnTop", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

'Empilamento dos erros de um processamento

Public Sub fgRaiseError(ByVal pstrComponente As String, _
                        ByVal pstrClasse As String, _
                        ByVal pstrMetodo As String, _
                        ByRef plngCodigoErroNegocio As Long, _
               Optional ByRef piNumeroSequencialErro As Integer = 0, _
               Optional ByVal pstrComplemento As String = "", _
               Optional ByRef pblnGravarErro As Boolean = False)
               
Dim strTexto                                As String
Dim lngErrNumber                            As Long
Dim strErrDescription                       As String
Dim strErrSource                            As String
Dim lngErrLastDllError                      As Long
Dim lngErrHelpContext                       As Long
Dim strErrHelpFile                          As String

Dim objDOMErro                              As MSXML2.DOMDocument40
Dim objElement                              As IXMLDOMElement

    If plngCodigoErroNegocio <> 0 Then
        Err.Clear
        On Error GoTo ErrHandler
        lngErrNumber = vbObjectError + 513 + plngCodigoErroNegocio
        strErrSource = pstrComponente
        strErrDescription = "Obter descrição de erro de negócio"
    Else
        lngErrNumber = Err.Number
        strErrDescription = Err.Description
        strErrSource = Err.Source
        lngErrLastDllError = Err.LastDllError
        lngErrHelpContext = Err.HelpContext
        strErrHelpFile = Err.HelpFile
        On Error GoTo ErrHandler
    End If
    
    Set objDOMErro = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not objDOMErro.loadXML(strErrDescription) Then
        fgAppendNode objDOMErro, "", "Erro", ""
        fgAppendNode objDOMErro, "Erro", "Grupo_ErrorInfo", ""
        fgAppendNode objDOMErro, "Grupo_ErrorInfo", "Number", lngErrNumber
        fgAppendNode objDOMErro, "Grupo_ErrorInfo", "Description", strErrDescription
        fgAppendNode objDOMErro, "Erro", "Repet_Origem", ""
    End If
        
    fgAppendNode objDOMErro, "Repet_Origem", "Grupo_Origem", ""
    
        
    Set objElement = objDOMErro.createElement("Origem")
    objElement.Text = pstrComponente & " - " & pstrClasse & " - " & pstrMetodo
    
    fgAppendXML objDOMErro, "Grupo_Origem", objElement.xml
    
    Set objElement = Nothing
    
    Set objElement = objDOMErro.createElement("Complemento")
    objElement.Text = pstrComplemento
    objDOMErro.selectSingleNode("//Repet_Origem/Grupo_Origem[position()=last()]").appendChild objElement
    Set objElement = Nothing
        
    strTexto = objDOMErro.xml
    
    Set objDOMErro = Nothing
    
    Err.Raise lngErrNumber, pstrComponente & " - " & pstrClasse & " - " & pstrMetodo, strTexto

ErrHandler:
    Err.Raise Err.Number, pstrComponente & " - " & pstrClasse & " - " & pstrMetodo, Err.Description, strErrHelpFile, lngErrHelpContext
End Sub

'Modificar o cursor Normal ou Ampulheta

Public Sub fgCursor(Optional pbStatus As Boolean = False)
    
    If pbStatus Then
        Screen.MousePointer = vbHourglass
    Else
        Screen.MousePointer = vbDefault
    End If
    
End Sub

'Obter valores de um arquivo .INI

Private Function flGetPrivateProfileString(ByVal pstrSecao As String, _
                                           ByVal pstrChave As String, _
                                           ByVal pstrNomeArquivo As String) As String

Dim lngReturnCode                           As Long
Dim strValorKey                             As String

On Error GoTo ErrHandler
    
    If Dir(pstrNomeArquivo) = vbNullString Then
        lngCodigoErroNegocio = 0
        Err.Raise 513, , "Arquivo de Configuração não encontrado."
    Else
    
        strValorKey = String(1000, Chr(0))
        lngReturnCode = GetPrivateProfileString(pstrSecao, pstrChave, "", strValorKey, Len(strValorKey), pstrNomeArquivo)
    
        If lngReturnCode <> 0 Then
            flGetPrivateProfileString = Mid(strValorKey, 1, InStr(1, strValorKey, Chr(0)) - 1)
        End If
    End If
    
    Exit Function
ErrHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basStartScreen", "flGetPrivateProfileString", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Obter os valores dos parametros

Public Sub fgObterConfiguracoes()

Dim strCommandLine                           As String
Dim strParametros                            As Variant

On Error GoTo ErrorHandler
    
    strCommandLine = Command()
    strParametros = Split(strCommandLine, ";")
    
    gstrAmbiente = strParametros(LBound(strParametros))
    #If EnableSoap = 1 Then
        gblnRegistraTLB = False
    #Else
        gblnRegistraTLB = strParametros(LBound(strParametros) + 1)
    #End If
    
    gstrSource = flGetPrivateProfileString(gstrAmbiente, "Source", App.Path & "\" & gstrArquivoINI)
    
    gstrPathA6CC = flGetPrivateProfileString(gstrAmbiente, "Path A6CC", App.Path & "\" & gstrArquivoINI)
    gstrPathA6 = flGetPrivateProfileString(gstrAmbiente, "Path A6", App.Path & "\" & gstrArquivoINI)
    gstrPathA7 = flGetPrivateProfileString(gstrAmbiente, "Path A7", App.Path & "\" & gstrArquivoINI)
    gstrPathA8 = flGetPrivateProfileString(gstrAmbiente, "Path A8", App.Path & "\" & gstrArquivoINI)
    
    gstrHelpFileA6 = flGetPrivateProfileString(gstrAmbiente, "HelpFile A6", App.Path & "\" & gstrArquivoINI)
    gstrHelpFileA7 = flGetPrivateProfileString(gstrAmbiente, "HelpFile A7", App.Path & "\" & gstrArquivoINI)
    gstrHelpFileA8 = flGetPrivateProfileString(gstrAmbiente, "HelpFile A8", App.Path & "\" & gstrArquivoINI)
    
    gstrURLWebService = flGetPrivateProfileString(gstrAmbiente, "Url Web Service", App.Path & "\" & gstrArquivoINI)
    glngTimeOut = CLng(flGetPrivateProfileString(gstrAmbiente, "Timeout", App.Path & "\" & gstrArquivoINI))
    
    gstrPrint = flGetPrivateProfileString(gstrAmbiente, "Print", App.Path & "\" & gstrArquivoINI)
    
    gstrPathAuditoria = flGetPrivateProfileString(gstrAmbiente, "Path Auditoria", App.Path & "\" & gstrArquivoINI)
    
    Exit Sub
    
ErrorHandler:
    
    Err.Raise vbObjectError + 266, "strParametros", "Parâmetros Inválidos- Command Line"
    
End Sub

'Ativa a aplicação

Public Function fgActiveApplication(AppName As String) As Boolean

Dim lngPrevHwnd                             As Long
Dim lngResult                               As Long

    lngPrevHwnd = FindWindow("ThunderRT6Main", AppName)
    
    If lngPrevHwnd > 0 Then
        'Get handle to previous window.
        lngPrevHwnd = GetWindow(lngPrevHwnd, GW_HWNDPREV)
        
        'Restore the program.
        lngResult = OpenIcon(lngPrevHwnd)
        
        'Activate the application.
        lngResult = SetForegroundWindow(lngPrevHwnd)
        
        fgActiveApplication = True
    Else
        fgActiveApplication = False
    End If

End Function

'Instanciar um objeto ( A6A7A8Miu.dll ou MSSOAPLib30.SoapClient30)

Public Function fgCriarObjetoMIU(ByVal pstrNomeClasse As String) As Object

Dim strWSDL                                 As String
Dim strServico                              As String
Dim strPorta                                As String

Dim objSoapClient                           As MSSOAPLib30.SoapClient30

On Error GoTo ErrorHandler

    #If EnableSoap = 1 Then
        strServico = UCase$(Split(pstrNomeClasse, ".")(0))
        strWSDL = gstrURLWebService & "/" & strServico & ".WSDL"
        strPorta = Split(pstrNomeClasse, ".")(1) & "SoapPort"
        
        Set objSoapClient = CreateObject("MSSOAP.SoapClient30")
        Call objSoapClient.MSSoapInit(strWSDL, strServico, strPorta)
        objSoapClient.ConnectorProperty("Timeout") = glngTimeOut * 1000
        
        Set fgCriarObjetoMIU = objSoapClient
    #Else
        Set fgCriarObjetoMIU = CreateObject(pstrNomeClasse)
    #End If

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, "basA8LQS", "fgCriarObjetoMIU", 0
   
End Function

'Inicializa uma das aplicações A6,A7,A8,Trilha de auditoria

Public Function fgShowApplication(ByVal pstrAppPath As String, _
                                  ByVal pstrParametros As String, _
                                  ByVal pstrPermissao As String, _
                                  ByRef pudtProcessInfo As PROCESS_INFORMATION) As Boolean

#If EnableSoap = 1 Then
    Dim objSeguranca                        As MSSOAPLib30.SoapClient30
#Else
    Dim objSeguranca                        As A6A7A8Miu.clsPerfil
#End If

Dim udtProcessInfo                          As PROCESS_INFORMATION
Dim udtStartupInfo                          As STARTUPINFO
Dim strNull                                 As String
Dim lngSuccess                              As Long

On Error GoTo ErrorHandler

    Set objSeguranca = fgCriarObjetoMIU("A6A7A8Miu.clsPerfil")
            
    If objSeguranca.VerificarPermissao(gstrUsuario, pstrPermissao) Then
        
        If Dir(pstrAppPath) = vbNullString Then
            Err.Raise 520, "fgShowApplication", "Módulo executável não encontrado"
        Else
            udtStartupInfo.cb = Len(udtStartupInfo)
            
            fgLockWindow frmStartScreen.hwnd
            
            lngSuccess = CreateProcess(strNull, _
                                       pstrAppPath & Space(1) & pstrParametros, _
                                       ByVal 0&, _
                                       ByVal 0&, _
                                       1&, _
                                       NORMAL_PRIORITY_CLASS, _
                                       ByVal 0&, _
                                       strNull, _
                                       udtStartupInfo, _
                                       udtProcessInfo)
    
            pudtProcessInfo = udtProcessInfo
            
            fgLockWindow 0
        End If
    Else
        Err.Raise 525, "fgShowApplication", "Acesso não permitido"
    End If
    
    Set objSeguranca = Nothing

    Exit Function
ErrorHandler:
    
    fgLockWindow 0
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError("A6A7A8StartScreen", "basStartScreen", "fgShowApplication Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
    
Public Sub Main()

#If EnableSoap = 1 Then
    Dim objUsuario                          As MSSOAPLib30.SoapClient30
#Else
    Dim objUsuario                          As A6A7A8Miu.clsUsuario
#End If

Dim objDOMDOC                               As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    If App.PrevInstance Then
        fgSetFocus "SLCC"
        End
        Exit Sub
    End If
            
    flLimpaRegistry
    fgObterConfiguracoes
        
    'Obtem o usuário de rede e o nome da estação client,
    'esses dados serão usados para log das operações efetuadas pelo usuário
    gstrUsuarioRede = flObterUsuarioRede()
    gstrEstacaoTrabalho = flObterEstacaoTrabalho()

    Set objUsuario = fgCriarObjetoMIU("A6A7A8Miu.clsUsuario")
    '   será utilizado como sistema de controle de acesso o MBS.
    '   Como o MBS utiliza-se do mesmo usuário de rede, não apresenta o form de logon
    
    'gstrUsuario = objUsuario.ObterUsuario
    
    gstrUsuario = flObterUsuarioRede()
    
    frmStartScreen.StatusBar.ToolTipText = gstrUsuario & " - " & gstrEstacaoTrabalho
    
    objUsuario.LogOn gstrUsuario, gstrUsuarioRede, gstrEstacaoTrabalho, False
    
    Set objUsuario = Nothing
    
    frmStartScreen.Show
        
    Exit Sub
    
ErrorHandler:
    
    Set objUsuario = Nothing
    
    Set objDOMDOC = CreateObject("MSXML2.DOMDocument.4.0")
        
    objDOMDOC.loadXML Err.Description
    
    If objDOMDOC.parseError.errorCode <> 0 Then
        MsgBox Err.Description & vbCrLf & _
               vbCrLf & _
               "O Sistema será fechado."
    Else
        MsgBox objDOMDOC.selectSingleNode("Erro/Grupo_ErrorInfo/Number").Text & "-" & objDOMDOC.selectSingleNode("Erro/Grupo_ErrorInfo/Description").Text & vbCrLf & _
               vbCrLf & _
               "O Sistema será fechado."
    End If
    
    Set objDOMDOC = Nothing
    
    End

End Sub

'Obter nome da estação de trabalho local

Private Function flObterEstacaoTrabalho() As String

Dim strEstacao                               As String
Dim lngLen                                   As Long

On Error GoTo ErrorHandler
    
    lngLen = MAX_COMPUTERNAME_LENGTH + 1
    strEstacao = String(lngLen, "X")
    
    GetComputerName strEstacao, lngLen
    strEstacao = Left(strEstacao, lngLen)
    flObterEstacaoTrabalho = strEstacao

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "modStartScreen", "flObterEstacaoTrabalho Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Obter ID do usuário local

Private Function flObterUsuarioRede() As String

Dim strUserName                              As String
Dim lngLen                                   As Long

On Error GoTo ErrorHandler

    lngLen = 100
    strUserName = String(lngLen, Chr$(0))
    GetUserName strUserName, lngLen
    strUserName = Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)
    flObterUsuarioRede = UCase$(Left$(strUserName, 8))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "modStartScreen", "flObterUsuarioRede Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Finalizar um processo

Public Sub fgTerminateProcess(pudtProcessInfo As PROCESS_INFORMATION)

Dim lngReturn                                 As Long

On Error GoTo ErrorHandler
    
                
    lngReturn = DestroyWindow(pudtProcessInfo.hProcess)
    lngReturn = TerminateProcess(pudtProcessInfo.hProcess, 0&)
    lngReturn = CloseHandle(pudtProcessInfo.hThread)
    lngReturn = CloseHandle(pudtProcessInfo.hProcess)

    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basStartScreen", "fgTerminateProcess Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

'Verifica se é um processo do Windows

Public Function fgIsProcess(pudtProcessInfo As PROCESS_INFORMATION) As Boolean

On Error GoTo ErrorHandler

    fgIsProcess = (OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, -1, pudtProcessInfo.dwProcessId) <> 0)

    Exit Function
    
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basStartScreen", "fgIsProcess Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Bloquer / desbloquear uma janela

Public Function fgLockWindow(Optional ByVal plWHnd As Long)
    
    LockWindowUpdate plWHnd

End Function

'Setfocus em uma janela

Public Function fgSetFocus(ByVal strTitle As String) As Boolean

Dim lngHwd                                  As Long
    
    lngHwd = ShowWindow(fgFindWindow(strTitle), IIf(strTitle = "SLCC", vbNormalFocus, 0))
    lngHwd = ShowWindow(fgFindWindow(strTitle), IIf(strTitle = "SLCC", vbNormalFocus, 3))

End Function

'Procura janela aberta

Public Function fgFindWindow(ByVal strTitle As String) As Long

    fgFindWindow = FindWindow(vbNullString, strTitle)

End Function

'Limpar registry referentes ao SLCC

Private Sub flLimpaRegistry()

Dim strDataUltimaExecucao                   As String
'
'Existe a necessidade deste Resume Next para caso ele não encontre a seção.
'
On Error Resume Next

    strDataUltimaExecucao = GetSetting("A6A7A8", "DataUltimaExecucao", "Settings")
    
    If strDataUltimaExecucao <> Format$(Date, "YYYYMMDD") Then
       SaveSetting "A6A7A8", "DataUltimaExecucao", "Settings", Format$(Now, "YYYYMMDD")
       DeleteSetting "A6SBR\Form Filtro\"
       DeleteSetting "A7\Alerta\"
       DeleteSetting "A8LQS\Form Filtro\"
    End If
    
End Sub
