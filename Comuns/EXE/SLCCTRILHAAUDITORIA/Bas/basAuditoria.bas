Attribute VB_Name = "basAuditoria"
Option Explicit

Public gstrSource                           As String
Public gstrAmbiente                         As String
Public gstrUsuario                          As String
Public gblnAcessoOnLine                     As Boolean
Public gblnRegistraTLB                      As Boolean
Public gstrOwnerSLCC                        As String

Public gstrURLWebService                    As String
Public glngTimeOut                          As Long

Public Const gdtmDataVazia                  As String = "00:00:00"

Private Const MAX_COMPUTERNAME_LENGTH       As Long = 31

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFilename As String) As Long

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

Public Sub Main()

Dim strParametros()                         As String

On Error GoTo ErrMain
    
    If App.PrevInstance Then
        MsgBox "Já existe uma aplicação aberta.", vbCritical
        Exit Sub
    End If
    
    strParametros = Split(Command(), ";")
    
    If strParametros(0) <> "Desenv1" Then
        flObterConfiguracaoCommandLine
    Else
        gstrURLWebService = strParametros(3)
        glngTimeOut = strParametros(4)
    End If
    
    mdiTrilhaAuditoria.Show
    
    DoEvents
    
    Exit Sub
ErrMain:
    
    MsgBox "Erro-> " & Err.Description & vbCrLf & "O Sistema será finalizado "
    
    End
    
End Sub

'Centralizar um form

Sub fgCenterMe(NameFrm As Form)
   
    Dim iTop As Integer
    
    On Error Resume Next
        
    NameFrm.Left = (Screen.Width - NameFrm.Width) / 2   ' Center form horizontally.
    
    If NameFrm.MDIChild Then
       iTop = (Screen.Height - NameFrm.Height) / 2 - 640
    Else
       iTop = (Screen.Height - NameFrm.Height) / 2 + 200
    End If
    
    If iTop < 0 Then iTop = 0
    
    NameFrm.Top = iTop ' Center form vertically.
    
End Sub

'Setar o tipo de cursor (Normal ou Ampulheta)

Public Sub fgCursor(Optional pbStatus As Boolean = False)
    
    If pbStatus Then
        Screen.MousePointer = vbHourglass
    Else
        Screen.MousePointer = vbDefault
    End If
    
End Sub

'Obter os parâmetros de entrada

Private Sub flObterConfiguracaoCommandLine()

Dim strCommandLine                           As String
Dim strParametros                            As Variant

On Error GoTo ErrorHandler
    
    strCommandLine = Command()
    strParametros = Split(strCommandLine, ";")
    
    gstrAmbiente = strParametros(LBound(strParametros))
    gstrUsuario = strParametros(LBound(strParametros) + 1)
    gstrOwnerSLCC = UCase(strParametros(LBound(strParametros) + 2))
    
    gstrURLWebService = strParametros(LBound(strParametros) + 3)
    glngTimeOut = strParametros(LBound(strParametros) + 4)
    
    Exit Sub
ErrorHandler:
    
    Err.Raise vbObjectError + 266, "strParametros", "Parâmetros Inválidos- Command Line"
    
End Sub

'Obter o nome da estação de trabalho local

Public Function fgObterEstacaoTrabalho() As String

Dim strEstacao                               As String
Dim lngLen                                   As Long

On Error GoTo ErrorHandler
    
    lngLen = MAX_COMPUTERNAME_LENGTH + 1
    strEstacao = String(lngLen, "X")
    
    GetComputerName strEstacao, lngLen
    strEstacao = Left(strEstacao, lngLen)
    fgObterEstacaoTrabalho = strEstacao

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basAuditoria", "fgObterEstacaoTrabalho Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Obter ID do usuário Local

Public Function fgObterUsuarioRede() As String

Dim lsUserName                              As String
Dim lngLen                                  As Long

On Error GoTo ErrorHandler
    
    lngLen = 100
    lsUserName = String(lngLen, Chr$(0))
    GetUserName lsUserName, lngLen
    fgObterUsuarioRede = Left$(lsUserName, InStr(lsUserName, Chr$(0)) - 1)

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basAuditoria", "fgObterUsuarioRede Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Adicionar um erro ao xml de lista de erro

Public Sub fgRaiseError(ByVal strComponente As String, _
                        ByVal strClasse As String, _
                        ByVal strMetodo As String, _
                        ByRef lngCodigoErroNegocio As Long, _
               Optional ByRef intNumeroSequencialErro As Integer = 0, _
               Optional ByVal strComplemento As String = "", _
               Optional ByRef blnGravarErro As Boolean = False)

Dim strTexto                                As String
Dim ErrNumber                               As Long
Dim ErrDescription                          As String
Dim ErrSource                               As String
Dim ErrLastDllError                         As Long
Dim ErrHelpContext                          As Long
Dim ErrHelpFile                             As String

Dim objDOMErro                              As MSXML2.DOMDocument40
Dim objElement                              As IXMLDOMElement


    If lngCodigoErroNegocio <> 0 Then
        Err.Clear
        On Error GoTo ErrHandler
        ErrNumber = vbObjectError + 513 + lngCodigoErroNegocio
        ErrSource = strComponente
        ErrDescription = "Obter descrição de erro de negócio"
    Else
        ErrNumber = Err.Number
        ErrDescription = Err.Description
        ErrSource = Err.Source
        ErrLastDllError = Err.LastDllError
        ErrHelpContext = Err.HelpContext
        ErrHelpFile = Err.HelpFile
        On Error GoTo ErrHandler
    End If
    
    Set objDOMErro = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not objDOMErro.loadXML(ErrDescription) Then
        fgAppendNode objDOMErro, "", "Erro", ""
        fgAppendNode objDOMErro, "Erro", "Grupo_ErrorInfo", ""
        fgAppendNode objDOMErro, "Grupo_ErrorInfo", "Number", ErrNumber
        fgAppendNode objDOMErro, "Grupo_ErrorInfo", "Description", ErrDescription
        fgAppendNode objDOMErro, "Erro", "Repet_Origem", ""
    End If
        
    fgAppendNode objDOMErro, "Repet_Origem", "Grupo_Origem", ""
            
    Set objElement = objDOMErro.createElement("Origem")
    objElement.Text = strComponente & " - " & strClasse & " - " & strMetodo
    
    objDOMErro.selectSingleNode("//Repet_Origem/Grupo_Origem[position()=last()]").appendChild objElement
        
    Set objElement = Nothing
    
    Set objElement = objDOMErro.createElement("Complemento")
    objElement.Text = strComplemento
    objDOMErro.selectSingleNode("//Repet_Origem/Grupo_Origem[position()=last()]").appendChild objElement
    Set objElement = Nothing
        
    strTexto = objDOMErro.xml
    
    Set objDOMErro = Nothing
    
    Err.Raise ErrNumber, strComponente & " - " & strClasse & " - " & strMetodo, strTexto

ErrHandler:
    Err.Raise Err.Number, strComponente & " - " & strClasse & " - " & strMetodo, Err.Description, ErrHelpFile, ErrHelpContext
End Sub
