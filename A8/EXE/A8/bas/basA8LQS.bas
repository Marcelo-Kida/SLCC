Attribute VB_Name = "basA8LQS"

' Este componente tem como objetivo agrupar métodos utilizados na camada de interface do sistema A8.

Option Explicit

Public gstrSource                           As String
Public gstrAmbiente                         As String
Public gstrUsuario                          As String
Public gblnAcessoOnLine                     As Boolean
Public gblnRegistraTLB                      As Boolean
Public gblnEnableSOAP                       As Boolean
Public gintTipoBackoffice                   As Integer

'Informações sobre a versão dos componentes envolvidos
Public gstrVersao                           As String

Public gstrURLWebService                    As String
Public glngTimeOut                          As Long
Public gstrHelpFile                         As String
Public gstrPrint                            As String

'Utilizada para controlar o perfil MANUTENÇÃO
'   Habilita botões de manutenção nos formulários
Public gblnPerfilManutencao                 As Boolean

'Utilizada como timer da Disponibilização de Alertas
Public glngContaMinutosAlerta               As Long
Public glngTempoAlerta                      As Long

'Utilizada como timer para verificar se existe algum sistema em contingência
Public glngContaMinutosContingencia         As Long
Public glngTempoContingencia                As Long

Public strHoraInicioVerificacao             As String
Public strHoraFimVerificacao                As String

'------------------ API ----------------------------------------
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFilename As String) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'---------------------------------------------------------------

'------------------ Constants ----------------------------------
Public Const gsMaskInt                      As String = "##,###,###,###,###,##0;[vbRed](##,###,###,###,###,##0)"
Public Const gsMaskDec                      As String = "##,###,###,###,###,##0.00;[vbRed](##,###,###,###,###,##0.00)"
Public gstrMascaraDataDtp                   As String
Public gstrSeparadorData                    As String

'Caracteres especiais
Public Const CAR_APOSTROFE                  As Long = 39       ' Caracter '
Public Const CAR_ABRE_CHAVE                 As Long = 123      ' Caracter {
Public Const CAR_FECHA_CHAVE                As Long = 125      ' Caracter }
Public Const CAR_ENTER                      As Long = 13       ' Caracter {Enter}
Public Const CAR_LINEFEED                   As Long = 10       ' Caracter {LF}
Public Const CAR_SUBST                      As Long = 127      ' Caracter que substituirá o {Enter}
Public Const CAR_ASPAS                      As Long = 34       ' "
Public Const CAR_PORCENTO                   As Long = 37       ' %
Public Const CAR_INTERROGACAO               As Long = 63       ' ?
Public Const CAR_CASP1                      As Long = 96       ' `
Public Const CAR_CASP2                      As Long = 180      ' ´
Public Const CAR_ASPAS1                     As Long = 145      ' ‘
Public Const CAR_ASPAS2                     As Long = 146      ' ’
Public Const CAR_BARRA                      As Long = 47       ' /
Public Const CAR_PONTO                      As Long = 46       ' .
Public Const CAR_HIFEN                      As Long = 45       ' -

'Tipos de operação para XML DOMs
Public Const gstrOperIncluir                As String = "Incluir"
Public Const gstrOperExcluir                As String = "Excluir"
Public Const gstrOperAlterar                As String = "Alterar"
Public Const gstrOperNone                   As String = "None"
Public Const gstrOperConsultar              As String = "Consultar"
Public Const gstrOperLer                    As String = "Ler"
Public Const gstrOperLerTodos               As String = "LerTodos"
'---------------------------------------------------------------

'------------------ Variaveis LQS ------------------------------
Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

'---------------------------------------------------------------
Public lngErrNumber                          As Long

'---------------Constantes para uso de ctlSysTray---------------------------
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const BDR_RAISEDOUTER = &H1&
Public Const BDR_RAISEDINNER = &H4&
Public Const BF_LEFT = &H1&             ' Border flags
Public Const BF_TOP = &H2&
Public Const BF_RIGHT = &H4&
Public Const BF_BOTTOM = &H8&
Public Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
Public Const BF_SOFT = &H1000&          ' For softer buttons

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

Public Const WM_MOUSEMOVE = &H200&
Public Const WM_LBUTTONDOWN = &H201&
Public Const WM_LBUTTONUP = &H202&
Public Const WM_LBUTTONDBLCLK = &H203&
Public Const WM_RBUTTONDOWN = &H204&
Public Const WM_RBUTTONUP = &H205&
Public Const WM_RBUTTONDBLCLK = &H206&


Public Const TRAY_CALLBACK = (WM_USER + 101&)
Public Const NIM_ADD = &H0&
Public Const NIM_MODIFY = &H1&
Public Const NIM_DELETE = &H2&
Public Const NIF_MESSAGE = &H1&
Public Const NIF_ICON = &H2&
Public Const NIF_TIP = &H4&

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

Public gintRowPositionAnt                   As Integer
Public gintIndexWorksheets                  As Integer

Public gxmlCombosFiltro                     As MSXML2.DOMDocument40
Public gblnExibirTipoBackOffice             As Boolean

'Verifica se está no ambiente de desenvolvimento (pelo Command$)
Private blnDesenv       As Boolean

'Função utilizada pelo ctlSystray
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

'-----------------------------------------------------------------------
Public Function fgIconeApp() As StdPicture
    Set fgIconeApp = mdiLQS.Icon
End Function

'Obter detalhes da versão
Public Function fgObterDetalhesVersoes() As String

#If EnableSoap = 1 Then
    Dim objVersao           As MSSOAPLib30.SoapClient30
#Else
    Dim objVersao           As A8MIU.clsVersao
#End If

Dim xmlVersoes              As MSXML2.DOMDocument40
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    If gstrVersao <> vbNullString Then
        fgObterDetalhesVersoes = gstrVersao
        Exit Function
    End If

    Set xmlVersoes = CreateObject("MSXML2.DOMDocument.4.0")
    fgAppendNode xmlVersoes, "", "Componentes", ""
    flAdicionaDadosVersao xmlVersoes

    Set objVersao = fgCriarObjetoMIU("A8MIU.clsVersao")
    xmlVersoes.loadXML objVersao.ObterVersoesComponentes(xmlVersoes.xml, _
                                                         vntCodErro, _
                                                         vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    gstrVersao = xmlVersoes.xml
    fgObterDetalhesVersoes = gstrVersao

    Set objVersao = Nothing
    Set xmlVersoes = Nothing

Exit Function
ErrorHandler:

    Set objVersao = Nothing
    Set xmlVersoes = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, "basA8LQS", "fgObterDetalhesVersoes", 0
End Function

'Adicionar os dados da versão
Private Sub flAdicionaDadosVersao(ByRef xmlVersao As MSXML2.DOMDocument40)

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objDomNodePropriedade                   As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler
    
    Set objDomNode = xmlVersao.createElement("Componente")
    
    Set objDomNodePropriedade = xmlVersao.createElement("Title")
    objDomNodePropriedade.Text = App.Title
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("Tipo")
    objDomNodePropriedade.Text = fgObterTipoComponente
    objDomNode.appendChild objDomNodePropriedade

    Set objDomNodePropriedade = xmlVersao.createElement("Major")
    objDomNodePropriedade.Text = App.Major
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("Minor")
    objDomNodePropriedade.Text = App.Minor
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("Revision")
    objDomNodePropriedade.Text = App.Revision
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("FileDescription")
    objDomNodePropriedade.Text = App.FileDescription
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("Date")
    objDomNodePropriedade.Text = flDataComponente
    objDomNode.appendChild objDomNodePropriedade
    
    xmlVersao.documentElement.appendChild objDomNode
    
    Set objDomNode = Nothing
    Set objDomNodePropriedade = Nothing
    
Exit Sub
ErrorHandler:

    Set objDomNode = Nothing
    Set objDomNodePropriedade = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "flAdicionaDadosVersao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Obter a data dos componentes
Private Function flDataComponente() As String

    On Error GoTo ErrorHandler
    
    flDataComponente = fgDtHr_To_Xml(FileDateTime(App.Path & "\" & App.EXEName & ".exe"))

Exit Function
ErrorHandler:
    
    flDataComponente = fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora))
    
End Function

'Adicionar dias uteis a uma data
Public Function fgAdicionarDiasUteis(ByVal pdatData As Date, _
                                     ByVal pintQtdeDias As Integer, _
                                     ByVal plngMovimento As enumPaginacao) As Date

#If EnableSoap = 1 Then
    Dim objA6A7A8Funcoes                    As MSSOAPLib30.SoapClient30
#Else
    Dim objA6A7A8Funcoes                    As A8MIU.clsA6A7A8Funcoes
#End If

Dim intQtdeDiasValidos                      As Integer
Dim intIncremento                           As Integer
Dim datRetorno                              As Date
Dim vntArray()                              As Variant
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
Dim strData                                 As String
Dim strDataRetorno                          As String

On Error GoTo ErrHandler

    strData = CStr(Format(pdatData, "DD/MM/YYYY"))

    Set objA6A7A8Funcoes = fgCriarObjetoMIU("A8MIU.clsA6A7A8Funcoes")
    strDataRetorno = objA6A7A8Funcoes.AdicionarDiasUteis(strData, _
                                                         pintQtdeDias, _
                                                         plngMovimento, _
                                                         vntCodErro, _
                                                         vntMensagemErro)
                                                               
    fgAdicionarDiasUteis = CDate(Format(strDataRetorno, "DD/MM/YYYY"))
    
    If vntCodErro <> 0 Then
        GoTo ErrHandler
    End If
    
    Set objA6A7A8Funcoes = Nothing

Exit Function
ErrHandler:
    Set objA6A7A8Funcoes = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgAdicionarDiasUteis", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter a descrição de débito e crédito
Public Function fgTraduzDebitoCredito(ByVal plngIndicadorDebitoCredito As Long) As String

    Select Case plngIndicadorDebitoCredito
        Case enumTipoDebitoCredito.Credito
            fgTraduzDebitoCredito = "Crédito"
        Case enumTipoDebitoCredito.Debito
            fgTraduzDebitoCredito = "Débito"
        Case enumTipoDebitoCreditoEstorno.EstornoCredito
            fgTraduzDebitoCredito = "Estorno Crédito"
        Case enumTipoDebitoCreditoEstorno.EstornoDebito
            fgTraduzDebitoCredito = "Estorno Débito"
    End Select

End Function

'Obter a chave do combo
Public Function fgObterCodigoCombo(ByVal strConteudoCombo As String) As String

Dim intPOSSeparator                         As Integer

    intPOSSeparator = InStr(1, strConteudoCombo, "-")
    If intPOSSeparator = 0 Then
        fgObterCodigoCombo = vbNullString
    Else
        fgObterCodigoCombo = Trim$(Left$(strConteudoCombo, intPOSSeparator - 1))
    End If
    
End Function

'Obter a descrição do combo
Public Function fgObterDescricaoCombo(ByVal strConteudoCombo As String) As String

Dim intPOSSeparator                         As Integer

    intPOSSeparator = InStr(1, strConteudoCombo, "-")
    If intPOSSeparator = 0 Then
        fgObterDescricaoCombo = vbNullString
    Else
        fgObterDescricaoCombo = Trim$(Mid$(strConteudoCombo, intPOSSeparator + 1))
    End If
    
End Function

'Devolve string extra que foi gravada em outro controle
Public Function fgObterCampoExtraCombo(cbo As VB.ComboBox, lst As VB.ListBox) As String
    
    Dim idx As Long
    idx = cbo.ListIndex
    
    If idx = -1 Then
        fgObterCampoExtraCombo = ""
    Else
        lst.ListIndex = idx
        fgObterCampoExtraCombo = lst.Text
    End If
   
End Function

'Validar o email
Public Function fgValidarEmail(ByVal pstrMail As String) As Boolean
        
Dim strTlds                            As Variant
Dim strLeft                              As String
        
    fgValidarEmail = True
    
    If Mid(Trim(pstrMail), 1, 1) = "@" Then
        fgValidarEmail = False
    ElseIf InStr(1, pstrMail, "@") = 0 Then
        fgValidarEmail = False
    ElseIf InStr(1, pstrMail, ".") = 0 Then
        fgValidarEmail = False
    ElseIf Not (InStr(1, pstrMail, " ") = 0) Then
        fgValidarEmail = False
    Else
        strTlds = Split(pstrMail, ".")
        strLeft = strTlds(UBound(strTlds))
        strLeft = LCase(strLeft)
        'Domain is a TLD.
        '??|com|net|org|edu|int|mil|gov|arpa|biz|aero|name|coop|info|pro|museum
        If Not (Len(strLeft) = 2) Then
            If Len(strLeft) = 3 Then
                If Not (strLeft = "com" Or strLeft = "net" Or strLeft = "org" Or strLeft = "edu" Or strLeft = "int" Or strLeft = "mil" Or strLeft = "gov" Or strLeft = "biz" Or strLeft = "pro") Then
                    fgValidarEmail = False
                End If
            ElseIf Len(strLeft) = 4 Then
                If Not (strLeft = "arpa" Or strLeft = "aero" Or strLeft = "name" Or strLeft = "coop" Or strLeft = "info") Then
                    fgValidarEmail = False
                End If
            ElseIf Len(strLeft) = 6 Then
                If Not (strLeft = "museum") Then
                    fgValidarEmail = False
                End If
            Else
                fgValidarEmail = False
            End If
        End If
    End If
   
End Function

'criar a intancia do objeto MIU
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

Dim strCommandLine()                        As String

On Error GoTo ErrorHandler
    
    If App.PrevInstance Then End
    
    App.OleRequestPendingTimeout = 20000
    App.OleRequestPendingMsgTitle = "A8 - Liquidação e Controle das Câmaras"
    App.OleRequestPendingMsgText = "Servidor processando. Aguarde."
    
    gstrMascaraDataDtp = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "sShortDate")
    gstrSeparadorData = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "sDate")
    
    strCommandLine = Split(Command(), ";")
    
    If strCommandLine(0) <> "Desenv" Then
        blnDesenv = False
        flObterConfiguracaoCommandLine
        
        If gblnRegistraTLB Then
            fgRegistraComponentes
        End If
        flExibeVersao
        fgControlarAcesso
    Else
        blnDesenv = True
        gblnPerfilManutencao = True
        gstrURLWebService = strCommandLine(5)
        glngTimeOut = strCommandLine(6)
        gstrPrint = strCommandLine(7)
        gstrHelpFile = strCommandLine(8)
    End If
    
    fgCarregarIntervalos
    
    mdiLQS.Show
    
    DoEvents
    
    Exit Sub
ErrorHandler:
    
    MsgBox "Erro-> " & Err.Description & vbCrLf & "O Sistema será finalizado "
    
    End
    
End Sub
 
'carregar o intervalo para disponibilizar alerta
Public Sub fgCarregarIntervalos()

Dim strIntervalo                            As String

    strIntervalo = GetSetting("A8LQS", "DisponibilizacaoAlerta", "Intervalo", "5")
    glngTempoAlerta = CLng("0" & strIntervalo)
    If glngTempoAlerta <= 0 Then
       glngTempoAlerta = 1
    End If
    glngContaMinutosAlerta = glngTempoAlerta
    
    glngTempoContingencia = 5
    glngContaMinutosContingencia = glngTempoContingencia

End Sub
 
'Converter a data do parâmetro para o XML
Public Function fgDate_To_DtXML(ByVal pdtData As Date) As String

    fgDate_To_DtXML = Format(Year(pdtData), "0000") & Format(Month(pdtData), "00") & Format(Day(pdtData), "00")

End Function

'Converter a data/hora do parâmetro para o XML
Public Function fgDateHr_To_DtHrXML(ByVal pdtDataHora As String) As String

    fgDateHr_To_DtHrXML = Format(Year(pdtDataHora), "0000") & Format(Month(pdtDataHora), "00") & Format(Day(pdtDataHora), "00") & _
                          Format(Hour(pdtDataHora), "00") & Format(Minute(pdtDataHora), "00") & Format(Second(pdtDataHora), "00")

End Function

'Obter os veículos legais
Public Sub fgLerCarregarVeiculoLegal(ByRef pcboGrupoVeicLegal As ComboBox, _
                                     ByRef pcboVeiculoLegal As ComboBox, _
                                     ByRef pxmlPropriedades As MSXML2.DOMDocument40, _
                            Optional ByRef parrSistema As Variant)

#If EnableSoap = 1 Then
    Dim objConsulta                         As MSSOAPLib30.SoapClient30
#Else
    Dim objConsulta                         As A8MIU.clsMIU
#End If

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim strLeitura                              As String
Dim intCont                                 As Integer
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    Set objConsulta = fgCriarObjetoMIU("A8MIU.clsMIU")
        
    DoEvents
    Call fgCursor(True)

    Call fgAppendNode(pxmlPropriedades, "Grupo_Propriedades", "CO_GRUP_VEIC_LEGA", "")
    
    pxmlPropriedades.selectSingleNode("//*/@Objeto").Text = "A6A7A8.clsVeiculoLegal"
    pxmlPropriedades.selectSingleNode("//*/@Operacao").Text = "LerTodos"
    
    If pcboGrupoVeicLegal.ListIndex > 0 Then
        pxmlPropriedades.selectSingleNode("//*/CO_GRUP_VEIC_LEGA").Text = fgObterCodigoCombo(pcboGrupoVeicLegal.Text)
    Else
        pxmlPropriedades.selectSingleNode("//*/CO_GRUP_VEIC_LEGA").Text = vbNullString
    End If
    
    strLeitura = objConsulta.Executar(pxmlPropriedades.xml, _
                                      vntCodErro, _
                                      vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Call fgRemoveNode(pxmlPropriedades, "CO_GRUP_VEIC_LEGA")
    
    pcboVeiculoLegal.Clear
    pcboVeiculoLegal.AddItem "<-- Todos -->"
    
    If strLeitura <> "" Then
        Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlLeitura.loadXML(strLeitura) Then
            Call fgErroLoadXML(xmlLeitura, App.EXEName, "frmFiltroMonitoracao", "fgLerCarregarVeiculoLegal")
        End If
        
        If Not IsMissing(parrSistema) Then
            intCont = 1
            ReDim parrSistema(intCont To xmlLeitura.documentElement.selectNodes("//Repeat_VeiculoLegal/*").length)
        End If
        
        For Each objDomNode In xmlLeitura.documentElement.selectNodes("//Repeat_VeiculoLegal/*")
            pcboVeiculoLegal.AddItem objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
            
            If Not IsMissing(parrSistema) Then
                parrSistema(intCont) = objDomNode.selectSingleNode("SG_SIST").Text & "k_" & _
                                       objDomNode.selectSingleNode("CO_VEIC_LEGA").Text
            
                intCont = intCont + 1
            End If
        Next
    End If
    
    pcboVeiculoLegal.ListIndex = 0
    
    Set objConsulta = Nothing
    Set xmlLeitura = Nothing
    Set objDomNode = Nothing
    
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:

    Set objConsulta = Nothing
    Set xmlLeitura = Nothing
    Set objDomNode = Nothing

    Call fgCursor(False)
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "basA8LQS", "fgLerCarregarVeiculoLegal", 0
    
End Sub

'Centralizar o Form
Public Sub fgCenterMe(NameFrm As Form)

Dim iTop                                    As Integer

On Error Resume Next
        
    NameFrm.Left = (mdiLQS.ScaleWidth - NameFrm.Width) / 2   ' Center form horizontally.
    
    If NameFrm.MDIChild Then
       iTop = (mdiLQS.ScaleHeight - NameFrm.Height) / 2 - 640
    Else
       iTop = (mdiLQS.ScaleHeight - NameFrm.Height) / 2 + 200
    End If
    
    If iTop < 0 Then iTop = 0
    
    NameFrm.Top = iTop ' Center form vertically.
    
End Sub

'Converter o cursor
Public Sub fgCursor(Optional pbStatus As Boolean = False)
    
    If pbStatus Then
        Screen.MousePointer = vbHourglass
    Else
        Screen.MousePointer = vbDefault
    End If
    
End Sub

'Verifica se existe o item no list view
Public Function fgExisteItemLvw(ByRef objListView As MSComctlLib.ListView, _
                                ByVal strKey As String) As Boolean
Dim objListItem                             As ListItem
On Error GoTo ErrorHandler
                                
    For Each objListItem In objListView.ListItems
        If objListItem.Key = strKey Then
           fgExisteItemLvw = True
           Exit Function
        End If
    Next
                                
ErrorHandler:

    fgExisteItemLvw = False
    Err.Clear

End Function

'Perquisar o combo
Public Sub fgSearchItemCombo(ByRef pcbo As ComboBox, _
                    Optional ByVal pitem As Integer = 0, _
                    Optional ByVal ptext As String = vbNullString)

Dim lItem                                   As Integer

On Error GoTo ErrorHandler
    
    pcbo.ListIndex = -1
    
    For lItem = 0 To pcbo.ListCount - 1
        If Trim(ptext) <> "" Then
            If Left$(pcbo.List(lItem), Len(Trim$(ptext))) = Trim(ptext) Then
                pcbo.ListIndex = lItem
                Exit For
            End If
        Else
            If pcbo.ItemData(lItem) = pitem Then
                pcbo.ListIndex = lItem
                Exit For
            End If
        End If
    Next
    
    Exit Sub
    
ErrorHandler:
    
    fgRaiseError App.EXEName, "basA8LQS", "fgSearchItemCombo", 0

End Sub

'Limpar os caracteres inválidos
Public Function fgLimpaCaracterInvalido(pTexto As String)

Dim sRet                                As String

On Error GoTo ErrorHandler

    sRet = ""
    sRet = Replace(pTexto, Chr(CAR_APOSTROFE), "")
    sRet = Replace(sRet, Chr(CAR_ABRE_CHAVE), "")
    sRet = Replace(sRet, Chr(CAR_FECHA_CHAVE), "")
    sRet = Replace(sRet, Chr(CAR_SUBST), "")
    sRet = Replace(sRet, Chr(CAR_ASPAS), "")
    sRet = Replace(sRet, Chr(CAR_CASP1), "")
    sRet = Replace(sRet, Chr(CAR_CASP2), "")
    sRet = Replace(sRet, Chr(CAR_ASPAS1), "")
    sRet = Replace(sRet, Chr(CAR_ASPAS2), "")
    
    fgLimpaCaracterInvalido = sRet

    Exit Function

ErrorHandler:
    
    fgRaiseError App.EXEName, "basA8LQS", "fgLimpaCaracterInvalido", 0
    
End Function

'Função genérica para tratamento de erros.
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
'    If ErrDescription = vbNullString Or InStr(1, ErrDescription, "Grupo_ErrorInfo") = 0 Then
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

'Carregar combo com informações do xml
Public Sub fgCarregarCombos(ByRef cboComboBox As ComboBox, _
                            ByRef xmlMapaNavegacao As MSXML2.DOMDocument40, _
                            ByRef strTagName As String, _
                            ByRef strCodigoBaseName As String, _
                            ByRef strDescricaoBaseName As String, _
                   Optional ByVal blnAdicionaItemTodos As Boolean = False, _
                   Optional ByRef lstCampoExtra As ListBox, _
                   Optional ByVal strCampoExtra As String, _
                   Optional ByVal strCondicao As String)

Dim objDomNode                              As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler
    
    cboComboBox.Clear
    
    If blnAdicionaItemTodos Then
        cboComboBox.AddItem "<-- Todos -->"
        If Not lstCampoExtra Is Nothing Then
            lstCampoExtra.AddItem ""
        End If
    End If
    
    If xmlMapaNavegacao.parseError = 0 Then
        For Each objDomNode In xmlMapaNavegacao.documentElement.selectNodes("//Repeat_" & strTagName & "/*" & strCondicao)
            cboComboBox.AddItem objDomNode.selectSingleNode(strCodigoBaseName).Text & " - " & _
                                objDomNode.selectSingleNode(strDescricaoBaseName).Text
                                    
            'Coloca campos extras no listbox
            If Not lstCampoExtra Is Nothing Then
                lstCampoExtra.AddItem objDomNode.selectSingleNode(strCampoExtra).Text
            End If
        Next
    End If
    
    If blnAdicionaItemTodos Then
        cboComboBox.ListIndex = 0
    Else
        cboComboBox.ListIndex = -1
    End If
    Set objDomNode = Nothing
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:
    Set objDomNode = Nothing
    Call fgCursor(False)
    
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgCarregarCombos", 0)
    
End Sub

'carregar os grupos de veículo legal
Public Sub fgCarregarGrupoVeicLegal(ByRef cboComboBox As ComboBox, _
                                    ByRef xmlMapaNavegacao As MSXML2.DOMDocument40)

Dim objDomNode                              As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler
    
    cboComboBox.Clear
      
    For Each objDomNode In xmlMapaNavegacao.documentElement.selectNodes("//Repeat_GrupoVeiculoLegal/*")
        cboComboBox.AddItem objDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text & " - " & _
                            objDomNode.selectSingleNode("DE_GRUP_VEIC_LEGA").Text
        cboComboBox.ItemData(cboComboBox.NewIndex) = objDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text
    Next
    
    cboComboBox.ListIndex = 0
    
    Set objDomNode = Nothing
    
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:

    Set objDomNode = Nothing

    Call fgCursor(False)
    
    fgRaiseError App.EXEName, "basA8LQS", "fgCarregarGrupoVeicLegal", 0
    
End Sub

'Carregar o tipo backoffice
Public Sub fgCarregarTipoBackOffice(ByRef cboComboBox As ComboBox, _
                                    ByRef xmlMapaNavegacao As MSXML2.DOMDocument40)

Dim objDomNode                              As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler
    
    cboComboBox.Clear
      
    For Each objDomNode In xmlMapaNavegacao.documentElement.selectNodes("//Repeat_TipoBackOffice/*")
        cboComboBox.AddItem objDomNode.selectSingleNode("TP_BACK_OFFC").Text & " - " & _
                            objDomNode.selectSingleNode("DE_BACK_OFFC").Text
        cboComboBox.ItemData(cboComboBox.NewIndex) = objDomNode.selectSingleNode("TP_BACK_OFFC").Text
    Next
    
    cboComboBox.ListIndex = -1
    
    Set objDomNode = Nothing
    
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:

    Set objDomNode = Nothing

    Call fgCursor(False)
    
    fgRaiseError App.EXEName, "basA8LQS", "fgCarregarTipoBackOffice", 0
    
End Sub

'Verificar a maior data do parâmetro
Public Function fgMaiorData(ByVal pdatData1 As Date, ByVal pdatData2 As Date) As Date
    fgMaiorData = IIf(pdatData1 > pdatData2, pdatData1, pdatData2)
End Function

Public Function fgLockWindow(Optional ByVal lngWHnd As Long)

    LockWindowUpdate lngWHnd

End Function

'Obter a descrição do tipo da ação
Public Function fgDescricaoTipoAcao(ByVal plngTipoAcao As Long) As String

    Select Case plngTipoAcao
        Case enumTipoAcao.AlteracaoHorarioAgendamento
            fgDescricaoTipoAcao = "Alteração Horário Agendamento"
        Case enumTipoAcao.AlteracaoTipoCompromisso
            fgDescricaoTipoAcao = "Alteração Tipo Compromisso"
        Case enumTipoAcao.CancelamentoSolicitado
            fgDescricaoTipoAcao = "Cancelamento Solicitado"
        Case enumTipoAcao.CancelamentoEnviado
            fgDescricaoTipoAcao = "Cancelamento Enviado"
        Case enumTipoAcao.EstornoSolicitado
            fgDescricaoTipoAcao = "Estorno Solicitado"
        Case enumTipoAcao.EstornoEnviado
            fgDescricaoTipoAcao = "Estorno Enviado"
        Case enumTipoAcao.RejeicaoConcordancia
            fgDescricaoTipoAcao = "Rejeição Concordância"
        Case enumTipoAcao.RejeicaoDiscordancia
            fgDescricaoTipoAcao = "Rejeição Discordância"
        Case enumTipoAcao.EnviadaLDL1002
            fgDescricaoTipoAcao = "Enviada LDL1002"
        Case enumTipoAcao.EnviadaSEL1023
            fgDescricaoTipoAcao = "Enviada SEL1023"
        Case enumTipoAcao.EnviadaLDL1016
            fgDescricaoTipoAcao = "Enviada LDL1016"
        Case enumTipoAcao.EnviadaLDL1006
            fgDescricaoTipoAcao = "Enviada LDL1006"
        Case enumTipoAcao.RejeicaoConcordanciaAdmArea
            fgDescricaoTipoAcao = "Rejeição Concordância Adm. Área"
        Case enumTipoAcao.RejeicaoDiscordanciaAdmArea
            fgDescricaoTipoAcao = "Rejeição Discordância Adm. Área"
        Case enumTipoAcao.EnviadaLDL0003Concordancia
            fgDescricaoTipoAcao = "Enviada LDL0003 Concordância"
        Case enumTipoAcao.EnviadaLDL0003Discordancia
            fgDescricaoTipoAcao = "Enviada LDL0003 Discordância"
        Case enumTipoAcao.EnviadoPagamento
            fgDescricaoTipoAcao = "Enviado Pagamento"
        Case enumTipoAcao.EnviadoPagamentoContingencia
            fgDescricaoTipoAcao = "Enviado Pagamento Contingência"
        Case enumTipoAcao.AjusteValor
            fgDescricaoTipoAcao = "Ajuste de Valor"
        Case enumTipoAcao.EnviadaLTR0002Concordancia
            fgDescricaoTipoAcao = "Enviada LTR0002 Concordância"
        Case enumTipoAcao.EnviadaLTR0002Discordancia
            fgDescricaoTipoAcao = "Enviada LTR0002 Discordância"
        Case enumTipoAcao.EnviadoPagamentoBACEN
            fgDescricaoTipoAcao = "Enviado Pagamento BACEN"
        Case enumTipoAcao.EnviadoPagamentoSTR
            fgDescricaoTipoAcao = "Enviado Pagamento STR"
        Case enumTipoAcao.EnviadaLTR0008Concordancia
            fgDescricaoTipoAcao = "Enviada LTR0008 Concordância"
        Case enumTipoAcao.EnviadaLTR0008Discordancia
            fgDescricaoTipoAcao = "Enviada LTR0008 Discordância"
        Case enumTipoAcao.ConcordanciaEmSer
            fgDescricaoTipoAcao = "Concordância Em Ser"
        Case enumTipoAcao.ConcordanciaManualEmSer
            fgDescricaoTipoAcao = "Concordância Manual Em Ser"
        Case enumTipoAcao.RejeicaoConcordanciaEmSer
            fgDescricaoTipoAcao = "Rejeição Concordância Em Ser"
        Case enumTipoAcao.RejeicaoConcordanciaManualEmSer
            fgDescricaoTipoAcao = "Rejeição Concordância Manual Em Ser"
        Case enumTipoAcao.Liberacao
            fgDescricaoTipoAcao = "Liberação"
        Case enumTipoAcao.LiberacaoAntecipada
            fgDescricaoTipoAcao = "Liberação Antecipada"
        Case enumTipoAcao.Liquidacao
            fgDescricaoTipoAcao = "Liquidação"
        Case enumTipoAcao.EnviadaSTR0007
            fgDescricaoTipoAcao = "Enviada STR0007"
        Case enumTipoAcao.EnviadaLDL1003
            fgDescricaoTipoAcao = "Enviada LDL1003"
        Case enumTipoAcao.Concordancia
            fgDescricaoTipoAcao = "Concordância"
        Case enumTipoAcao.EnviadaSEL1007
            fgDescricaoTipoAcao = "Enviada SEL1007"
        Case enumTipoAcao.CancelamentoPendente
            fgDescricaoTipoAcao = "Cancelamento Pendente"
        Case enumTipoAcao.CancelamentoRejeitado
            fgDescricaoTipoAcao = "Cancelamento Rejeitado"
        Case enumTipoAcao.CancelamentoSolicitadoComMensagem
            fgDescricaoTipoAcao = "Cancelamento Solicitado com Mensagem"
        Case enumTipoAcao.EnviadoRecebimento
            fgDescricaoTipoAcao = "Enviado Recebimento"
        Case enumTipoAcao.EnviadaBMC0012
            fgDescricaoTipoAcao = "Enviada BMC0012"
        Case enumTipoAcao.EnviadaBMC0102
            fgDescricaoTipoAcao = "Enviada BMC0102"
        Case enumTipoAcao.RegistroContingencia
            fgDescricaoTipoAcao = "Registro em Contingência"
        Case enumTipoAcao.DiscordanciaAdmBO
            fgDescricaoTipoAcao = "Discordância Adm. BO"
        Case enumTipoAcao.LTR0001ComISPBJaExistente
            fgDescricaoTipoAcao = "LTR0001 com ISPB Creditado já existente"
        Case enumTipoAcao.LTR0007ComISPBJaExistente
            fgDescricaoTipoAcao = "LTR0007 com ISPB Creditado já existente"
        Case enumTipoAcao.RejeicaoPorDuplicidade
            fgDescricaoTipoAcao = "Identificador da Operação em Duplicidade"
        Case enumTipoAcao.EnviadaLTR0004Pagamento
            fgDescricaoTipoAcao = "Enviada LTR0004 Pagamento"
        Case enumTipoAcao.EnviadaLTR0003Pagamento
            fgDescricaoTipoAcao = "Enviada LTR0003 Pagamento"
        Case enumTipoAcao.EnviadaSTR0004Pagamento
            fgDescricaoTipoAcao = "Enviada STR0004 Pagamento"
        Case enumTipoAcao.EnviadaBMC0001
            fgDescricaoTipoAcao = "Enviada BMC0001"
        Case enumTipoAcao.EnviadaCAM0002
            fgDescricaoTipoAcao = "Enviada CAM0002"
        Case enumTipoAcao.EnviadaBMC0002
            fgDescricaoTipoAcao = "Enviada BMC0002"
        Case enumTipoAcao.EnviadaBMC0003
            fgDescricaoTipoAcao = "Enviada BMC0003"
        Case enumTipoAcao.EnviadaConfirmacaoContingencia
            fgDescricaoTipoAcao = "Enviada Confirmação Contingência"
        Case enumTipoAcao.PreviaLiquidada
            fgDescricaoTipoAcao = "Prévia Liquidada"
        Case enumTipoAcao.EnviadaCAM0054
            fgDescricaoTipoAcao = "Enviada CAM0054"
        Case enumTipoAcao.EnviadaCAM0006
            fgDescricaoTipoAcao = "Enviada CAM0006"
        Case enumTipoAcao.EnviadaCAM0009
            fgDescricaoTipoAcao = "Enviada CAM0009"
        Case enumTipoAcao.EnviadaCAM0007
            fgDescricaoTipoAcao = "Enviada CAM0007"
        Case enumTipoAcao.EnviadaCAM0010
            fgDescricaoTipoAcao = "Enviada CAM0010"
        Case enumTipoAcao.EnviadaCAM0014
            fgDescricaoTipoAcao = "Enviada CAM0014"
    End Select
    
End Function

'Obter a descrição do tipo da ação
Public Function fgDescricaoCanalVenda(ByVal plngCanalVenda As Long) As String

    Select Case plngCanalVenda
        Case enumCanalDeVenda.Nenhum
            fgDescricaoCanalVenda = "Nenhum"
        Case enumCanalDeVenda.SGM
            fgDescricaoCanalVenda = "SGM"
        Case enumCanalDeVenda.SGC
            fgDescricaoCanalVenda = "SGC"
        Case Else
            fgDescricaoCanalVenda = ""
    End Select
    
End Function

'Carregar a data de vigência
Public Sub fgCarregaDataVigencia(ByRef pdtpInicio As DTPicker, _
                                 ByRef pdtpFim As DTPicker, _
                                 ByVal pstrDataInicio As String, _
                                 ByVal pstrDataFim As String)
    
    pdtpInicio.Enabled = False
    pdtpInicio.MinDate = fgDtXML_To_Date(pstrDataInicio)
    pdtpInicio.value = pdtpInicio.MinDate
    If pdtpInicio.value > fgDataHoraServidor(Data) Then
        pdtpInicio.MinDate = fgDataHoraServidor(Data)
        pdtpInicio.Enabled = True
    End If
    
    If Trim(pstrDataFim) <> gstrDataVazia Then
        If fgDtXML_To_Date(pstrDataFim) < fgDataHoraServidor(Data) Then
            pdtpFim.MinDate = fgDtXML_To_Date(pstrDataFim)
            pdtpInicio.Enabled = True
        Else
            pdtpFim.MinDate = fgMaiorData(fgDataHoraServidor(Data), pdtpInicio.value)
        End If
        pdtpFim.value = fgDtXML_To_Date(pstrDataFim)
    Else
        pdtpFim.MinDate = fgMaiorData(fgDataHoraServidor(Data), pdtpInicio.value)
        pdtpFim.value = pdtpFim.MinDate
        pdtpFim.value = Null
    End If
                                 
End Sub

'Carregar a data de inicio de vigência
Public Sub fgDataVigenciaInicioChange(ByRef pdtpDataInicio As DTPicker, _
                                      ByRef pdtpDataFim As DTPicker)

    If pdtpDataInicio.value < fgDataHoraServidor(Data) Then
        pdtpDataInicio.value = fgDataHoraServidor(Data)
        pdtpDataInicio.MinDate = fgDataHoraServidor(Data)
    End If
    pdtpDataFim.MinDate = pdtpDataInicio.value
    pdtpDataFim.value = pdtpDataInicio.value
    pdtpDataFim.value = Null

End Sub

'Carregar a data de fim de vigência
Public Sub fgDataVigenciaFimChange(ByRef pdtpDataInicio As DTPicker, _
                                   ByRef pdtpDataFim As DTPicker)

    If Not IsNull(pdtpDataFim.value) Then
        If pdtpDataFim.value < fgDataHoraServidor(Data) Then
            pdtpDataFim.value = fgDataHoraServidor(Data)
            pdtpDataFim.MinDate = fgDataHoraServidor(Data)
        End If
    End If
    
    If pdtpDataInicio.value < fgDataHoraServidor(Data) And pdtpDataInicio.Enabled Then
        pdtpDataInicio.value = fgDataHoraServidor(Data)
        pdtpDataInicio.MinDate = pdtpDataInicio.value
    End If

End Sub

Public Function fgValidarMaxDateDTPicker(ByVal pobjDtPicker As DTPicker, ByVal pdatData As Date) As Date
    fgValidarMaxDateDTPicker = IIf(pdatData > pobjDtPicker.MaxDate, pobjDtPicker.MaxDate, pdatData)
End Function

Public Function fgValidarMinDateDTPicker(ByVal pobjDtPicker As DTPicker, ByVal pdatData As Date) As Date
    fgValidarMinDateDTPicker = IIf(pdatData < pobjDtPicker.MinDate, pobjDtPicker.MinDate, pdatData)
End Function

'Obter a configuração do commandline
Private Sub flObterConfiguracaoCommandLine()

Dim strCommandLine                           As String
Dim vntParametros                            As Variant

On Error GoTo ErrorHandler
    
    strCommandLine = Command()
    
    vntParametros = Split(strCommandLine, ";")
    
    gstrAmbiente = vntParametros(LBound(vntParametros))
    gstrSource = vntParametros(LBound(vntParametros) + 1)
    gstrUsuario = vntParametros(LBound(vntParametros) + 2)
    gblnAcessoOnLine = (UCase(vntParametros(LBound(vntParametros) + 3)) = "ON")
    
    #If EnableSoap = 1 Then
        gblnRegistraTLB = False
    #Else
        gblnRegistraTLB = vntParametros(LBound(vntParametros) + 4)
    #End If
    
    gstrURLWebService = vntParametros(LBound(vntParametros) + 5)
    glngTimeOut = vntParametros(LBound(vntParametros) + 6)
    gstrPrint = vntParametros(LBound(vntParametros) + 7)
    gstrHelpFile = vntParametros(LBound(vntParametros) + 8)
    
    Exit Sub
ErrorHandler:
    
    Err.Raise vbObjectError + 266, "strParametros", "Parâmetros Inválidos- Command Line"
    
End Sub

'Registrar os componentes
Public Sub fgRegistraComponentes()

Dim strCLIREG32                              As String
Dim strArquivo                               As String

On Error GoTo ErrorHandler
        
    strCLIREG32 = App.Path & "\CLIREG32.EXE"
        
    If gblnRegistraTLB Then
        'Registra os novos componentes
        strArquivo = App.Path & "\A8MIU"
        Shell strCLIREG32 & " """ & strArquivo & ".VBR"" -t """ & strArquivo & ".TLB"" -d -nologo -q -s " & gstrSource & " -l"
    End If
    
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA8LQS", "fgRegistraComponentes Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

'Desregistrar os componentes
Public Sub fgDesregistraComponentes()

Dim strCLIREG32                              As String
Dim strArquivo                               As String

On Error GoTo ErrorHandler
    
    strCLIREG32 = App.Path & "\CliReg32.Exe"
    
    'Caso tenha registrado os componentes, fuma tudo
    If gblnRegistraTLB Then
        
        'Desregistra os componentes
         strArquivo = App.Path & "\A8MIU"
         Shell strCLIREG32 & " """ & strArquivo & ".VBR"" -t """ & strArquivo & ".TLB"" -u -d -nologo -q -l"
            
    End If
    
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA8LQS", "fgDesregistraComponentes Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
        
End Sub

'Gerar informações da tela para o Excel
Public Sub flGeraDadosExcel(ByRef pobjExcel As Excel.Application, _
                            ByVal pControle As Control, _
                            ByVal blnPrimeiroGrid As Boolean)

Dim llCol                                   As Long
Static llRow                                As Long

Dim llMaxLen                                As Long
Dim llTotalLinhas                           As Long
Dim lsRange                                 As String
Dim lsSeparadorDecimal                      As String
Dim ListItem                                As MSComctlLib.ListItem
Dim strAux                                  As String
Dim llMaxLenHeader                          As Long

On Error GoTo ErrorHandler
    
    lsSeparadorDecimal = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "SDecimal")
    
    With pobjExcel.Worksheets(gintIndexWorksheets)
        
        If TypeOf pControle Is MSFlexGrid Then
        
            pControle.ReDraw = False
        
            llTotalLinhas = 3

            For llCol = 0 To pControle.Cols - 1
                
                For llRow = 0 To pControle.Rows - 1
                        
                        If pControle.ColWidth(llCol) <> 0 Then
                            If IsNumeric(pControle.TextMatrix(llRow, llCol)) Then
                                If Len(Trim(strAux)) <= 28 Then
                                    If InStr(1, pControle.TextMatrix(llRow, llCol), lsSeparadorDecimal) > 0 Then
                                        .Cells(llRow + llTotalLinhas, llCol + 1) = fgVlrXml_To_Interface(pControle.TextMatrix(llRow, llCol))
                                    Else
                                        .Cells(llRow + llTotalLinhas, llCol + 1) = pControle.TextMatrix(llRow, llCol)
                                    End If
                                    
                                    lsRange = Chr(65) + CStr(llCol + 1)
                                    DoEvents
                                    
                                    If InStr(1, pControle.TextMatrix(llRow, llCol), lsSeparadorDecimal) > 0 Then
                                        .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                                    Else
                                        .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "#,##0_);[Red](#,##0)"
                                    End If
                                Else
                                    .Cells(llTotalLinhas, llCol + 1).NumberFormat = "@"
                                    .Cells(llTotalLinhas, llCol + 1) = CVar(pControle.TextMatrix(llRow, llCol))
                                End If
                                
                            ElseIf IsDate(pControle.TextMatrix(llRow, llCol)) Then
                                If IsTime(pControle.TextMatrix(llRow, llCol)) Then
                                    .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "HH:MM"
                                    .Cells(llRow + llTotalLinhas, llCol + 1) = Format(pControle.TextMatrix(llRow, llCol), "HH:MM")
                                    lsRange = Chr(65) + CStr(llCol + 1)
                                    DoEvents
                                ElseIf Hour(pControle.TextMatrix(llRow, llCol)) <> 0 Or Minute(pControle.TextMatrix(llRow, llCol)) <> 0 Or Second(pControle.TextMatrix(llRow, llCol)) <> 0 Then
                                    .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                                    .Cells(llRow + llTotalLinhas, llCol + 1) = " " & pControle.TextMatrix(llRow, llCol)
                                    lsRange = Chr(65) + CStr(llCol + 1)
                                    DoEvents
                                Else
                                    .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "dd/mm/yyyy"
                                    .Cells(llRow + llTotalLinhas, llCol + 1) = " " & pControle.TextMatrix(llRow, llCol)
                                    lsRange = Chr(65) + CStr(llCol + 1)
                                    DoEvents
                                End If
                            Else
                                .Cells(llRow + llTotalLinhas, llCol + 1) = Trim(pControle.TextMatrix(llRow, llCol))
                            End If
            
                            If blnPrimeiroGrid Then
                            
                                pControle.Col = llCol
                                pControle.Row = llRow
                                
                                .Cells(llRow + llTotalLinhas, llCol + 1).Font.Bold = pControle.CellFontBold
                                .Cells(llRow + llTotalLinhas, llCol + 1).Font.Color = pControle.CellFontBold
                                .Cells(llRow + llTotalLinhas, llCol + 1).VerticalAlignment = xlBottom
                                
                                Select Case pControle.CellAlignment
                                    Case flexAlignCenterCenter
                                        .Cells(llRow + llTotalLinhas, llCol + 1).HorizontalAlignment = xlCenter
                                    Case flexAlignLeftCenter
                                        .Cells(llRow + llTotalLinhas, llCol + 1).HorizontalAlignment = xlLeft
                                    Case flexAlignRightCenter
                                        .Cells(llRow + llTotalLinhas, llCol + 1).HorizontalAlignment = xlRight
                                End Select
                                
                                If llMaxLen < Len(Trim(pControle.TextMatrix(llRow, llCol))) Then
                                    llMaxLen = Len(Trim(pControle.TextMatrix(llRow, llCol)))
                                End If
                            End If
                                           
                            If llRow <= pControle.FixedRows - 1 Then
                                .Cells(llRow + llTotalLinhas, llCol + 1).Font.Bold = True
                            End If
                               
                        End If
                Next
                
                If blnPrimeiroGrid Then
                    .Cells(llTotalLinhas, llCol + 1).ColumnWidth = llMaxLen
                    llMaxLen = 0
                End If
            
            Next
            
            If .Cells(1, 1).ColumnWidth > 0 Then
                .Cells(1, 1) = "" 'flObterNomeEmpresa
                .Cells(1, 1).Font.Bold = True
            Else
                .Cells(1, 2) = "" 'flObterNomeEmpresa
                .Cells(1, 2).Font.Bold = True
            End If

            
            pControle.ReDraw = True
        
        ElseIf TypeOf pControle Is ListView Then
        
            For llCol = 0 To pControle.ColumnHeaders.Count - 1
                
                'Se tamanho da coluna do listview = 0, o tamanho da coluna do excel deve ser 0
                If Val(pControle.ColumnHeaders(llCol + 1).Width) = 0 Then
                    .Cells(llTotalLinhas, llCol + 1).ColumnWidth = 0
                Else
                
                    llTotalLinhas = 3
                    llMaxLenHeader = Len(pControle.ColumnHeaders(llCol + 1).Text)
                    
                    .Cells(3, llCol + 1) = pControle.ColumnHeaders(llCol + 1).Text
                    .Cells(3, llCol + 1).Font.Bold = True
                    .Cells(3, llCol + 1).EntireColumn.AutoFit
                    
                    Select Case pControle.ColumnHeaders(llCol + 1).Alignment
                        Case lvwColumnCenter
                            .Cells(1, llCol + 1).HorizontalAlignment = xlCenter
                        Case lvwColumnLeft
                            .Cells(1, llCol + 1).HorizontalAlignment = xlLeft
                        Case lvwColumnRight
                            .Cells(1, llCol + 1).HorizontalAlignment = xlRight
                    End Select
                    
                    
                    For Each ListItem In pControle.ListItems
                        
                        llTotalLinhas = llTotalLinhas + 1
                        
                        If llCol = 0 Then
                            strAux = ListItem.Text
                        Else
                            strAux = ListItem.SubItems(llCol)
                        End If
                        
                        If IsNumeric(strAux) Then
                        
                            If Len(Trim(strAux)) <= 28 Then
                                If InStr(1, strAux, lsSeparadorDecimal) > 0 Then
                                    .Cells(llTotalLinhas, llCol + 1) = fgVlrXml_To_Interface(strAux)
                                Else
                                    .Cells(llTotalLinhas, llCol + 1) = strAux
                                End If
                                
                                lsRange = Chr(65) + CStr(llCol + 1)
                                DoEvents
                                
                                If InStr(1, strAux, lsSeparadorDecimal) > 0 Then
                                    .Cells(llTotalLinhas, llCol + 1).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                                Else
                                    .Cells(llTotalLinhas, llCol + 1).NumberFormat = "#,##0_);[Red](#,##0)"
                                End If
                            
                            Else
                                .Cells(llTotalLinhas, llCol + 1).NumberFormat = "@"
                                .Cells(llTotalLinhas, llCol + 1) = CVar(strAux)
                            End If
                            
                        ElseIf IsDate(strAux) Then
                            If Len(strAux) < 6 Then
                                .Cells(llTotalLinhas, llCol + 1) = strAux
                            ElseIf Hour(strAux) <> 0 Or Minute(strAux) <> 0 Or Second(strAux) <> 0 Then
                                If Len(strAux) > 9 Then
                                    .Cells(llTotalLinhas, llCol + 1) = " " & Format(strAux, "dd/mm/yyyy hh:mm:ss")
                                    lsRange = Chr(65) + CStr(llCol + 1)
                                    DoEvents
                                    .Cells(llTotalLinhas, llCol + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                                Else
                                    .Cells(llTotalLinhas, llCol + 1) = " " & Format(strAux, "hh:mm:ss")
                                    lsRange = Chr(65) + CStr(llCol + 1)
                                    DoEvents
                                    .Cells(llTotalLinhas, llCol + 1).NumberFormat = "hh:mm:ss"
                                End If
                            Else
                                .Cells(llTotalLinhas, llCol + 1) = " " & Format(strAux, "dd/mm/yyyy")
                                lsRange = Chr(65) + CStr(llCol + 1)
                                DoEvents
                                .Cells(llTotalLinhas, llCol + 1).NumberFormat = "dd/mm/yyyy"
                            End If
                            
                        Else
                            .Cells(llTotalLinhas, llCol + 1) = Trim(strAux)
                        End If
    
                        Select Case pControle.ColumnHeaders(llCol + 1).Alignment
                            Case lvwColumnCenter
                                .Cells(llTotalLinhas, llCol + 1).HorizontalAlignment = xlCenter
                            Case lvwColumnLeft
                                .Cells(llTotalLinhas, llCol + 1).HorizontalAlignment = xlLeft
                            Case lvwColumnRight
                                .Cells(llTotalLinhas, llCol + 1).HorizontalAlignment = xlRight
                        End Select
                            
                        If llMaxLen < Len(strAux) Then
                            llMaxLen = Len(strAux)
                        End If
                                       
                    Next ListItem
                    
                    If llMaxLen > 0 And llMaxLen > llMaxLenHeader Then
                        .Cells(llTotalLinhas, llCol + 1).ColumnWidth = llMaxLen
                        llMaxLen = 0
                    End If
                
                End If
                
            Next llCol
    
            .Cells(1, 1) = "" 'flObterNomeEmpresa
            .Cells(1, 1).Font.Bold = True

        ElseIf TypeOf pControle Is vaSpread Then
        
            If blnPrimeiroGrid Then
                llTotalLinhas = 3
            Else
                llTotalLinhas = llRow + 4
            End If
            
            pControle.ReDraw = False

            For llCol = 1 To pControle.MaxCols
                
                For llRow = 0 To pControle.MaxRows
                
                    pControle.Row = llRow
                    pControle.Col = llCol
                        
                    If pControle.ColWidth(llCol) <> 0 Then
                        
                        If IsNumeric(pControle.Text) Then
                            
                            If Len(Trim(strAux)) <= 28 Then
                                If InStr(1, pControle.Text, lsSeparadorDecimal) > 0 Then
                                    .Cells(llRow + llTotalLinhas, llCol + 1) = fgVlrXml_To_Interface(pControle.Text)
                                Else
                                    .Cells(llRow + llTotalLinhas, llCol + 1) = pControle.Text
                                End If
                                
                                lsRange = Chr(65) + CStr(llCol + 1)
                                DoEvents
                                
                                If InStr(1, pControle.Text, lsSeparadorDecimal) > 0 Then
                                    .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                                Else
                                    .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "#,##0_);[Red](#,##0)"
                                End If
                            Else
                                .Cells(llTotalLinhas, llCol + 1).NumberFormat = "@"
                                .Cells(llTotalLinhas, llCol + 1) = CVar(pControle.Text)
                            End If
                            
                        ElseIf IsDate(pControle.Text) Then
                            If IsTime(pControle.Text) Then
                                .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "HH:MM"
                                .Cells(llRow + llTotalLinhas, llCol + 1) = Format(pControle.Text, "HH:MM")
                                lsRange = Chr(65) + CStr(llCol + 1)
                                DoEvents
                            ElseIf Hour(pControle.Text) <> 0 Or Minute(pControle.Text) <> 0 Or Second(pControle.Text) <> 0 Then
                                .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                                .Cells(llRow + llTotalLinhas, llCol + 1) = " " & pControle.Text
                                lsRange = Chr(65) + CStr(llCol + 1)
                                DoEvents
                            Else
                                .Cells(llRow + llTotalLinhas, llCol + 1).NumberFormat = "dd/mm/yyyy"
                                .Cells(llRow + llTotalLinhas, llCol + 1) = " " & pControle.Text
                                lsRange = Chr(65) + CStr(llCol + 1)
                                DoEvents
                            End If
                        Else
                            .Cells(llRow + llTotalLinhas, llCol + 1) = Trim(pControle.Text)
                        End If
        
                        If blnPrimeiroGrid Then
                            .Cells(llRow + llTotalLinhas, llCol + 1).Font.Bold = pControle.FontBold
                            .Cells(llRow + llTotalLinhas, llCol + 1).Font.Color = vbAutomatic
                            .Cells(llRow + llTotalLinhas, llCol + 1).VerticalAlignment = xlBottom
                            
                            Select Case pControle.TypeHAlign
                                Case TypeHAlignCenter
                                    .Cells(llRow + llTotalLinhas, llCol + 1).HorizontalAlignment = xlCenter
                                Case TypeHAlignLeft
                                    .Cells(llRow + llTotalLinhas, llCol + 1).HorizontalAlignment = xlLeft
                                Case TypeHAlignRight
                                    .Cells(llRow + llTotalLinhas, llCol + 1).HorizontalAlignment = xlRight
                            End Select
                            
                            If llMaxLen < Len(Trim(pControle.Text)) Then
                                llMaxLen = Len(Trim(pControle.Text))
                            End If
                        End If
                           
                    End If
                Next
                
                If blnPrimeiroGrid Then
                    .Cells(llTotalLinhas, llCol + 1).ColumnWidth = llMaxLen
                    llMaxLen = 0
                End If
            
            Next
            
            If .Cells(1, 1).ColumnWidth > 0 Then
                .Cells(1, 1) = "" 'flObterNomeEmpresa
                .Cells(1, 1).Font.Bold = True
            Else
                .Cells(1, 2) = "" 'flObterNomeEmpresa
                .Cells(1, 2).Font.Bold = True
            End If
            
            pControle.ReDraw = True

        End If
        
    End With

Exit Sub
ErrorHandler:
    pControle.ReDraw = True
    mdiLQS.uctlogErros.MostrarErros Err, "basA8LQS"
End Sub

'Gerar informações da tela para o Excel
Public Sub fgExportaExcel(ByVal pForm As Form, _
                 Optional ByVal pvControle As Variant)

Dim pControle                               As Control
Dim objExcel                                As Excel.Application
Dim blnPrimeiroGrid                         As Boolean

    On Error GoTo ErrorHandler
    
    fgCursor True
    
    Set objExcel = CreateObject("Excel.Application")
    gintIndexWorksheets = 1
    objExcel.Workbooks.Add
    blnPrimeiroGrid = True
    If Not IsMissing(pvControle) Then
        objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets   '****** Só Inclui esta linha
        If TypeOf pvControle Is MSFlexGrid Then
            If pControle.Rows > pControle.FixedRows Then
                Call flGeraDadosExcel(objExcel, pvControle, blnPrimeiroGrid)
            End If
        ElseIf TypeOf pvControle Is ListView Then
            Call flGeraDadosExcel(objExcel, pvControle, blnPrimeiroGrid)
        ElseIf TypeOf pvControle Is vaSpread Then
            Call flGeraDadosExcel(objExcel, pvControle, blnPrimeiroGrid)
        End If
        
        blnPrimeiroGrid = False
    Else
        For Each pControle In pForm.Controls
            If TypeOf pControle Is MSFlexGrid Then
                If pControle.Rows > pControle.FixedRows Then
                    If objExcel.Worksheets.Count < gintIndexWorksheets Then
                       objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                    End If
                    objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets   '****** Só Inclui esta linha
                    Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
                    gintIndexWorksheets = gintIndexWorksheets + 1
                    blnPrimeiroGrid = False
                End If
            ElseIf TypeOf pControle Is ListView Then
                If objExcel.Worksheets.Count < gintIndexWorksheets Then
                   objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                End If
                objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets   '****** Só Inclui esta linha
                Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
                gintIndexWorksheets = gintIndexWorksheets + 1
                blnPrimeiroGrid = False
            ElseIf TypeOf pControle Is vaSpread Then
                If objExcel.Worksheets.Count < gintIndexWorksheets Then
                   objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                End If
                objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets   '****** Só Inclui esta linha
                Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
                gintIndexWorksheets = gintIndexWorksheets + 1
                blnPrimeiroGrid = False
            End If
        Next
    End If
    
    If blnPrimeiroGrid = True Then
        MsgBox "Não existem dados à serem exportados para o Excel.", vbInformation, "Atenção"
    Else
        objExcel.Visible = True
    End If
    
    Set objExcel = Nothing

    fgCursor

    Exit Sub

ErrorHandler:
    fgCursor
    Set objExcel = Nothing
    mdiLQS.uctlogErros.MostrarErros Err, "basA8LQS"
End Sub

'Gerar informações da tela para o PDF
Public Sub fgExportaPDF(ByVal pForm As Form, _
               Optional ByVal pvControle As Variant)

Dim pControle                               As Control
Dim objExcel                                As Excel.Application
Dim blnPrimeiroGrid                         As Boolean

    On Error GoTo ErrorHandler
    
    fgCursor True
    
    Set objExcel = CreateObject("Excel.Application")
    gintIndexWorksheets = 1
    objExcel.Workbooks.Add
    blnPrimeiroGrid = True
    If Not IsMissing(pvControle) Then
        objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets   '****** Só Inclui esta linha
        If TypeOf pvControle Is MSFlexGrid Then
            If pControle.Rows > pControle.FixedRows Then
                Call flGeraDadosExcel(objExcel, pvControle, blnPrimeiroGrid)
            End If
        ElseIf TypeOf pvControle Is ListView Then
            Call flGeraDadosExcel(objExcel, pvControle, blnPrimeiroGrid)
        ElseIf TypeOf pvControle Is vaSpread Then
            Call flGeraDadosExcel(objExcel, pvControle, blnPrimeiroGrid)
        End If
        
        blnPrimeiroGrid = False
    Else
        For Each pControle In pForm.Controls
            'FlexGrid
            If TypeOf pControle Is MSFlexGrid Then
                If pControle.Rows > pControle.FixedRows Then
                    If objExcel.Worksheets.Count < gintIndexWorksheets Then
                       objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                    End If
                    objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets   '****** Só Inclui esta linha
                    Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
                    
                    'Para cada Sheet gera um arquivo PDF
                    objExcel.Worksheets(gintIndexWorksheets).Select
                    With objExcel.Worksheets(gintIndexWorksheets).PageSetup
                        .Orientation = xlLandscape
                        .Zoom = 70
                        .LeftMargin = Application.InchesToPoints(0.393700787401575)
                        .RightMargin = Application.InchesToPoints(0.393700787401575)
                        .TopMargin = Application.InchesToPoints(0.393700787401575)
                        .BottomMargin = Application.InchesToPoints(0.393700787401575)
                        .HeaderMargin = Application.InchesToPoints(0.47244094488189)
                        .FooterMargin = Application.InchesToPoints(0.47244094488189)
                    End With
                    
                    objExcel.Worksheets(gintIndexWorksheets).Cells.Select
                    objExcel.Worksheets(gintIndexWorksheets).Cells.EntireColumn.AutoFit
                    
                    objExcel.Worksheets(gintIndexWorksheets).PrintOut ActivePrinter:=gstrPrint, Collate:=True
                    
                    gintIndexWorksheets = gintIndexWorksheets + 1
                    blnPrimeiroGrid = False
                    
                End If
            'ListView
            ElseIf TypeOf pControle Is ListView Then
                If objExcel.Worksheets.Count < gintIndexWorksheets Then
                   objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                End If
                objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets   '****** Só Inclui esta linha
                Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
                
                'Para cada Sheet gera um arquivo PDF
                objExcel.Worksheets(gintIndexWorksheets).Select
                With objExcel.Worksheets(gintIndexWorksheets).PageSetup
                    .Orientation = xlLandscape
                    .Zoom = 70
                    .LeftMargin = Application.InchesToPoints(0.393700787401575)
                    .RightMargin = Application.InchesToPoints(0.393700787401575)
                    .TopMargin = Application.InchesToPoints(0.393700787401575)
                    .BottomMargin = Application.InchesToPoints(0.393700787401575)
                    .HeaderMargin = Application.InchesToPoints(0.47244094488189)
                    .FooterMargin = Application.InchesToPoints(0.47244094488189)
                End With
                
                objExcel.Worksheets(gintIndexWorksheets).Cells.Select
                objExcel.Worksheets(gintIndexWorksheets).Cells.EntireColumn.AutoFit
                
                objExcel.Worksheets(gintIndexWorksheets).PrintOut ActivePrinter:=gstrPrint, Collate:=True
                
                gintIndexWorksheets = gintIndexWorksheets + 1
                blnPrimeiroGrid = False
            'Spread
            ElseIf TypeOf pControle Is vaSpread Then
                If objExcel.Worksheets.Count < gintIndexWorksheets Then
                   objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                End If
                objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pForm.Name, 26) & " " & gintIndexWorksheets   '****** Só Inclui esta linha
                Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
                
                'Para cada Sheet gera um arquivo PDF
                objExcel.Worksheets(gintIndexWorksheets).Select
                With objExcel.Worksheets(gintIndexWorksheets).PageSetup
                    .Orientation = xlLandscape
                    .Zoom = 70
                    .LeftMargin = Application.InchesToPoints(0.393700787401575)
                    .RightMargin = Application.InchesToPoints(0.393700787401575)
                    .TopMargin = Application.InchesToPoints(0.393700787401575)
                    .BottomMargin = Application.InchesToPoints(0.393700787401575)
                    .HeaderMargin = Application.InchesToPoints(0.47244094488189)
                    .FooterMargin = Application.InchesToPoints(0.47244094488189)
                End With
                
                objExcel.Worksheets(gintIndexWorksheets).Cells.Select
                objExcel.Worksheets(gintIndexWorksheets).Cells.EntireColumn.AutoFit
                
                objExcel.Worksheets(gintIndexWorksheets).PrintOut ActivePrinter:=gstrPrint, Collate:=True
                
                gintIndexWorksheets = gintIndexWorksheets + 1
                blnPrimeiroGrid = False
            End If
        Next
    End If
    
    If blnPrimeiroGrid = True Then
        MsgBox "Não existem dados à serem exportados para o PDF.", vbInformation, "Atenção"
    End If
    
    Set objExcel = Nothing

    fgCursor

    Exit Sub

ErrorHandler:
    fgCursor
    Set objExcel = Nothing
    mdiLQS.uctlogErros.MostrarErros Err, "basA8LQS"
End Sub

Public Function fgMaiorValor(ByVal pdblValor1 As Double, ByVal pdblValor2 As Double) As Double
    fgMaiorValor = IIf(pdblValor1 > pdblValor2, pdblValor1, pdblValor2)
End Function

'Marcar/Desmarcar todos os registros
Public Sub fgMarcarDesmarcarTodas(ByVal lstListView As ListView, _
                                  ByVal plngTipoSelecao As enumTipoSelecao)

Dim lngLinha                                As Long

    On Error GoTo ErrorHandler

    For lngLinha = 1 To lstListView.ListItems.Count
        lstListView.ListItems(lngLinha).Checked = (plngTipoSelecao = enumTipoSelecao.MarcarTodas)
    Next

    Exit Sub

ErrorHandler:
   fgRaiseError App.EXEName, "basA8LQS", "fgMarcarDesmarcarTodas", 0

End Sub

Public Function fgMenorValor(ByVal pdblValor1 As Double, ByVal pdblValor2 As Double) As Double
    fgMenorValor = IIf(pdblValor1 < pdblValor2, pdblValor1, pdblValor2)
End Function

'Classificar o List View
Public Sub fgClassificarListview(ByRef List As ListView, _
                                 ByVal Coluna As Integer, _
                        Optional ByVal blnManterOrdemClassificacao As Boolean = False)

Dim blnTestDate                          As Boolean
Dim blnTestNumber                        As Boolean
Dim vntConteudo                          As Variant
Dim objItemAux                           As ListItem

    If Coluna < 1 Then Exit Sub
    
    blnTestDate = True
    
    For Each objItemAux In List.ListItems
        If Coluna = 1 Then
            vntConteudo = objItemAux.Text
        Else
            vntConteudo = objItemAux.SubItems(Coluna - 1)
        End If
        
        If Not IsDate(vntConteudo) Then
            blnTestDate = False
            Exit For
        End If
    Next
    
    blnTestNumber = True
    
    For Each objItemAux In List.ListItems
        If Coluna = 1 Then
            vntConteudo = objItemAux.Text
        Else
            vntConteudo = objItemAux.SubItems(Coluna - 1)
        End If
        
        If Not IsNumeric(vntConteudo) Then
            blnTestNumber = False
            Exit For
        End If
    Next
    
    With List
        If blnTestDate Or blnTestNumber Then
            .ColumnHeaders.Add , "AUX", "DATA", 0
            For Each objItemAux In List.ListItems
                If Coluna = 1 Then
                    vntConteudo = objItemAux.Text
                Else
                    vntConteudo = objItemAux.SubItems(Coluna - 1)
                End If
                
                If blnTestDate Then
                    vntConteudo = Format$(vntConteudo, "yyyymmdd hh:mm:ss")
                Else
                    vntConteudo = Format$(vntConteudo, "000000000.000000000")
                End If
                
                objItemAux.SubItems(.ColumnHeaders.Count - 1) = vntConteudo
            Next
            Coluna = .ColumnHeaders.Count
        End If
            
        If Not .Sorted Then
            .Sorted = True
            .SortKey = Coluna - 1
            .SortOrder = lvwAscending
        Else
            .SortKey = Coluna - 1
            If Not blnManterOrdemClassificacao Then
                If .SortOrder = lvwAscending Then
                    .SortOrder = lvwDescending
                Else
                    .SortOrder = lvwAscending
                End If
            End If
        End If
        
        If blnTestDate Or blnTestNumber Then
            Call .ColumnHeaders.Remove(.ColumnHeaders.Item("AUX").Index)
        End If
        
    
        
    End With
    
End Sub

'Verifica o controle de acesso
Public Function fgControlarAcesso()

#If EnableSoap = 1 Then
    Dim objPerfil                           As MSSOAPLib30.SoapClient30
#Else
    Dim objPerfil                           As A6A7A8Miu.clsPerfil
#End If

Dim xmlControleAcesso                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strControleAcesso                       As String
Dim objControl                              As Control
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    For Each objControl In mdiLQS.Controls
        If TypeName(objControl) = "Menu" Then
            If objControl.Caption <> "-" Then
                objControl.Enabled = False
            End If
        End If
    Next

    Set objPerfil = fgCriarObjetoMIU("A6A7A8Miu.clsPerfil")
    strControleAcesso = objPerfil.ObterControleAcesso("A8", _
                                                      vntCodErro, _
                                                      vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objPerfil = Nothing
   
    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlControleAcesso.loadXML(strControleAcesso) Then
        fgErroLoadXML xmlControleAcesso, App.EXEName, "basA8LQS", "fgControlarAcesso"
    End If
    
    For Each objControl In mdiLQS.Controls
        If TypeName(objControl) = "Menu" Then
            If objControl.Caption <> "-" Then
                Set xmlNode = xmlControleAcesso.selectSingleNode("//Grupo_Acesso[Perfil='" & UCase(objControl.Name) & "']/Perfil")
                If Not xmlNode Is Nothing Then
                    objControl.Enabled = True
                End If
            End If
        End If
    Next
   
    'Verifica se o usuário está associado ao GRUPO MANUTENÇÃO,
    'através da função << MANUT >>
    Set xmlNode = xmlControleAcesso.selectSingleNode("//Grupo_Acesso[Perfil='MANUT']/Perfil")
    If Not xmlNode Is Nothing Then
        gblnPerfilManutencao = True
    End If
    
    mdiLQS.mnuAjuda.Enabled = True
    mdiLQS.mnuAjudaManual.Enabled = True
    mdiLQS.mnuSobre.Enabled = True
    
    Set xmlControleAcesso = Nothing
    
    Exit Function
    
ErrorHandler:
    Set xmlControleAcesso = Nothing
    Set objPerfil = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "basA8LQS", "fgControlarAcesso", 0
    
End Function

'Mostar o filtro
Public Function fgMostraFiltro(ByVal pstrFiltroXML As String, _
                      Optional ByVal pblnVerificaDataAtual As Boolean = True) As Boolean

Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim blnRetorno                              As Boolean
Dim intCont                                 As Integer
Dim strDataXML                              As String

    On Error GoTo ErrorHandler
    
    'Carrega o XML de filtro e verifica se é válido
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(pstrFiltroXML) Then
        Call fgErroLoadXML(xmlDomFiltro, App.EXEName, "basA8LQS", "fgMostraFiltro")
    End If
    
    'Verifica se a DATA INICIAL existe...
    If Not xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataIni") Is Nothing Then
        If pblnVerificaDataAtual Then
            'Captura apenas a data em formato invertido...
            For intCont = 1 To Len(xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataIni").Text)
                If IsNumeric(Mid(xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataIni").Text, intCont, 1)) Then
                    strDataXML = strDataXML & Mid(xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataIni").Text, intCont, 1)
                End If
            Next
        
            '...e compara com a DATA do SERVIDOR
            If Len(strDataXML) <> 8 Then
                strDataXML = Left$(strDataXML, 8)
            End If
            
            If fgDtXML_To_Date(strDataXML) <> fgDataHoraServidor(DataAux) Then
                blnRetorno = True
            End If
        End If
    Else
        blnRetorno = True
    End If
    
    fgMostraFiltro = blnRetorno

    Exit Function
    
ErrorHandler:
    Set xmlDomFiltro = Nothing
    
    fgRaiseError App.EXEName, "basA8LQS", "fgMostraFiltro", 0

End Function

'Verificar o regsitro no flexgrid
Public Sub fgPositionRowFlexGrid(ByVal intRow As Integer, _
                                 ByVal flxFlexGrid As MSFlexGrid)

Dim intCol                                  As Integer
Dim intRowClear                             As Integer

    On Error GoTo ErrorHandler
        
    With flxFlexGrid
        If .Rows = 1 Then Exit Sub
        
        .ReDraw = False
        
        .FillStyle = flexFillRepeat
        .Row = 1
        .Col = 0
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .CellBackColor = vbWhite
        .CellForeColor = vbAutomatic
        
        .Row = intRow
        .Col = 0
        .RowSel = intRow
        .ColSel = .Cols - 1
        .CellBackColor = &H8000000D
        .CellForeColor = vbWhite
        
        .GridLinesFixed = flexGridInset
        
        .ReDraw = True
    End With

'
'
'        'Limpa o FlexGrid.
'        If flxFlexGrid.Rows > (flxFlexGrid.FixedRows + 1) Then
'           If gintRowPositionAnt = 0 Then
'              For intCol = flxFlexGrid.FixedCols To flxFlexGrid.Cols - 1
'                  flxFlexGrid.Col = intCol
'                  For intRowClear = flxFlexGrid.FixedRows To flxFlexGrid.Rows - 1
'                      flxFlexGrid.Row = intRowClear
'                      flxFlexGrid.CellBackColor = vbWhite
'                      flxFlexGrid.CellForeColor = vbAutomatic
'                  Next
'              Next
'           End If
'        End If
'
'        If intRow <> gintRowPositionAnt Then
'           flxFlexGrid.Row = IIf(gintRowPositionAnt = 0, flxFlexGrid.FixedRows, gintRowPositionAnt)
'           'Pinta a linha Anterior posicionada no Grid com as Cores Preto e Branco.
'           For intCol = flxFlexGrid.FixedCols To flxFlexGrid.Cols - 1
'                flxFlexGrid.Col = intCol
'                flxFlexGrid.CellBackColor = vbWhite
'                flxFlexGrid.CellForeColor = vbAutomatic
'           Next
'
'           'Pinta a linha posicionada no Grid de Azul e Branco
'           For intCol = flxFlexGrid.FixedCols To flxFlexGrid.Cols - 1
'               flxFlexGrid.Row = intRow
'               flxFlexGrid.Col = intCol
'               flxFlexGrid.CellBackColor = &H8000000D
'               flxFlexGrid.CellForeColor = vbWhite
'           Next
'           gintRowPositionAnt = intRow
'        End If
'
'        flxFlexGrid.ReDraw = True
        
    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, "basA8LQS", "fgPositionRowFlexGrid", 0
    
End Sub

'Configurar o tamanho das colunas no listview
Public Sub flResizeCollumns(objForm As Form, ByVal lvwListView As ListView)

Dim objListItem                             As ColumnHeader
Dim sngTamForm                              As Single
Dim sngSomaWidthCollumns                    As Single

    sngTamForm = objForm.Width
    sngSomaWidthCollumns = 0
    
    For Each objListItem In lvwListView.ColumnHeaders
        sngSomaWidthCollumns = sngSomaWidthCollumns + objListItem.Width
    Next

    sngSomaWidthCollumns = (lvwListView.Width - (sngSomaWidthCollumns + 100)) \ lvwListView.ColumnHeaders.Count
    
    If sngSomaWidthCollumns < 0 Then Exit Sub
    
    For Each objListItem In lvwListView.ColumnHeaders
        objListItem.Width = objListItem.Width + sngSomaWidthCollumns
    Next

End Sub

'Exibir a versão dos componentes
Private Sub flExibeVersao()
Dim strTexto                                As String
Dim xmlVersoes                              As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode


On Error GoTo ErrorHandler

    Set xmlVersoes = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlVersoes.loadXML(fgObterDetalhesVersoes)
    
    For Each objDomNode In xmlVersoes.documentElement.childNodes
        strTexto = strTexto & _
                   objDomNode.selectSingleNode("Tipo").Text & ":" & _
                   objDomNode.selectSingleNode("Major").Text & "." & _
                   objDomNode.selectSingleNode("Minor").Text & "." & _
                   objDomNode.selectSingleNode("Revision").Text & " - "
    Next objDomNode
    strTexto = Mid$(strTexto, 1, Len(strTexto) - 3)
    mdiLQS.staLQS.Panels("Versao").Text = strTexto
    
    Set xmlVersoes = Nothing

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, "basA8LQS", "flExibeVersao", 0
    
End Sub

Public Function fgDesenv() As Boolean
    
    'Verifica se está no ambiente de Desenvolvimento
    fgDesenv = blnDesenv
    
End Function

'Verificar item checked
Public Function fgItemsCheckedListView(pListView As ListView) As Long

    Dim i As Long
    Dim lstItem As ListItem
    i = 0
    For Each lstItem In pListView.ListItems
        If lstItem.Checked Then
            i = i + 1
        End If
    Next
    fgItemsCheckedListView = i

End Function

'Ajustar o valor da operação
Function fgAjustarValorOperacao(ByVal pvntNU_SEQU_OPER_ATIV As Variant) As Boolean

    'Se ajustou,  TRUE
    'Se cancelou, FALSE
    
    Unload frmAjusteOperacao
    
    frmAjusteOperacao.SequenciaOperacao = pvntNU_SEQU_OPER_ATIV
    frmAjusteOperacao.Show vbModal
    
    fgAjustarValorOperacao = frmAjusteOperacao.blnSalvou

    Unload frmAjusteOperacao

End Function

'Obter o usuário de rede
Public Function fgObterUsuario() As String

#If EnableSoap = 1 Then
    Dim objUsuario          As MSSOAPLib30.SoapClient30
#Else
    Dim objUsuario          As A6A7A8Miu.clsUsuario
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Set objUsuario = fgCriarObjetoMIU("A6A7A8Miu.clsUsuario")
    fgObterUsuario = objUsuario.ObterUsuario(vntCodErro, _
                                             vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objUsuario = Nothing
    
Exit Function
ErrorHandler:
    Set objUsuario = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "basA8LQS", "fgObterUsuario", 0
End Function

Public Sub fgObterIntervaloVerificacao()

#If EnableSoap = 1 Then
    Dim objMonitoracao      As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracao      As A8MIU.clsOperacao
#End If

Dim xmlIntervalo            As MSXML2.DOMDocument40
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Set xmlIntervalo = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMonitoracao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    
    If xmlIntervalo.loadXML(objMonitoracao.ObterIntervaloVerificaServer(vntCodErro, vntMensagemErro)) Then

        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If

        If xmlIntervalo.selectSingleNode("@HoraInicio") Is Nothing Then
            strHoraInicioVerificacao = Trim$(xmlIntervalo.selectSingleNode("//@HoraInicio").Text)
        End If
        
        If xmlIntervalo.selectSingleNode("@HoraFim") Is Nothing Then
            strHoraFimVerificacao = Trim$(xmlIntervalo.selectSingleNode("//@HoraFim").Text)
        End If

    End If

    Set xmlIntervalo = Nothing
    Set objMonitoracao = Nothing

Exit Sub
ErrorHandler:
    
    Set xmlIntervalo = Nothing
    Set objMonitoracao = Nothing
    
    If Not IsEmpty(vntCodErro) Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, "basA8LQS", "fgObterIntervaloVerificacao", 0

End Sub

Public Sub fgCarregarXMLGeralTelaFiltro()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim xmlAuxiliar                             As MSXML2.DOMDocument40
Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim xmlCombo                                As MSXML2.DOMDocument40

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Dim objXMLNode                              As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set gxmlCombosFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlAuxiliar = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlCombo = CreateObject("MSXML2.DOMDocument.4.0")

    Call fgAppendNode(gxmlCombosFiltro, vbNullString, "Repeat_CombosFiltro", vbNullString)
    
    Call fgAppendNode(xmlAuxiliar, vbNullString, "Repeat_Acesso", vbNullString)
    Call fgAppendNode(xmlAuxiliar, "Repeat_Acesso", "Grupo_Acesso", vbNullString)
    Call fgAppendAttribute(xmlAuxiliar, "Grupo_Acesso", "Operacao", "LerTodos")
    Call fgAppendAttribute(xmlAuxiliar, "Grupo_Acesso", "Objeto", "A6A7A8.clsGrupoVeiculoLegal")
    Call fgAppendNode(xmlAuxiliar, "Grupo_Acesso", "TP_VIGE", "S")
    Call fgAppendNode(xmlAuxiliar, "Grupo_Acesso", "TP_SEGR", "S")
    Call fgAppendNode(xmlAuxiliar, "Grupo_Acesso", "TP_BKOF", "99")
    
    Call xmlAuxiliar.loadXML(objMIU.Executar(xmlAuxiliar.xml, vntCodErro, vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    gblnExibirTipoBackOffice = False
    
    For Each objXMLNode In xmlAuxiliar.selectNodes("//Repeat_GrupoVeiculoLegal/*")
        If Trim$(UCase$(objXMLNode.selectSingleNode("NO_GRUP_VEIC_LEGA").Text)) = "A8_GVL_BACKOFFICETODOS" Then
            gblnExibirTipoBackOffice = True
            Exit For
        End If
    Next
    
    If gblnExibirTipoBackOffice Then
        xmlLeitura.loadXML xmlAuxiliar.documentElement.selectSingleNode("//Grupo_GrupoVeiculoLegal").xml
        xmlLeitura.documentElement.selectSingleNode("@Operacao").Text = gstrOperLerTodos
        xmlLeitura.documentElement.selectSingleNode("TP_VIGE").Text = "S"
        xmlLeitura.documentElement.selectSingleNode("TP_SEGR").Text = "N"
        xmlLeitura.documentElement.selectSingleNode("TP_BKOF").Text = vbNullString
        
        xmlLeitura.documentElement.selectSingleNode("@Objeto").Text = "A6A7A8.clsGrupoVeiculoLegal"
        xmlCombo.loadXML objMIU.Executar(xmlLeitura.xml, vntCodErro, vntMensagemErro)
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        If xmlCombo.xml <> vbNullString Then Call fgAppendXML(gxmlCombosFiltro, "Repeat_CombosFiltro", xmlCombo.xml)
    Else
        If xmlAuxiliar.xml <> vbNullString Then Call fgAppendXML(gxmlCombosFiltro, "Repeat_CombosFiltro", xmlAuxiliar.xml)
    End If
    
    xmlLeitura.loadXML xmlAuxiliar.documentElement.selectSingleNode("//Grupo_GrupoVeiculoLegal").xml
    xmlLeitura.documentElement.selectSingleNode("@Operacao").Text = gstrOperLerTodos
    xmlLeitura.documentElement.selectSingleNode("TP_VIGE").Text = "S"
    xmlLeitura.documentElement.selectSingleNode("TP_SEGR").Text = IIf(gblnExibirTipoBackOffice, "N", "S")
    xmlLeitura.documentElement.selectSingleNode("TP_BKOF").Text = vbNullString
    
    xmlLeitura.documentElement.selectSingleNode("@Objeto").Text = "A6A7A8.clsEmpresa"
    xmlCombo.loadXML objMIU.Executar(xmlLeitura.xml, vntCodErro, vntMensagemErro)
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    If xmlCombo.xml <> vbNullString Then Call fgAppendXML(gxmlCombosFiltro, "Repeat_CombosFiltro", xmlCombo.xml)
    
    xmlLeitura.documentElement.selectSingleNode("@Objeto").Text = "A6A7A8.clsLocalLiquidacao"
    xmlCombo.loadXML objMIU.Executar(xmlLeitura.xml, vntCodErro, vntMensagemErro)
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    If xmlCombo.xml <> vbNullString Then Call fgAppendXML(gxmlCombosFiltro, "Repeat_CombosFiltro", xmlCombo.xml)
    
    xmlLeitura.documentElement.selectSingleNode("@Objeto").Text = "A6A7A8.clsTipoOperacao"
    xmlCombo.loadXML objMIU.Executar(xmlLeitura.xml, vntCodErro, vntMensagemErro)
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    If xmlCombo.xml <> vbNullString Then Call fgAppendXML(gxmlCombosFiltro, "Repeat_CombosFiltro", xmlCombo.xml)
    
    xmlLeitura.documentElement.selectSingleNode("@Objeto").Text = "A6A7A8.clsTipoLiquidacao"
    xmlCombo.loadXML objMIU.Executar(xmlLeitura.xml, vntCodErro, vntMensagemErro)
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    If xmlCombo.xml <> vbNullString Then Call fgAppendXML(gxmlCombosFiltro, "Repeat_CombosFiltro", xmlCombo.xml)
    
    xmlLeitura.documentElement.selectSingleNode("@Objeto").Text = "A8LQS.clsSituacaoProcesso"
    xmlCombo.loadXML objMIU.Executar(xmlLeitura.xml, vntCodErro, vntMensagemErro)
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    If xmlCombo.xml <> vbNullString Then Call fgAppendXML(gxmlCombosFiltro, "Repeat_CombosFiltro", xmlCombo.xml)

    If gblnExibirTipoBackOffice Then
        xmlLeitura.documentElement.selectSingleNode("@Objeto").Text = "A8LQS.clsTipoBackOffice"
        xmlCombo.loadXML objMIU.Executar(xmlLeitura.xml, vntCodErro, vntMensagemErro)
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        If xmlCombo.xml <> vbNullString Then Call fgAppendXML(gxmlCombosFiltro, "Repeat_CombosFiltro", xmlCombo.xml)
    End If
    
    xmlLeitura.documentElement.selectSingleNode("@Objeto").Text = "A8LQS.clsMensagem"
    xmlCombo.loadXML objMIU.Executar(xmlLeitura.xml, vntCodErro, vntMensagemErro)
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    If xmlCombo.xml <> vbNullString Then Call fgAppendXML(gxmlCombosFiltro, "Repeat_CombosFiltro", xmlCombo.xml)

    Set objMIU = Nothing
    Set xmlLeitura = Nothing
    Set xmlCombo = Nothing

    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    Set gxmlCombosFiltro = Nothing
    Set xmlLeitura = Nothing
    Set xmlCombo = Nothing
        
    If Not IsEmpty(vntCodErro) Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, "basA8LQS", "fgCarregarXMLGeralTelaFiltro", 0

End Sub

Public Function fgVerificaJanelaVerificacao() As Boolean

Dim dtmDataHoraAtual                        As Date
Dim dtmDataHoraInicio                       As Date
Dim dtmDataHoraFim                          As Date

On Error GoTo ErrorHandler

    dtmDataHoraAtual = Now
    dtmDataHoraInicio = fgDtHrStr_To_DateTime(Format$(Date, "YYYYMMDD") & fgCompletaString(strHoraInicioVerificacao, "0", 4, True) & "00")
    dtmDataHoraFim = fgDtHrStr_To_DateTime(Format$(Date, "YYYYMMDD") & fgCompletaString(strHoraFimVerificacao, "0", 4, True) & "00")
    
    If dtmDataHoraAtual >= dtmDataHoraInicio And dtmDataHoraAtual <= dtmDataHoraFim Then
        fgVerificaJanelaVerificacao = True
    Else
        fgVerificaJanelaVerificacao = False
    End If
    
    Exit Function

ErrorHandler:
         
    fgRaiseError App.EXEName, "basA8LQS", "fgVerificaJanelaVerificacao", 0

End Function

'Executar o Método Genérico 'Executar' do objeto A8MIU.clsMIU
Public Function fgMIUExecutarGenerico(ByVal pstrOperacao As String, _
                                      ByVal pstrObjeto As String, _
                                      ByVal pxmlFiltro As MSXML2.DOMDocument40) As String

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim xmlLeitura              As MSXML2.DOMDocument40
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

    On Error GoTo ErrorHandler
    
    If pxmlFiltro.xml = vbNullString Then
        Call fgAppendNode(pxmlFiltro, vbNullString, "Repeat_Filtro", vbNullString)
    End If
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlLeitura, "", "Repeat_Filtros", "")
    Call fgAppendNode(xmlLeitura, "Repeat_Filtros", "Grupo_Filtros", "")
    Call fgAppendAttribute(xmlLeitura, "Grupo_Filtros", "Operacao", pstrOperacao)
    Call fgAppendAttribute(xmlLeitura, "Grupo_Filtros", "Objeto", pstrObjeto)
    Call fgAppendXML(xmlLeitura, "Grupo_Filtros", pxmlFiltro.xml)
                    
    fgMIUExecutarGenerico = objMIU.Executar(xmlLeitura.xml, _
                                            vntCodErro, _
                                            vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    Set xmlLeitura = Nothing
    
    Exit Function

ErrorHandler:
    Set objMIU = Nothing
    Set xmlLeitura = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgMIUExecutarGenerico", 0)
    
End Function

'Mostar o resultado do processamento em lote

Public Sub fgMostrarResultado(ByVal pstrResultadoOperacao As String, ByVal pstrAcaoProcessamento As String)

    On Error GoTo ErrorHandler

    With frmResultOperacaoLote
        .strDescricaoOperacao = " " & pstrAcaoProcessamento
        .Resultado = pstrResultadoOperacao
        .Show vbModal
    End With

    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, "basA8LQS", "fgMostrarResultado", 0

End Sub

Public Function fgObterDescricaoProduto(ByVal plngProduto As Long) As String

#If EnableSoap = 1 Then
    Dim objMensagem         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem         As A8MIU.clsMensagem
#End If

Dim objNode                                 As MSXML2.IXMLDOMNode
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
Dim xmlProduto                              As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    If plngProduto = 0 Then
        fgObterDescricaoProduto = vbNullString
        Exit Function
    End If
    
    Set xmlProduto = CreateObject("MSXML2.DOMDocument.4.0")
    
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    xmlProduto.loadXML objMensagem.LerTodosDominioTabela(enumCodigoEmpresa.Meridional, _
                                                         "PJ.TB_PRODUTO", _
                                                         plngProduto, _
                                                         "", _
                                                         "", _
                                                         vntCodErro, _
                                                         vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If Not xmlProduto.selectSingleNode("//DESCRICAO") Is Nothing Then
        fgObterDescricaoProduto = xmlProduto.selectSingleNode("//DESCRICAO").Text
    End If
    
    
    Set objMensagem = Nothing
    Set xmlProduto = Nothing

Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    Set xmlProduto = Nothing

    fgRaiseError App.EXEName, "basA8LQS", "fgObterDescricaoProduto", 0
End Function



