Attribute VB_Name = "basA6SBR"

' Este compomente tem como objetivo agrupar métodos de utilização geral pelo sistema.
' Funções genéricas, constantes e variáveis públicas, também encontram-se declaradas neste componente.

Option Explicit

Public gstrSource                           As String
Public gstrAmbiente                         As String
Public gstrUsuario                          As String
Public gblnAcessoOnLine                     As Boolean
Public gblnRegistraTLB                      As Boolean
Public glngContaMinutosAlerta               As Long
Public gblnPerfilManutencao                 As Boolean

Public gstrURLWebService                    As String
Public glngTimeOut                          As Long

Public gstrHelpFile                         As String
Public gstrPrint                            As String

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'------------------ API ----------------------------------------
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFilename As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'API para Obter ID usuario logado
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'---------------------------------------------------------------

'------------------ Constants ----------------------------------
'Constants
Public Const datDataVazia                   As Date = "00:00:00"
Public Const lngCorEnabledFalse             As Long = &H808080
Public Const lngCorEnabledTrue              As Long = vbBlack

Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

'------------------ Variaveis ERRO ------------------------------
Public intNumeroSequencialErro              As Integer
Public lngCodigoErroNegocio                 As Long

Private Const MAX_COMPUTERNAME_LENGTH       As Long = 31

Public gintRowPositionAnt                   As Integer
Public gintIndexWorksheets                  As Integer

Public gstrVersao                           As String

Public gxmlCombosFiltro                     As MSXML2.DOMDocument40
Public gblnExibirTipoBackOffice             As Boolean

Public Function fgIconeApp() As StdPicture
    Set fgIconeApp = mdiSBR.Icon
End Function

Public Sub fgCarregarXMLGeralTelaFiltro()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A6MIU.clsMIU
#End If

Dim xmlAuxiliar                             As MSXML2.DOMDocument40
Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim xmlCombo                                As MSXML2.DOMDocument40

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

    On Error GoTo ErrorHandler

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
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
    
    gblnExibirTipoBackOffice = IIf(xmlAuxiliar.selectSingleNode("//Repeat_GrupoVeiculoLegal/Grupo_GrupoVeiculoLegal[CO_GRUP_VEIC_LEGA=99]") Is Nothing, False, True)
    
    If gblnExibirTipoBackOffice Then
        xmlLeitura.loadXML xmlAuxiliar.documentElement.selectSingleNode("//Grupo_GrupoVeiculoLegal").xml
        xmlLeitura.documentElement.selectSingleNode("@Operacao").Text = "LerTodos"
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
    xmlLeitura.documentElement.selectSingleNode("@Operacao").Text = "LerTodos"
    xmlLeitura.documentElement.selectSingleNode("TP_VIGE").Text = "S"
    xmlLeitura.documentElement.selectSingleNode("TP_SEGR").Text = IIf(gblnExibirTipoBackOffice, "N", "S")
    xmlLeitura.documentElement.selectSingleNode("TP_BKOF").Text = vbNullString
    
    xmlLeitura.documentElement.selectSingleNode("@Objeto").Text = "A6A7A8.clsEmpresa"
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

    Set objMIU = Nothing
    Set xmlLeitura = Nothing
    Set xmlCombo = Nothing

    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    Set gxmlCombosFiltro = Nothing
    Set xmlLeitura = Nothing
    Set xmlCombo = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "basA6SBR", "fgCarregarXMLGeralTelaFiltro", 0

End Sub

' Obtém detalhes de versões de componentes a serem exibidos aos usuários.

Public Function fgObterDetalhesVersoes() As String

#If EnableSoap = 1 Then
    Dim objVersao           As MSSOAPLib30.SoapClient30
#Else
    Dim objVersao           As A6MIU.clsVersao
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

    Set objVersao = fgCriarObjetoMIU("A6MIU.clsVersao")
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

' Compõe dados de versões de componentes a serem exibidos aos usuários.

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

' Retorna a data do componente.
Private Function flDataComponente() As String

On Error GoTo ErrorHandler
    
    flDataComponente = fgDtHr_To_Xml(FileDateTime(App.Path & "\" & App.EXEName & ".exe"))

Exit Function
ErrorHandler:
    
    flDataComponente = fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora))
    
End Function

' Função genérica para carregamento de combos a partir de XML passado como parâmetro.

Public Sub fgCarregarCombos(ByRef cboComboBox As ComboBox, _
                            ByRef xmlMapaNavegacao As MSXML2.DOMDocument40, _
                            ByRef strTagName As String, _
                            ByRef strCodigoBaseName As String, _
                            ByRef strDescricaoBaseName As String, _
                   Optional ByVal blnAdicionaItemTodos As Boolean = False)

Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    cboComboBox.Clear
    
    If blnAdicionaItemTodos Then
        cboComboBox.AddItem "<-- Todos -->"
    End If
    
    For Each objDomNode In xmlMapaNavegacao.documentElement.selectNodes("//Repeat_" & strTagName & "/*")
        cboComboBox.AddItem objDomNode.selectSingleNode(strCodigoBaseName).Text & " - " & _
                            objDomNode.selectSingleNode(strDescricaoBaseName).Text
    Next
    
    If blnAdicionaItemTodos Then
        cboComboBox.ListIndex = 0
    Else
        cboComboBox.ListIndex = -1
    End If

    Set objDomNode = Nothing
    
Exit Sub
ErrorHandler:
    Set objDomNode = Nothing
    Call fgRaiseError(App.EXEName, "basA6SBR", "fgCarregarCombos", 0)
    
End Sub

' Simula a posição Always On Top do windows, para as janelas do sistema a serem exibidas.

Public Sub flAlwaysOnTop(myfrm As Form, SetOnTop As Boolean)

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

   fgRaiseError App.EXEName, "basA6SBR", "flAlwaysOnTop", 0
    
End Sub

' Centraliza o formulário na tela.

Public Sub fgCenterMe(NameFrm As Form)

Dim intTop                                  As Integer

On Error Resume Next
        
    NameFrm.Left = (mdiSBR.ScaleWidth - NameFrm.Width) / 2   ' Center form horizontally.
    
    If NameFrm.MDIChild Then
       intTop = (mdiSBR.ScaleHeight - NameFrm.Height) / 2 - 640
    Else
       intTop = (mdiSBR.ScaleHeight - NameFrm.Height) / 2 + 200
    End If
    
    If intTop < 0 Then intTop = 0
    
    NameFrm.Top = intTop ' Center form vertically.
    
End Sub

' Classifica colunas de listviews.
Public Sub fgClassificarListview(List As MSComctlLib.ListView, Coluna As Integer)

Dim blnTestDate                          As Boolean
Dim blnTestNumber                        As Boolean
Dim vntConteudo                          As Variant
Dim objItemAux                           As MSComctlLib.ListItem

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
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        End If
        
        If blnTestDate Or blnTestNumber Then
            Call .ColumnHeaders.Remove(.ColumnHeaders.Item("AUX").Index)
        End If
    End With
End Sub

' Set cursor para default ou vbhourglass.

Public Sub fgCursor(Optional blnStatus As Boolean = False, _
                    Optional blnStop As Boolean = True)

    'If Not blnStatus And blnStop Then Stop
    
    If blnStatus Then
        Screen.MousePointer = vbHourglass
    Else
        Screen.MousePointer = vbDefault
    End If

End Sub

' Busca um item específico de um combo.

Public Sub fgSearchItemCombo(ByRef pcbo As ComboBox, ByVal pitem As Integer, Optional ByVal ptext As String)

Dim lItem                                   As Integer

On Error GoTo ErrorHandler

    pcbo.ListIndex = -1

    For lItem = 0 To pcbo.ListCount - 1
        If Trim(ptext) <> "" Then
            If Trim(pcbo.List(lItem)) = Trim(ptext) Then
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

    mdiSBR.uctLogErros.MostrarErros Err, "fgSearchItemCombo"

End Sub

' Seleciona a linha de um componente Spread.

Public Sub fgSelecionarLinhaSpread(ByVal pvasLista As vaSpread, ByVal plngCol As Long, ByVal lngRow As Long)

Dim varConteudoCelula                       As Variant
Dim lngColunaGrid                           As Long

    With pvasLista
        .BlockMode = True
        .Col = 1
        .Row = 2
        .Col2 = 7
        .Row2 = .MaxRows
        .BackColor = &H8000000E
    
        .Col = 9
        .Row = 2
        .Col2 = .MaxCols
        .Row2 = .MaxRows
        .BackColor = vbWhite
        
        lngColunaGrid = 1
        Do
            .GetText lngColunaGrid, lngRow, varConteudoCelula
            If varConteudoCelula <> vbNullString Then Exit Do
            lngColunaGrid = lngColunaGrid + 1
        Loop
        
        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = 1
        .ForeColor = vbWhite
        
        .Col = lngColunaGrid
        .Row = lngRow
        .Col2 = 7
        .Row2 = lngRow
        .BackColor = &HC0FFFF
        
        .Col = 9
        .Row = lngRow
        .Col2 = .MaxCols
        .Row2 = lngRow
        .BackColor = &HC0FFFF
    End With
    
End Sub

' Utilizada para a criação, na camada intermediária, de um objeto a ser referenciado. Necessário para a utilização
' do SOAP.

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

' Carrega configurações de linhas de comando nas propriedades do projeto.

Public Sub Main()

Dim strCommandLine()                        As String

On Error GoTo ErrorHandler

    App.OleRequestPendingTimeout = 20000
    App.OleRequestPendingMsgTitle = "A6 - Controle das Sub-reservas"
    App.OleRequestPendingMsgText = "Servidor processando. Aguarde."
    
    strCommandLine = Split(Command(), ";")
    
    If strCommandLine(0) <> "Desenv" _
    And strCommandLine(0) <> "ProvasIntegradas" Then
        
        flObterConfiguracaoCommandLine
        
        If gblnRegistraTLB Then
            fgRegistraComponentes
        End If
        
        fgControlarAcesso
    Else
        gblnPerfilManutencao = True
        gstrURLWebService = strCommandLine(5)
        glngTimeOut = strCommandLine(6)
        gstrPrint = strCommandLine(7)
        gstrHelpFile = strCommandLine(8)
    End If

    mdiSBR.Show

    DoEvents

Exit Sub
ErrorHandler:

    MsgBox "Erro-> " & Err.Description & vbCrLf & "O Sistema será finalizado "
   
    End
    
End Sub

' Trava janela evitando efeitos indesejáveis ou clicks desnecessários de usuários.

Public Function fgLockWindow(Optional ByVal lngWHnd As Long)

    LockWindowUpdate lngWHnd

End Function

' Retorna o texto que vier à esquerda de um conteúdo separado por "-".

Public Function fgObterCodigoCombo(ByVal strConteudoCombo As String) As String

Dim intPOSSeparator                         As Integer

    intPOSSeparator = InStr(1, strConteudoCombo, "-")
    If intPOSSeparator = 0 Then
        fgObterCodigoCombo = vbNullString
    Else
        fgObterCodigoCombo = Trim$(Left$(strConteudoCombo, intPOSSeparator - 1))
    End If
    
End Function

' Retorna o texto que vier à direita de um conteúdo separado por "-".

Public Function fgObterDescricaoCombo(ByVal strConteudoCombo As String) As String

Dim intPOSSeparator                         As Integer

    intPOSSeparator = InStr(1, strConteudoCombo, "-")
    If intPOSSeparator = 0 Then
        fgObterDescricaoCombo = vbNullString
    Else
        fgObterDescricaoCombo = Trim$(Mid$(strConteudoCombo, intPOSSeparator + 1))
    End If
    
End Function

' Função genérica para o carregamento de treeviews de itens de caixa para diversos formulários.

Public Sub fgCarregarTreeViewFluxoCaixa(ByRef pstrTagXML As String, _
                                        ByRef ptreTView As MSComctlLib.TreeView, _
                                        ByVal pstrarrCodQuebras As String, _
                                        ByVal pstrarrDescQuebras As String, _
                                        ByVal pintTipoBackOfficeUsuario As Integer, _
                                        ByVal pstrRegistros As String)

Dim strarrCodQuebras()                      As String
Dim strarrDescQuebras()                     As String
Dim intContNiveis                           As Integer
Dim intContTags                             As Integer

Dim xmlDomRegistros                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objTrvNode                              As MSComctlLib.Node

Dim strNodeKey                              As String
Dim strarrKeysNivelAnt()                    As String
Dim intContNodes                            As Integer

Dim intNivelItemCaixa                       As Integer
Dim laIndexParent(1 To 5)                   As String

Dim strGrupoVeiculoLegal                    As String

On Error GoTo ErrorHandler

    ptreTView.Nodes.Clear
    ptreTView.Sorted = False
    
    Set xmlDomRegistros = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlDomRegistros.loadXML(pstrRegistros) Then
        Exit Sub
    End If

    strarrCodQuebras = Split(pstrarrCodQuebras, ";")
    strarrDescQuebras = Split(pstrarrDescQuebras, ";")
    
    ReDim strarrKeysNivelAnt(xmlDomRegistros.selectNodes("/Repeat_" & pstrTagXML & "/*").length)
    
    For intContNiveis = 0 To UBound(strarrCodQuebras)
        intContNodes = 0
        
        For Each objDomNode In xmlDomRegistros.selectNodes("/Repeat_" & pstrTagXML & "/*")

            If intContNiveis = 0 Then
                If objDomNode.selectSingleNode(strarrDescQuebras(intContNiveis)).Text <> strGrupoVeiculoLegal Then
                    strNodeKey = "k_" & _
                                 objDomNode.selectSingleNode(strarrCodQuebras(intContNiveis)).Text
    
                    Set objTrvNode = ptreTView.Nodes.Add(, , strNodeKey, _
                                     objDomNode.selectSingleNode(strarrDescQuebras(intContNiveis)).Text, "itemgrupo")
                                     
                    strGrupoVeiculoLegal = objDomNode.selectSingleNode(strarrDescQuebras(intContNiveis)).Text
                    
                End If
            Else
                strNodeKey = strarrKeysNivelAnt(intContNodes) & _
                             "k_" & objDomNode.selectSingleNode(strarrCodQuebras(intContNiveis)).Text & _
                             "k_" & objDomNode.selectSingleNode("SG_SIST").Text & _
                             "k_" & objDomNode.selectSingleNode("TP_BKOF").Text

                Set objTrvNode = ptreTView.Nodes.Add(strarrKeysNivelAnt(intContNodes), tvwChild, strNodeKey, _
                                 objDomNode.selectSingleNode(strarrDescQuebras(intContNiveis)).Text, "itemgrupo")

                If objDomNode.selectSingleNode("TP_BKOF").Text <> pintTipoBackOfficeUsuario Then
                    objTrvNode.ForeColor = lngCorEnabledFalse
                End If
            End If
            
            strarrKeysNivelAnt(intContNodes) = strNodeKey
            If intContNiveis = 0 Then objTrvNode.EnsureVisible
            intContNodes = intContNodes + 1
            
        Next
    Next
    
Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, "basA6SBR", "fgCarregarTreeViewFluxoCaixa", 0

End Sub

' Função genérica para o carregamento de treeviews de itens de caixa para diversos formulários.

Public Sub fgCarregarTreItemCaixa(ByRef treCadastro As MSComctlLib.TreeView, _
                                  ByRef xmlMapaNavegacao As MSXML2.DOMDocument40, _
                                  ByVal objForm As Form)
                                  
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objNode                                 As MSComctlLib.Node
Dim strIcon                                 As String

Dim strTipoBackOffice                       As String * 1
Dim strTipoCaixa                            As String * 1
Dim strCodigoItemCaixa                      As String * 15
Dim strChavePai                             As String * 20
Dim strChaveFilho                           As String * 20

On Error GoTo ErrorHandler

    treCadastro.Nodes.Clear
    treCadastro.HideSelection = False
    
    With xmlMapaNavegacao
        .resolveExternals = False
        .validateOnParse = False
        .setProperty "SelectionLanguage", "XPath"
    End With
    
    With treCadastro
    
        For Each objDomNode In xmlMapaNavegacao.documentElement.selectNodes("//Repeat_ItemCaixa/*")

            strTipoBackOffice = objDomNode.selectSingleNode("TP_BKOF").Text
            strTipoCaixa = objDomNode.selectSingleNode("TP_CAIX").Text
            strCodigoItemCaixa = Mid(objDomNode.selectSingleNode("CO_ITEM_CAIX").Text, 2)
                          
            If objDomNode.selectSingleNode("CO_ITEM_CAIX_PAI").Text = "" Then
                strIcon = "OpenReservaFuturo"
                strChavePai = "K" & strTipoBackOffice & _
                              "K" & strTipoCaixa & _
                              "K" & strCodigoItemCaixa
                strChaveFilho = strChavePai
                
                Set objNode = .Nodes.Add(, tvwFirst, strChavePai, _
                                         objDomNode.selectSingleNode("DE_ITEM_CAIX").Text, strIcon)

            Else
                If objForm.Name = "frmItemCaixa" Or objDomNode.selectSingleNode("DE_ITEM_CAIX").Text <> gstrItemGenerico Then
                    If objDomNode.selectSingleNode("TP_ITEM_CAIX").Text = enumTipoItemCaixa.Grupo Then
                        strIcon = "Open"
                    Else
                        strIcon = "Leaf"
                    End If
                    
                    strChavePai = "K" & strTipoBackOffice & _
                                  "K" & strTipoCaixa & _
                                  "K" & Mid(objDomNode.selectSingleNode("CO_ITEM_CAIX_PAI").Text, 2)
                    
                    strChaveFilho = "K" & strTipoBackOffice & _
                                    "K" & strTipoCaixa & _
                                    "K" & strCodigoItemCaixa
                    
                    Set objNode = .Nodes.Add(strChavePai, tvwChild, strChaveFilho, _
                                             objDomNode.selectSingleNode("DE_ITEM_CAIX").Text, strIcon)
                End If
            End If
        
            If Not objNode Is Nothing Then
                If fgObterNivelItemCaixa(objDomNode.selectSingleNode("CO_ITEM_CAIX").Text) = 1 Then
                    objNode.EnsureVisible
                    objNode.Expanded = False
                End If
                objNode.Tag = objDomNode.selectSingleNode("TP_ITEM_CAIX").Text
                Set objNode = Nothing
            End If
        Next
    End With
    
    If treCadastro.Nodes.Count > 0 Then
        treCadastro.Nodes.Item(1).Selected = True
        Set objNode = treCadastro.SelectedItem
    End If
    
    Set objNode = Nothing
    
Exit Sub
ErrorHandler:
    Set objNode = Nothing
    fgRaiseError App.EXEName, "basA6SBR", "fgCarregarTreItemCaixa", 0

End Sub

' Obtém nome da estação de trabalho do usuário logado.

Public Function fgObterEstacaoTrabalho() As String

Dim strEstacao                              As String
Dim lngLen                                  As Long

On Error GoTo ErrorHandler

    lngLen = MAX_COMPUTERNAME_LENGTH + 1
    strEstacao = String(lngLen, "X")

    GetComputerName strEstacao, lngLen
    strEstacao = Left$(strEstacao, lngLen)

    fgObterEstacaoTrabalho = UCase(strEstacao)

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6SBR", "fgObterEstacaoTrabalho", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Função genéria para o acionamento da obtenção da posição atual do caixa.

Public Function fgObterPosicaoCaixaSubReserva(ByVal pstrVeiculoLegal As String, _
                                              ByVal pstrSiglaSistema As String) As String
    
#If EnableSoap = 1 Then
    Dim objCaixaSubReserva      As MSSOAPLib30.SoapClient30
#Else
    Dim objCaixaSubReserva      As A6MIU.clsCaixaSubReserva
#End If

Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant

On Error GoTo ErrorHandler

    Set objCaixaSubReserva = fgCriarObjetoMIU("A6MIU.clsCaixaSubReserva")
    fgObterPosicaoCaixaSubReserva = objCaixaSubReserva.ObterPosicaoCaixaSubReserva(pstrVeiculoLegal, _
                                                                                   pstrSiglaSistema, _
                                                                                   vbNullString, _
                                                                                   True, _
                                                                                   vntCodErro, _
                                                                                   vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objCaixaSubReserva = Nothing

    Exit Function

ErrorHandler:
    Set objCaixaSubReserva = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' Obtém tipo de backoffice do usuário logado.

Public Function fgObterTipoBackOfficeUsuario() As enumTipoBackOffice
    
#If EnableSoap = 1 Then
    Dim objControleAcesso       As MSSOAPLib30.SoapClient30
#Else
    Dim objControleAcesso       As A6MIU.clsControleAcesso
#End If

Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant
    
On Error GoTo ErrorHandler

    Set objControleAcesso = fgCriarObjetoMIU("A6MIU.clsControleAcesso")
    fgObterTipoBackOfficeUsuario = objControleAcesso.ObterTipoBackOfficeUsuario(vntCodErro, _
                                                                                vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objControleAcesso = Nothing

Exit Function
ErrorHandler:
    Set objControleAcesso = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' Obtém código do usuário logado na rede.

Public Function fgObterUsuarioRede() As String

Dim strUserName                              As String
Dim lngLen                                   As Long

On Error GoTo ErrorHandler

    lngLen = 100
    strUserName = String(lngLen, Chr$(0))
    GetUserName strUserName, lngLen
    If Len(Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)) > 8 Then
        fgObterUsuarioRede = Left(UCase(Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)), 8)
    Else
        fgObterUsuarioRede = UCase(Left$(strUserName, InStr(strUserName, Chr$(0)) - 1))
    End If

    Exit Function
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' Função genérica para o acionamento da leitura e carregamento da lista de veículos legais.

Public Sub fgLerCarregarVeiculoLegal(ByRef pcboGrupoVeicLegal As ComboBox, _
                                     ByRef pcboVeiculoLegal As ComboBox, _
                                     ByRef pxmlPropriedades As MSXML2.DOMDocument40, _
                            Optional ByRef parrSistema As Variant)

#If EnableSoap = 1 Then
    Dim objConsulta                         As MSSOAPLib30.SoapClient30
#Else
    Dim objConsulta                         As A6MIU.clsMIU
#End If

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim strLeitura                              As String
Dim intCont                                 As Integer
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    Set objConsulta = fgCriarObjetoMIU("A6MIU.clsMIU")
            
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
            Call fgErroLoadXML(xmlLeitura, App.EXEName, "frmFiltro", "fgLerCarregarVeiculoLegal")
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
    pcboVeiculoLegal.Enabled = True
    
    Set objConsulta = Nothing
    Set xmlLeitura = Nothing
    Set objDomNode = Nothing
    
Exit Sub
ErrorHandler:
    Set objConsulta = Nothing
    Set xmlLeitura = Nothing
    Set objDomNode = Nothing

    fgCursor
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmFiltro - fgLerCarregarVeiculoLegal")
    
End Sub

' Compara duas datas e retorna a maior.

Public Function fgMaiorData(ByVal pdatData1 As Date, ByVal pdatData2 As Date) As Date
    fgMaiorData = IIf(pdatData1 > pdatData2, pdatData1, pdatData2)
End Function

' Tratamento genérico de erros ocorridos no sistema.

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

' Aciona o registro de componentes.

Public Sub fgRegistraComponentes()

Dim strCLIREG32                              As String
Dim strArquivo                               As String

On Error GoTo ErrorHandler
        
    strCLIREG32 = App.Path & "\CLIREG32.EXE"
        
    If gblnRegistraTLB Then
        'Registra os novos componentes
        strArquivo = App.Path & "\A6MIU"
        Shell strCLIREG32 & " """ & strArquivo & ".VBR"" -t """ & strArquivo & ".TLB"" -d -nologo -q -s " & gstrSource & " -l"
    End If
    
Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA6SBR", "fgRegistraComponentes Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

' Torna componentes não-registrados.

Public Sub fgDesregistraComponentes()

Dim strCLIREG32                              As String
Dim strArquivo                               As String

On Error GoTo ErrorHandler
    
    strCLIREG32 = App.Path & "\CliReg32.Exe"
    
    'Caso tenha registrado os componentes, fuma tudo
    If gblnRegistraTLB Then
        
        'Desregistra os componentes
        strArquivo = App.Path & "\A6MIU"
        Shell strCLIREG32 & " """ & strArquivo & ".VBR"" -t """ & strArquivo & ".TLB"" -u -d -nologo -q -l"
            
    End If
    
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA6SBR", "fgDesregistraComponentes Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
        
End Sub

' Retorna conteúdo da linha de comando nas propriedades do projeto.

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

'Verifica se existe o item no list view

Public Function fgExisteItemLvw(ByRef objListView As MSComctlLib.ListView, _
                                ByVal strKey As String) As Boolean
Dim objListItem                             As MSComctlLib.ListItem
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

' Função genérica para a exportação de grids de pesquisa em geral, para o Excel.

Public Sub fgExportaExcel(ByVal pcolControles As Collection, _
                 Optional ByVal pblnPlanilhaNova As Boolean = True, _
                 Optional ByVal pstrCaminhoPlanilha As String = vbNullString, _
                 Optional ByVal pblnSobreporDados As Boolean = False)

Dim pControle                               As Control
Dim objExcel                                As Excel.Application
Dim blnPrimeiroGrid                         As Boolean

On Error GoTo ErrorHandler
    
    Set objExcel = CreateObject("Excel.Application")
    
    blnPrimeiroGrid = True
    
    If pblnPlanilhaNova Then
        gintIndexWorksheets = 1
        objExcel.Workbooks.Add
    Else
        If pblnSobreporDados Then
            gintIndexWorksheets = 1
        Else
            gintIndexWorksheets = 9999
        End If
        objExcel.Workbooks.Open pstrCaminhoPlanilha
    End If
    
    For Each pControle In pcolControles
        If TypeOf pControle Is MSFlexGrid Then
            If pControle.Rows > pControle.FixedRows Then
                If objExcel.Worksheets.Count < gintIndexWorksheets Then
                    objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                    gintIndexWorksheets = objExcel.Worksheets.Count
                End If
                objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pControle.Tag, 28) & " " & Format$(gintIndexWorksheets, "00")
                Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
                gintIndexWorksheets = gintIndexWorksheets + 1
                blnPrimeiroGrid = False
            End If
        ElseIf TypeOf pControle Is ListView Then
            If objExcel.Worksheets.Count < gintIndexWorksheets Then
                objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                gintIndexWorksheets = objExcel.Worksheets.Count
            End If
            objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pControle.Tag, 28) & " " & Format$(gintIndexWorksheets, "00")
            Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
            gintIndexWorksheets = gintIndexWorksheets + 1
            blnPrimeiroGrid = False
        ElseIf TypeOf pControle Is vaSpread Then
            If objExcel.Worksheets.Count < gintIndexWorksheets Then
                objExcel.Worksheets.Add , objExcel.Worksheets(objExcel.Worksheets.Count)
                gintIndexWorksheets = objExcel.Worksheets.Count
            End If
            objExcel.Worksheets(gintIndexWorksheets).Name = Left$(pControle.Tag, 28) & " " & Format$(gintIndexWorksheets, "00")
            Call flGeraDadosExcel(objExcel, pControle, blnPrimeiroGrid)
            gintIndexWorksheets = gintIndexWorksheets + 1
            blnPrimeiroGrid = False
        End If
    Next
    
    objExcel.Visible = True
    
    Set objExcel = Nothing

Exit Sub
ErrorHandler:
    fgCursor
    Set objExcel = Nothing
    mdiSBR.uctLogErros.MostrarErros Err, "basA6SBR"

End Sub

' Gera dados para a exportação para o Excel.

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
        
            pControle.Redraw = False
        
'            If blnPrimeiroGrid Then
                llTotalLinhas = 3
 '           Else
  '              llTotalLinhas = llRow + 4
   '         End If

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

            
            pControle.Redraw = True
        
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
                                    .Cells(llTotalLinhas, llCol + 1) = fgVlrXml_To_Interface(fgVlrXml_To_Decimal(strAux))
                                Else
                                    .Cells(llTotalLinhas, llCol + 1) = fgVlrXml_To_Decimal(strAux)
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
            
            pControle.Redraw = False

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
            
            pControle.Redraw = True

        End If
        
    End With

Exit Sub
ErrorHandler:
    pControle.Redraw = True
    mdiSBR.uctLogErros.MostrarErros Err, "basA6SBR"
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
    mdiSBR.uctLogErros.MostrarErros Err, "basA6SBR"
End Sub

' Carrega datas de vigência para componentes dtPicker.

Public Sub fgCarregaDataVigencia(ByRef pdtpInicio As DTPicker, _
                                 ByRef pdtpFim As DTPicker, _
                                 ByVal pstrDataInicio As String, _
                                 ByVal pstrDataFim As String)
    
    pdtpInicio.Enabled = False
    pdtpInicio.MinDate = fgDtXML_To_Date(pstrDataInicio)
    pdtpInicio.Value = pdtpInicio.MinDate
    If pdtpInicio.Value > fgDataHoraServidor(Data) Then
        pdtpInicio.MinDate = fgDataHoraServidor(Data)
        pdtpInicio.Enabled = True
    End If
    
    If Trim(pstrDataFim) <> datDataVazia Then
        If fgDtXML_To_Date(pstrDataFim) < fgDataHoraServidor(Data) Then
            pdtpFim.MinDate = fgDtXML_To_Date(pstrDataFim)
            pdtpInicio.Enabled = True
        Else
            pdtpFim.MinDate = fgMaiorData(fgDataHoraServidor(Data), pdtpInicio.Value)
        End If
        pdtpFim.Value = fgDtXML_To_Date(pstrDataFim)
    Else
        pdtpFim.MinDate = fgMaiorData(fgDataHoraServidor(Data), pdtpInicio.Value)
        pdtpFim.Value = pdtpFim.MinDate
        pdtpFim.Value = Null
    End If
                                 
End Sub

' Trata mudanças no conteúdo de um dtPicker no que diz respeito ao valor mínimo permitido.

Public Sub fgDataVigenciaInicioChange(ByRef pdtpDataInicio As DTPicker, _
                                      ByRef pdtpDataFim As DTPicker)

    If pdtpDataInicio.Value < fgDataHoraServidor(Data) Then
        pdtpDataInicio.Value = fgDataHoraServidor(Data)
        pdtpDataInicio.MinDate = fgDataHoraServidor(Data)
    End If
    pdtpDataFim.MinDate = pdtpDataInicio.Value
    pdtpDataFim.Value = pdtpDataInicio.Value
    pdtpDataFim.Value = Null

End Sub

' Trata mudanças no conteúdo de um dtPicker no que diz respeito ao valor máximo permitido.

Public Sub fgDataVigenciaFimChange(ByRef pdtpDataInicio As DTPicker, _
                                   ByRef pdtpDataFim As DTPicker)

    If Not IsNull(pdtpDataFim.Value) Then
        If pdtpDataFim.Value < fgDataHoraServidor(Data) Then
            pdtpDataFim.Value = fgDataHoraServidor(Data)
            pdtpDataFim.MinDate = fgDataHoraServidor(Data)
        End If
    End If
    
    If pdtpDataInicio.Value < fgDataHoraServidor(Data) And pdtpDataInicio.Enabled Then
        pdtpDataInicio.Value = fgDataHoraServidor(Data)
        pdtpDataInicio.MinDate = pdtpDataInicio.Value
    End If

End Sub

' Aciona o controle de acesso do usuário ao sistema.

Public Function fgControlarAcesso()

#If EnableSoap = 1 Then
    Dim objPerfil                           As MSSOAPLib30.SoapClient30
#Else
    Dim objPerfil                           As A6A7A8Miu.clsPerfil
#End If

Dim xmlControleAcesso                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strControleAcesso                       As String
Dim lngCont                                 As Long
Dim objControl                              As Control
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    For Each objControl In mdiSBR.Controls
        If TypeName(objControl) = "Menu" Then
            If objControl.Caption <> "-" Then
                objControl.Enabled = False
            End If
        End If
    Next

    Set objPerfil = fgCriarObjetoMIU("A6A7A8Miu.clsPerfil")
    strControleAcesso = objPerfil.ObterControleAcesso("A6", _
                                                      vntCodErro, _
                                                      vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objPerfil = Nothing
   
    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlControleAcesso.loadXML(strControleAcesso) Then
        fgErroLoadXML xmlControleAcesso, App.EXEName, "basA6SBR", "fgControlarAcesso"
    End If
    
    For Each objControl In mdiSBR.Controls
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
    
    mdiSBR.mnuAjuda.Enabled = True
    mdiSBR.mnuAjudaManual.Enabled = True
    mdiSBR.mnuAjudaSobre.Enabled = True
    
    Set xmlControleAcesso = Nothing
    
Exit Function
ErrorHandler:
    
    Set objPerfil = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "basA6SBR", "fgControlarAcesso", 0
    
End Function

' Posiciona na linha de corrente de um componente FlexGrid.

Public Sub fgPositionRowFlexGrid(ByVal intRow As Integer, _
                                 ByVal flxFlexGrid As MSFlexGrid)

Dim intCol                                  As Integer
Dim intRowClear                             As Integer

On Error GoTo ErrorHandler
        
        flxFlexGrid.Redraw = False
        
        'Limpa o FlexGrid.
        If flxFlexGrid.Rows > (flxFlexGrid.FixedRows + 1) Then
           If gintRowPositionAnt = 0 Then
              For intCol = flxFlexGrid.FixedCols To flxFlexGrid.Cols - 1
                  flxFlexGrid.Col = intCol
                  For intRowClear = flxFlexGrid.FixedRows To flxFlexGrid.Rows - 1
                      flxFlexGrid.Row = intRowClear
                      flxFlexGrid.CellBackColor = vbWhite
                      flxFlexGrid.CellForeColor = vbAutomatic
                  Next
              Next
           End If
        End If
        
        If intRow <> gintRowPositionAnt Then
           flxFlexGrid.Row = IIf(gintRowPositionAnt = 0, flxFlexGrid.FixedRows, gintRowPositionAnt)
           'Pinta a linha Anterior posicionada no Grid com as Cores Preto e Branco.
           For intCol = flxFlexGrid.FixedCols To flxFlexGrid.Cols - 1
                flxFlexGrid.Col = intCol
                flxFlexGrid.CellBackColor = vbWhite
                flxFlexGrid.CellForeColor = vbAutomatic
           Next
            
           'Pinta a linha posicionada no Grid de Azul e Branco
           For intCol = flxFlexGrid.FixedCols To flxFlexGrid.Cols - 1
               flxFlexGrid.Row = intRow
               flxFlexGrid.Col = intCol
               flxFlexGrid.CellBackColor = &H8000000D
               flxFlexGrid.CellForeColor = vbWhite
           Next
           gintRowPositionAnt = intRow
        End If
        
        flxFlexGrid.Redraw = True
        
Exit Sub

ErrorHandler:

    fgRaiseError App.EXEName, "basA8LQS", "fgPositionRowFlexGrid", 0
    
End Sub

' Trata a exibição da tela de filtro, de acordo com a data de filtro aplicada anteriormente,
' encontrada no registry do windows.

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
        Call fgErroLoadXML(xmlDomFiltro, App.EXEName, "basA6SBR", "fgMostraFiltro")
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

' Valida valor mínimo de data para um componente dtPicker.

Public Function fgValidarMinDateDTPicker(ByVal pobjDtPicker As DTPicker, ByVal pdatData As Date) As Date
    fgValidarMinDateDTPicker = IIf(pdatData < pobjDtPicker.MinDate, pobjDtPicker.MinDate, pdatData)
End Function

' Retorna a descrição do estado de caixa.

Public Function fgDescricaoEstadoCaixa(ByVal penumEstadoCaixa As enumEstadoCaixa) As String
    
Dim strRetorno                              As String
    
    Select Case penumEstadoCaixa
        Case enumEstadoCaixa.Disponivel
            strRetorno = "Disponível"
        Case enumEstadoCaixa.Aberto
            strRetorno = "Aberto"
        Case enumEstadoCaixa.Fechado
            strRetorno = "Fechado"
    End Select

    fgDescricaoEstadoCaixa = strRetorno

End Function
