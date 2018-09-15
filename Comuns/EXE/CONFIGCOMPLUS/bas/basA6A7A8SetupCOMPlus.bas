Attribute VB_Name = "basA6A7A8SetupCOMPlus"
'Empresa        : Regerbanc
'Pacote         :
'Classe         : basA6A7A8SetupCOMPlus
'Data Criação   : 10/10/2003
'Objetivo       : Instalar e configurar os componentes do SLCC
'
'Analista       : Adilson Gonçalves Damasceno
'
'Programador    : Eder Andrade
'Data           : 14/10/2003
'
'Teste          :
'Autor          :
'
'Data Alteração : 05/11/2003
'Autor          : Eder Andrade
'Objetivo       : Exibição e gravação dos logs das operações
'
'Data Alteração : 04/12/2003
'Autor          : Eder Andrade
'Objetivo       : Implementação da rotina que desregistra os componentes anteriores


Option Explicit

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private strNomeComputador                   As String
Private strNomeAplicacao                    As String

Public xmlLog                               As MSXML2.DOMDocument40

Private Enum enumIndicadorSimNao
    Nao = 0
    Sim = 1
End Enum

Private Function flDiretorioSistema() As String

Dim strBuffer                               As String * 255

    Call GetSystemDirectory(strBuffer, 255)
    flDiretorioSistema = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)

End Function

Private Function flBlnValue(ByVal intValor As enumIndicadorSimNao) As Boolean
    flBlnValue = intValor = enumIndicadorSimNao.Sim
End Function

Public Sub fgDesinstalar()
Dim xmlDomDescritor                         As MSXML2.DOMDocument40

Dim objCOMAdminCatalogCollection            As COMAdmin.COMAdminCatalogCollection
Dim objCOMAdminCatalogObject                As COMAdmin.COMAdminCatalogObject
'Dim objCOMAdminCatalog                      As COMAdmin.COMAdminCatalog
Dim objCOMAdminCatalog                      As Object

Dim blnExisteAplicacao                      As Boolean

Dim strCaminhoAplicacao                     As String

'Váriaveis auxiliares
Dim lngCont                                 As Long

On Error GoTo ErrorHandler

    'Lê o descritor da aplicação
    Set xmlDomDescritor = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomDescritor.Load(App.Path & "\SLCC_Componentes_COMPLUS.xml") Then
        Err.Raise vbObjectError, , "Xml descritor inválido!"
    End If
    
    strNomeAplicacao = xmlDomDescritor.documentElement.selectSingleNode("Properties/Name").Text
    strCaminhoAplicacao = xmlDomDescritor.documentElement.selectSingleNode("Repeat_Components/@Path").Text
    
    'Obtem o nome do computador
    strNomeComputador = String(100, " ")
    GetComputerName strNomeComputador, 99
   
    strNomeComputador = RTrim(strNomeComputador)
    strNomeComputador = Left(strNomeComputador, Len(strNomeComputador) - 1)
        
    'Conecta-se ao servidor COM+
    Set objCOMAdminCatalog = CreateObject("COMAdmin.COMAdminCatalog")
    objCOMAdminCatalog.Connect strNomeComputador
    
    Set objCOMAdminCatalogCollection = objCOMAdminCatalog.GetCollection("Applications")
    objCOMAdminCatalogCollection.Populate
    
    'Encontra a aplicação e remove seus componentes do COM+
    For lngCont = 0 To objCOMAdminCatalogCollection.Count - 1
        Set objCOMAdminCatalogObject = objCOMAdminCatalogCollection.Item(lngCont)
        If objCOMAdminCatalogObject.Name = strNomeAplicacao Then
            
            objCOMAdminCatalog.ShutdownApplication (strNomeAplicacao)
            
            Call fgRemoveComponentes(objCOMAdminCatalogCollection, objCOMAdminCatalogObject)
            Exit For
        End If
    Next
    
'    Shell strCaminhoAplicacao & "\A6A8ValidaRemessa.exe /unregserver"
'    Call fgAdicionarLog("A6A8ValidaRemessa.exe", "Desregistrado", "")

'    Shell "NET STOP A7VerificaServer", vbHide
'    Call fgAdicionarLog("SLCCServico.exe", "Desregistrado", "")

'    Shell "NET STOP A6A8AtivaValidaRemessa", vbHide
'    Call fgAdicionarLog("A6A8AtivaValidaRemessa.exe", "Desregistrado", "")
    
Exit Sub
ErrorHandler:

    Set objCOMAdminCatalogCollection = Nothing
    Set objCOMAdminCatalogObject = Nothing
    Set objCOMAdminCatalog = Nothing
    
    Set xmlDomDescritor = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub



Public Sub fgInstalar()

Dim xmlDomDescritor                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMElement

Dim objCOMAdminCatalogCollection            As COMAdmin.COMAdminCatalogCollection
Dim objCOMAdminCatalogObject                As COMAdmin.COMAdminCatalogObject
'Dim objCOMAdminCatalog                      As COMAdmin.COMAdminCatalog
Dim objCOMAdminCatalog                      As Object

Dim strNomeComponente                       As String
Dim strCaminhoAplicacao                     As String
Dim strCaminhoComponente                    As String

Dim blnExisteAplicacao                      As Boolean

Dim strArquivoTBL                           As String

'Váriaveis auxiliares
Dim lngCont                                 As Long

Dim lngControle                             As Long
Dim lngMaximo                               As Long

Dim strElemento                             As String
Dim strAcao                                 As String
Dim strComplemento                          As String

On Error GoTo ErrorHandler

    'Lê o descritor da aplicação
    Set xmlDomDescritor = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomDescritor.Load(App.Path & "\SLCC_Componentes_COMPLUS.xml") Then
        Err.Raise vbObjectError, , "Xml descritor inválido!"
    End If
    
    strNomeAplicacao = xmlDomDescritor.documentElement.selectSingleNode("Properties/Name").Text
    
    'Obtem o nome do computador
    strNomeComputador = String(100, " ")
    GetComputerName strNomeComputador, 99
   
    strNomeComputador = RTrim(strNomeComputador)
    strNomeComputador = Left(strNomeComputador, Len(strNomeComputador) - 1)
        
    'Conecta-se ao servidor COM+
    Set objCOMAdminCatalog = CreateObject("COMAdmin.COMAdminCatalog")
    objCOMAdminCatalog.Connect strNomeComputador
    
    Set objCOMAdminCatalogCollection = objCOMAdminCatalog.GetCollection("Applications")
    objCOMAdminCatalogCollection.Populate
    
    'Encontra a aplicação e remove seus componentes do COM+
    For lngCont = 0 To objCOMAdminCatalogCollection.Count - 1
        Set objCOMAdminCatalogObject = objCOMAdminCatalogCollection.Item(lngCont)
        If objCOMAdminCatalogObject.Name = strNomeAplicacao Then
            blnExisteAplicacao = True
            
            objCOMAdminCatalog.ShutdownApplication (strNomeAplicacao)
            
            Call fgRemoveComponentes(objCOMAdminCatalogCollection, objCOMAdminCatalogObject)
            Exit For
        End If
    Next
    
    'Instalar Aplicacao
    If Not blnExisteAplicacao Then
        Set objCOMAdminCatalogObject = objCOMAdminCatalogCollection.Add
        
        'Configura as propriedades
        For Each objDomNode In xmlDomDescritor.documentElement.selectSingleNode("Properties").childNodes
            If objDomNode.selectSingleNode("@Boolean") Is Nothing Then
                objCOMAdminCatalogObject.Value(objDomNode.baseName) = objDomNode.Text
            Else
                objCOMAdminCatalogObject.Value(objDomNode.baseName) = flBlnValue(objDomNode.Text)
            End If
        Next objDomNode
        
        fgAdicionarLog objCOMAdminCatalogObject.Name, "Criada no COM+ ", vbNullString
        
    End If
    
    objCOMAdminCatalogCollection.SaveChanges
    
    'Adiciona os componentes
    strCaminhoAplicacao = xmlDomDescritor.documentElement.selectSingleNode("Repeat_Components/@Path").Text
    lngMaximo = xmlDomDescritor.documentElement.selectNodes("Repeat_Components/*/*").length
    lngControle = 0
    
    'Registra A6A8ValidaRemessa
    'Shell strCaminhoAplicacao & "\A6A8ValidaRemessa.exe "
    'Call fgAdicionarLog("A6A8ValidaRemessa.exe", "Registrado", "")
    
    For Each objDomNode In xmlDomDescritor.documentElement.selectNodes("Repeat_Components/*/*")
        
        lngControle = lngControle + 1
        Call fgBarraStatus(lngMaximo, lngControle, "Adicionando Componentes")
        
        strNomeComponente = objDomNode.baseName
        strCaminhoComponente = strCaminhoAplicacao & "\" & strNomeComponente
    
        'Se componente usa o arquivo TLB
        If Not objDomNode.selectSingleNode("@UseTlbFile") Is Nothing Then
            If objDomNode.selectSingleNode("@UseTlbFile").Text = enumIndicadorSimNao.Sim Then
                strArquivoTBL = strCaminhoAplicacao & "\" & Split(objDomNode.baseName, ".")(0) & ".TLB"
                strComplemento = " Adicionado " & Split(objDomNode.baseName, ".")(0) & ".TLB"
            Else
                strArquivoTBL = vbNullString
                strComplemento = vbNullString
            End If
        Else
            strArquivoTBL = vbNullString
            strComplemento = vbNullString
        End If
    
        'Se for uma classe de evento
        If Not objDomNode.selectSingleNode("@EventClass") Is Nothing Then
            objCOMAdminCatalog.InstallEventClass strNomeAplicacao, strCaminhoComponente, strArquivoTBL, ""
            strAcao = "Adicionado como Classe de evento"
        Else
            objCOMAdminCatalog.InstallComponent strNomeAplicacao, strCaminhoComponente, strArquivoTBL, ""
            strAcao = "Adicionado"
        End If
        
        fgAdicionarLog strNomeComponente, strAcao, strComplemento
    
    Next
    
    Call flConfiguraTransação(xmlDomDescritor)
    
    Call flConfiguraObjectConstructor(xmlDomDescritor)
    
'    Shell "NET START A7VerificaServer", vbHide
'    Call fgAdicionarLog("SLCCServico.exe", "Registrado", "")
    
'Finalizações
    Set xmlDomDescritor = Nothing
    
    Set objCOMAdminCatalog = Nothing

Exit Sub
ErrorHandler:

    Set objCOMAdminCatalogCollection = Nothing
    Set objCOMAdminCatalogObject = Nothing
    Set objCOMAdminCatalog = Nothing
    
    Set xmlDomDescritor = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub flConfiguraTransação(ByRef pxmlDomDescritor As MSXML2.DOMDocument40)

'Dim objCOMAdminCatalog                      As COMAdmin.COMAdminCatalog
Dim objCOMAdminCatalog                      As Object
Dim objCOMAdminCatalogCollection            As COMAdmin.COMAdminCatalogCollection
Dim objCOMAdminCatalogObject                As COMAdmin.COMAdminCatalogObject

Dim objCOMComponentes                       As COMAdmin.COMAdminCatalogCollection
Dim objCOMComponente                        As COMAdmin.COMAdminCatalogObject

Dim Instances                               As COMAdmin.COMAdminCatalogCollection
Dim Instance                                As COMAdmin.COMAdminCatalogObject

Dim strNomeComponente                       As String
Dim strNomeClasse                           As String

Dim lngMaximo                               As Long
Dim lngControle                             As Long

Dim strAcao                                 As String
Dim strComplemento                          As String

On Error GoTo ErrorHandler

    Set objCOMAdminCatalog = CreateObject("COMAdmin.COMAdminCatalog")
    
    objCOMAdminCatalog.Connect strNomeComputador
    
    Set objCOMAdminCatalogCollection = objCOMAdminCatalog.GetCollection("Applications")
    objCOMAdminCatalogCollection.Populate
    
    'Encontra a aplicação no COM+
    For Each objCOMAdminCatalogObject In objCOMAdminCatalogCollection
        If objCOMAdminCatalogObject.Name = strNomeAplicacao Then
            Exit For
        End If
    Next

    'Obtém os componentes da aplicação
    Set objCOMComponentes = objCOMAdminCatalogCollection.GetCollection("Components", objCOMAdminCatalogObject.Key)
    objCOMComponentes.Populate
    
    
    lngMaximo = objCOMComponentes.Count
    lngControle = 0
    For Each objCOMComponente In objCOMComponentes
    
        lngControle = lngControle + 1
        Call fgBarraStatus(lngMaximo, lngControle, "Configurando Transações...")
    
        strNomeComponente = Split(objCOMComponente.Name, ".")(0)
        strNomeClasse = Split(objCOMComponente.Name, ".")(1)
        
        Set Instances = objCOMComponentes.GetCollection("SubscriptionsForComponent", objCOMComponente.Key)
        
        strAcao = "Transação configurada "
        
        With pxmlDomDescritor.documentElement
        
            If Not .selectSingleNode("Repeat_Class_TransactionConfig/Grupo_Not_Supported/" & strNomeComponente & "[Class = '" & strNomeClasse & "']") Is Nothing Then
                objCOMComponente.Value("Transaction") = COMAdminTransactionOptions.COMAdminTransactionNone
                strComplemento = "Not Supported "
                
                If Not .selectSingleNode("Repeat_Class_TransactionConfig/Grupo_Not_Supported/" & strNomeComponente & "[Class = '" & strNomeClasse & "']/@ComponentTransactionTimeoutEnabled") Is Nothing Then
                    objCOMComponente.Value("ComponentTransactionTimeoutEnabled") = CBool(.selectSingleNode("Repeat_Class_TransactionConfig/Grupo_Not_Supported/" & strNomeComponente & "[Class = '" & strNomeClasse & "']/@ComponentTransactionTimeoutEnabled").Text)
                    objCOMComponente.Value("ComponentTransactionTimeout") = CBool(.selectSingleNode("Repeat_Class_TransactionConfig/Grupo_Not_Supported/" & strNomeComponente & "[Class = '" & strNomeClasse & "']/@ComponentTransactionTimeout").Text)
                End If
                
            ElseIf Not .selectSingleNode("Repeat_Class_TransactionConfig/Grupo_Required/" & strNomeComponente & "[Class = '" & strNomeClasse & "']") Is Nothing Then
                objCOMComponente.Value("Transaction") = COMAdminTransactionOptions.COMAdminTransactionRequired
                strComplemento = "Required "
                
                If Not .selectSingleNode("Repeat_Class_TransactionConfig/Grupo_Required/" & strNomeComponente & "/Class[.='" & strNomeClasse & "']/@ComponentTransactionTimeoutEnabled") Is Nothing Then
                    objCOMComponente.Value("ComponentTransactionTimeoutEnabled") = CBool(.selectSingleNode("Repeat_Class_TransactionConfig/Grupo_Required/" & strNomeComponente & "/Class[.='" & strNomeClasse & "']/@ComponentTransactionTimeoutEnabled").Text)
                    objCOMComponente.Value("ComponentTransactionTimeout") = .selectSingleNode("Repeat_Class_TransactionConfig/Grupo_Required/" & strNomeComponente & "/Class[.='" & strNomeClasse & "']/@ComponentTransactionTimeout").Text
                End If
                
            ElseIf Not .selectSingleNode("Repeat_Class_TransactionConfig/Grupo_RequiredNew/" & strNomeComponente & "[Class = '" & strNomeClasse & "']") Is Nothing Then
                objCOMComponente.Value("Transaction") = COMAdminTransactionOptions.COMAdminTransactionRequiresNew
                strComplemento = "Requires New "
                
                If Not .selectSingleNode("Repeat_Class_TransactionConfig/Grupo_RequiredNew/" & strNomeComponente & "[Class = '" & strNomeClasse & "']/@ComponentTransactionTimeoutEnabled") Is Nothing Then
                    objCOMComponente.Value("ComponentTransactionTimeoutEnabled") = CBool(.selectSingleNode("Repeat_Class_TransactionConfig/Grupo_Not_Supported/" & strNomeComponente & "[Class = '" & strNomeClasse & "']/@ComponentTransactionTimeoutEnabled").Text)
                    objCOMComponente.Value("ComponentTransactionTimeout") = CBool(.selectSingleNode("Repeat_Class_TransactionConfig/Grupo_Not_Supported/" & strNomeComponente & "[Class = '" & strNomeClasse & "']/@ComponentTransactionTimeout").Text)
                End If
                
            Else
                objCOMComponente.Value("Transaction") = COMAdminTransactionOptions.COMAdminTransactionSupported
                strComplemento = "Supported "
                
                If Not .selectSingleNode("Repeat_Class_TransactionConfig/Grupo_Not_Supported/" & strNomeComponente & "[Class = '" & strNomeClasse & "']/@ComponentTransactionTimeoutEnabled") Is Nothing Then
                    objCOMComponente.Value("ComponentTransactionTimeoutEnabled") = CBool(.selectSingleNode("Repeat_Class_TransactionConfig/Grupo_Not_Supported/" & strNomeComponente & "[Class = '" & strNomeClasse & "']/@ComponentTransactionTimeoutEnabled").Text)
                    objCOMComponente.Value("ComponentTransactionTimeout") = CBool(.selectSingleNode("Repeat_Class_TransactionConfig/Grupo_Not_Supported/" & strNomeComponente & "[Class = '" & strNomeClasse & "']/@ComponentTransactionTimeout").Text)
                End If
                
            End If
        
            If objCOMComponente.Value("IsEventClass") Then
                objCOMComponente.Value("FireInParallel") = True
            End If
            
            If Not .selectSingleNode("Subscriber_Info/" & strNomeComponente & "/Class[Name = '" & strNomeClasse & "']") Is Nothing Then
                flConfiguraSubscriber .selectSingleNode("Subscriber_Info/" & strNomeComponente & "/Class[Name = '" & strNomeClasse & "']"), _
                                      objCOMComponentes, _
                                      objCOMComponente
                                      
                fgAdicionarLog objCOMComponente.Name, "Inscrição configurada ", " Inscrito para " & .selectSingleNode("Subscriber_Info/" & strNomeComponente & "/Class[Name = '" & strNomeClasse & "']/SubscribesTo").Text
            End If
        
        End With
        
        objCOMComponentes.SaveChanges
        fgAdicionarLog objCOMComponente.Name, strAcao, strComplemento
    
    
    Next objCOMComponente
    
    
'Finalizações
    Set objCOMAdminCatalog = Nothing
    Set objCOMAdminCatalogCollection = Nothing
    Set objCOMComponentes = Nothing


Exit Sub
ErrorHandler:
    Set objCOMAdminCatalog = Nothing
    Set objCOMAdminCatalogCollection = Nothing
    Set objCOMAdminCatalogObject = Nothing
    
    Set objCOMComponentes = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub flConfiguraSubscriber(ByRef objDomNode As MSXML2.IXMLDOMNode, _
                                  ByRef objCOMComponentes As COMAdmin.COMAdminCatalogCollection, _
                                  ByRef objCOMCompSubscriber As COMAdmin.COMAdminCatalogObject)

Dim objCOMInterfaces                        As COMAdmin.COMAdminCatalogCollection
Dim objCOMInterfacePublisher                As COMAdmin.COMAdminCatalogObject
Dim objCOMInterfaceSubscriber               As COMAdmin.COMAdminCatalogObject

Dim objCOMInscricoes                        As COMAdmin.COMAdminCatalogCollection
Dim objCOMInscricao                         As COMAdmin.COMAdminCatalogObject

Dim objCOMCompPublisher                     As COMAdmin.COMAdminCatalogObject

Dim strNomeInteface                         As String

On Error GoTo ErrorHandler
    
    'Encontra o componente Publisher
    For Each objCOMCompPublisher In objCOMComponentes
        If objCOMCompPublisher.Name = objDomNode.selectSingleNode("SubscribesTo").Text Then
            Exit For
        End If
    Next objCOMCompPublisher
    
    strNomeInteface = "_" & Split(objDomNode.selectSingleNode("SubscribesTo").Text, ".")(1)
    
    'Configura a interface do publisher
    Set objCOMInterfaces = objCOMComponentes.GetCollection("InterfacesForComponent", objCOMCompPublisher.Key)
    objCOMInterfaces.Populate
    
    For Each objCOMInterfacePublisher In objCOMInterfaces
        If objCOMInterfacePublisher.Name = strNomeInteface Then
            objCOMInterfacePublisher.Value("QueuingEnabled") = True
            Exit For
        End If
    Next objCOMInterfacePublisher
    
    objCOMInterfaces.SaveChanges
    
    'Configura a interface do publisher
    Set objCOMInterfaces = objCOMComponentes.GetCollection("InterfacesForComponent", objCOMCompSubscriber.Key)
    objCOMInterfaces.Populate
    
    'Encontra a interface desejada
    For Each objCOMInterfaceSubscriber In objCOMInterfaces
        If objCOMInterfaceSubscriber.Name = strNomeInteface Then
            objCOMInterfaceSubscriber.Value("QueuingEnabled") = True
            Exit For
        End If
    Next objCOMInterfaceSubscriber

    objCOMInterfaces.SaveChanges
    
    Set objCOMInscricoes = objCOMComponentes.GetCollection("SubscriptionsForComponent", objCOMCompSubscriber.Key)
    objCOMInscricoes.Populate
    
    Set objCOMInscricao = objCOMInscricoes.Add
    
    With objCOMInterfacePublisher
        .Value("QueuingEnabled") = True
    End With
    
    objCOMInterfaces.SaveChanges
    
    With objCOMInscricao
        .Value("EventCLSID") = objCOMCompPublisher.Key
        .Value("InterfaceID") = objCOMInterfacePublisher.Key
        .Value("Name") = "A7BusServer"
        .Value("Enabled") = True
        .Value("Queued") = True
    End With
    
    objCOMInscricoes.SaveChanges
    
    Set objCOMInterfaces = Nothing
    Set objCOMInterfacePublisher = Nothing
    Set objCOMInscricoes = Nothing
    Set objCOMInscricao = Nothing
    Set objCOMCompPublisher = Nothing
    
    
Exit Sub
ErrorHandler:

    Set objCOMInterfaces = Nothing
    Set objCOMInterfacePublisher = Nothing
    Set objCOMInscricoes = Nothing
    Set objCOMInscricao = Nothing
    Set objCOMCompPublisher = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub fgBarraStatus(ByVal plngMax As Long, _
                         ByVal plngValor As Long, _
                         ByVal pstrTexto As String)
    
    DoEvents

    With frmA6A7A8SetupCOMPlus
        .prgStatus.Max = plngMax
        .prgStatus.Value = plngValor
        If pstrTexto <> vbNullString Then
            .lblStatus.Caption = pstrTexto
            .lblStatus.Visible = True
        Else
            .lblStatus.Visible = False
        End If
                
    End With

End Sub

Public Function fgAppendNode(ByRef xmlDocument As MSXML2.DOMDocument40, _
                             ByVal pstrNodeContext As String, _
                             ByVal pstrNodeName As String, _
                             ByVal pstrNodeValue As String, _
                    Optional ByVal pstrNodeRepetName As String = "") As Boolean

Dim objDOMNodeAux                           As MSXML2.IXMLDOMNode
Dim objDomNodeContext                       As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    If pstrNodeName = vbNullString Then
        'Parâmetro pstrNodeName deve ser diferente de vbNullString
        Err.Raise vbObjectError, App.EXEName & "." & "basRBLibrary", "Parâmetro pstrNodeName deve ser diferente de vbNullString"
    End If
    
    If pstrNodeContext = "" Then
        'Se a Tag for o root passar pstrNodeContext = Nome Tag Principal
        Set objDomNodeContext = xmlDocument
    Else
        'Para criar um Grupo (Ex.:Grupo_Usuario) dentro de Uma Repeticao (Ex. Repeat_Usuario)
        'Passar o argumento pNomeRepet = Repet_Usuario
        If pstrNodeRepetName <> "" Then
            Set objDomNodeContext = xmlDocument.selectSingleNode("//" & pstrNodeRepetName).childNodes.Item(xmlDocument.selectSingleNode("//" & pstrNodeRepetName).childNodes.length - 1)
        Else
            Set objDomNodeContext = xmlDocument.documentElement.selectSingleNode("//" & pstrNodeContext)
        End If
    End If
    
    Set objDOMNodeAux = xmlDocument.createElement(pstrNodeName)
    objDOMNodeAux.Text = pstrNodeValue
    objDomNodeContext.appendChild objDOMNodeAux

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Description, Err.Description
End Function


Public Sub fgAdicionarLog(ByVal pstrElemento As String, _
                          ByVal pstrAcao As String, _
                          ByVal pstrComplemento As String)

Dim lstItem                                 As ListItem
Dim objDomNodeAcao                          As MSXML2.IXMLDOMNode
Dim objDomNode                              As MSXML2.IXMLDOMNode
    
    On Error GoTo ErrorHandler
    
    If xmlLog Is Nothing Then
        Set xmlLog = CreateObject("MSXML2.DOMDocument.4.0")
        Call fgAppendNode(xmlLog, "", "Log", "")
    End If

    If xmlLog.documentElement.selectSingleNode(pstrElemento) Is Nothing Then
        fgAppendNode xmlLog, "Log", pstrElemento, ""
    End If

    
    Set objDomNodeAcao = xmlLog.createElement("Acao")
    
    Set objDomNode = xmlLog.createElement("Descricao")
    objDomNode.Text = pstrAcao
    objDomNodeAcao.appendChild objDomNode
    
    Set objDomNode = xmlLog.createElement("Hora")
    objDomNode.Text = fgDtHr_To_Xml(Now)
    objDomNodeAcao.appendChild objDomNode
        
    Set objDomNode = xmlLog.createElement("Complemento")
    objDomNode.Text = pstrComplemento
    objDomNodeAcao.appendChild objDomNode
        
    xmlLog.documentElement.selectSingleNode(pstrElemento).appendChild objDomNodeAcao
    

    Set lstItem = frmA6A7A8SetupCOMPlus.lvwStatus.ListItems.Add(, , pstrElemento)
    lstItem.SubItems(1) = pstrAcao
    lstItem.SubItems(2) = pstrComplemento
    lstItem.EnsureVisible
    
Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function fgDtHr_To_Xml(ByVal pdtmData As Date) As String

On Error GoTo ErrorHandler

    If Not IsDate(pdtmData) Then
        'Parâmetro pdtmData deve ser do tipo Date
        Err.Raise vbObjectError + 513, , "Parâmetro pdtmData deve ser do tipo 'Date'"
    End If

    fgDtHr_To_Xml = Format(pdtmData, "YYYYMMDDHHMMSS")

Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Sub fgRemoveComponentes(ByRef objCOMAdminCatalogCollection As COMAdmin.COMAdminCatalogCollection, _
                               ByRef objCOMAdminCatalogObject As COMAdmin.COMAdminCatalogObject)

Dim objCOMComponentes                       As COMAdmin.COMAdminCatalogCollection

Dim lngCont                                 As Long
Dim lngContDLL                              As Long
Dim lngControle                             As Long

Dim blnFlagDLL                              As Boolean

Dim strElemento                             As String
Dim strDLLs()                               As String
Dim strCaminhoDLL                           As String

On Error GoTo ErrorHandler

    Set objCOMComponentes = objCOMAdminCatalogCollection.GetCollection("Components", objCOMAdminCatalogObject.Key)
    objCOMComponentes.Populate

    lngControle = objCOMComponentes.Count
    lngCont = 0
    
    ReDim strDLLs(0)
    
    While objCOMComponentes.Count > 0
        lngCont = lngCont + 1
        strElemento = objCOMComponentes.Item(0).Name
        strCaminhoDLL = objCOMComponentes.Item(0).Value("DLL")
        Call fgBarraStatus(lngControle, lngCont, "Removendo componentes")
        objCOMComponentes.Remove 0
        
        blnFlagDLL = False
        For lngContDLL = LBound(strDLLs) To UBound(strDLLs)
            If strDLLs(lngContDLL) = strCaminhoDLL Then
                'A dll já está marcada
                blnFlagDLL = True
                Exit For
            End If
        Next lngContDLL
        
        If Not blnFlagDLL Then
            If strDLLs(0) = vbNullString Then
                strDLLs(0) = strCaminhoDLL
            Else
                ReDim Preserve strDLLs(UBound(strDLLs) + 1)
                strDLLs(UBound(strDLLs)) = strCaminhoDLL
            End If
        End If
        
        fgAdicionarLog strElemento, "Removido ", ""
    Wend
    
    objCOMComponentes.SaveChanges
    
    If strDLLs(0) <> vbNullString Then
        For lngContDLL = LBound(strDLLs) To UBound(strDLLs)
            Call fgBarraStatus(UBound(strDLLs), lngContDLL, "Desregistrando DLLs")
            Shell flDiretorioSistema & "\REGSVR32.exe /u /s " & Chr$(34) & strDLLs(lngContDLL) & Chr$(34)
            Call fgAdicionarLog(Mid$(strDLLs(lngContDLL), InStrRev(strDLLs(lngContDLL), "\") + 1), "Desregistrada", "")
        Next lngContDLL
    End If

Exit Sub
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub flConfiguraObjectConstructor(ByRef pxmlDomDescritor As MSXML2.DOMDocument40)

'Dim objCOMAdminCatalog                      As COMAdmin.COMAdminCatalog
Dim objCOMAdminCatalog                      As Object
Dim objCOMAdminCatalogCollection            As COMAdmin.COMAdminCatalogCollection
Dim objCOMAdminCatalogObject                As COMAdmin.COMAdminCatalogObject

Dim objCOMComponentes                       As COMAdmin.COMAdminCatalogCollection
Dim objCOMComponente                        As COMAdmin.COMAdminCatalogObject

Dim Instances                               As COMAdmin.COMAdminCatalogCollection
Dim Instance                                As COMAdmin.COMAdminCatalogObject

Dim strNomeComponente                       As String
Dim strNomeClasse                           As String

Dim lngMaximo                               As Long
Dim lngControle                             As Long
Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim strComplemento                          As String

On Error GoTo ErrorHandler

    Set objCOMAdminCatalog = CreateObject("COMAdmin.COMAdminCatalog")
    
    objCOMAdminCatalog.Connect strNomeComputador
    
    Set objCOMAdminCatalogCollection = objCOMAdminCatalog.GetCollection("Applications")
    objCOMAdminCatalogCollection.Populate
    
    'Encontra a aplicação no COM+
    For Each objCOMAdminCatalogObject In objCOMAdminCatalogCollection
        If objCOMAdminCatalogObject.Name = strNomeAplicacao Then
            Exit For
        End If
    Next

    'Obtém os componentes da aplicação
    Set objCOMComponentes = objCOMAdminCatalogCollection.GetCollection("Components", objCOMAdminCatalogObject.Key)
    objCOMComponentes.Populate
        
    lngMaximo = objCOMComponentes.Count
    lngControle = 0
    For Each objCOMComponente In objCOMComponentes
    
        lngControle = lngControle + 1
        Call fgBarraStatus(lngMaximo, lngControle, "Configurando Object Constructor...")
    
        strNomeComponente = Split(objCOMComponente.Name, ".")(0)
        strNomeClasse = Split(objCOMComponente.Name, ".")(1)
        
        With pxmlDomDescritor.documentElement
            
            Set xmlNode = .selectSingleNode("Repeat_Class_ObjectConstructor/Grupo_ObjectConstructor/" & strNomeComponente & "[Class = '" & strNomeClasse & "']")
            
            If Not xmlNode Is Nothing Then
                objCOMComponente.Value("ConstructionEnabled") = True
                strComplemento = xmlNode.selectSingleNode("Class/@ConstructorString").Text
                objCOMComponente.Value("ConstructorString") = strComplemento
                strComplemento = "(Constructor Ativado):" & strComplemento
            Else
                strComplemento = "(Constructor Desativado)"
            End If
        End With
        
        objCOMComponentes.SaveChanges
    
        fgAdicionarLog objCOMComponente.Name, "Constructor Configurado", strComplemento
    
    Next objCOMComponente
    
'Finalizações
    Set objCOMAdminCatalog = Nothing
    Set objCOMAdminCatalogCollection = Nothing
    Set objCOMComponentes = Nothing

Exit Sub
ErrorHandler:
    
    Set objCOMAdminCatalog = Nothing
    Set objCOMAdminCatalogCollection = Nothing
    Set objCOMAdminCatalogObject = Nothing
    
    Set objCOMComponentes = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

