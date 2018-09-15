VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMIU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Empresa            : Regerbanc - Partticipações , Negócios e Serviços LTDA\r\nComponente         : BJMIU\r\nClasse             : clsMIU\r\nData Criação       : 07/10/2002\r\nObjetivo           :\r\n                   :\r\nAnalista           : Adilson G. Damasceno/Carlos Fortes/Marcelo Kida\r\n\r\nProgramador        : Carlos Fortes\r\nData               : 07/10/2002\r\n\r\nData Teste         :\r\nAutor              :\r\n\r\nData Alteração     : 23/07/2003\r\nAutor              : Cassiano Nicolosi\r\nObjetivo           : Implementação da função ""ObterMapaFiltro"""
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3D9F1C270370"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e outras classes genéricas do objeto A7MIU.dll.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio As Long
Attribute lngCodigoErroNegocio.VB_VarDescription = "Variável utilizada para tratamento de erros"
Private intNumeroSequencialErro As Integer

Private objContext    As COMSVCSLib.ObjectContext

Implements COMSVCSLib.ObjectControl

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

'Montar mapa de navegação de acordo com sua configuração.
Private Function flMontarMapaNavegacao(ByRef pxmlMapaNavegacao As MSXML2.DOMDocument40, _
                                       ByVal pstrFuncionalidade As String) As Boolean
                                   
Dim xmlDomNode                             As MSXML2.IXMLDOMNode
Dim strXML                                 As String
Dim blnPrimeiroLerTodos                    As Boolean

On Error GoTo ErrorHandler
    
    blnPrimeiroLerTodos = True
    
    Call fgAppendNode(pxmlMapaNavegacao, pstrFuncionalidade, "Grupo_Propriedades", "", "")
    Call fgAppendNode(pxmlMapaNavegacao, pstrFuncionalidade, "Grupo_Dados", "", "")

    For Each xmlDomNode In pxmlMapaNavegacao.selectNodes("//Grupo_Objetos/*")
        Select Case xmlDomNode.attributes.getNamedItem("Metodo").Text
            Case "Propriedades"
                
                strXML = flObterPropriedades(xmlDomNode.Text)
                Call fgAppendXML(pxmlMapaNavegacao, "Grupo_Propriedades", strXML)
            
            Case "LerTodos"
                
                strXML = flLerTodos(xmlDomNode.Text)
                 
                If strXML <> "" Then
                    fgAppendXML pxmlMapaNavegacao, "Grupo_Dados", strXML
                End If
            Case "ObterDominioSPB"
                
                strXML = flObterDominioSPB(xmlDomNode.attributes.getNamedItem("NomeDominio").Text, _
                                           xmlDomNode.attributes.getNamedItem("NomeXML").Text)
                 
                If strXML <> "" Then
                    fgAppendXML pxmlMapaNavegacao, "Grupo_Dados", strXML
                End If
        End Select
        
    Next xmlDomNode

    fgRemoveNode pxmlMapaNavegacao, "Grupo_Objetos"
    
    flMontarMapaNavegacao = True

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarMapaNavegacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Executar o método LerTodos do objeto informado.
Private Function flLerTodos(ByVal pstrNomeObjeto As String) As String

Dim objLerTodos                             As Object

On Error GoTo ErrorHandler

    Set objLerTodos = CreateObject(pstrNomeObjeto)
    flLerTodos = objLerTodos.LerTodos()
    Set objLerTodos = Nothing

    Exit Function

ErrorHandler:
    
    Set objLerTodos = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLerTodos", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Obter todos os itens de um domínio SPB
Private Function flObterDominioSPB(ByVal pstrNomeDominio As String, _
                                   ByVal pstrTagRetorno As String) As String

Dim objDominio                            As Object

On Error GoTo ErrorHandler

    Set objDominio = CreateObject("A8MIU.clsA6A7A8Funcoes")
    
    flObterDominioSPB = objDominio.ObterDominioSPB(pstrNomeDominio, pstrTagRetorno)
    
    Set objDominio = Nothing

    Exit Function

ErrorHandler:
    
    Set objDominio = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterDominioSPB", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Executar o método ObterPropriedades do objeto informado.
Private Function flObterPropriedades(ByVal pstrNomeObjeto As String) As String

Dim objPropriedades                         As Object

On Error GoTo ErrorHandler

    Set objPropriedades = CreateObject(pstrNomeObjeto)
    flObterPropriedades = objPropriedades.ObterPropriedades()
    Set objPropriedades = Nothing

    Exit Function

ErrorHandler:
    
    Set objPropriedades = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterPropriedades", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Controlar o roteamento e execução dos métodos incluir, alterar, excluir, ler, lertodos para o objeto informado utilizando as classes clsConsulta e clsTransacao do A7Miu.
Public Function Executar(ByVal pstrPropriedades As String, _
                         ByRef pvntCodErro As Variant, _
                         ByRef pvntMensagemErro As Variant) As String
    
Dim xmlPropriedades                         As MSXML2.DOMDocument40
Dim strOperacao                             As String
Dim objTransacao                            As Object 'MIU.clsTransacao
Dim objConsulta                             As Object 'MIU.clsConsulta
Dim objExecutarGenericoRemoto               As Object 'A8LQS.clsExecutarGenericoRemoto

    On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlPropriedades = CreateObject("MSXML2.DOMDocument.4.0")

    pstrPropriedades = Replace(pstrPropriedades, "'", vbNullString)
    
    xmlPropriedades.preserveWhiteSpace = True

    If Not xmlPropriedades.loadXML(pstrPropriedades) Then
        '3 - XML Inválido.
        lngCodigoErroNegocio = 3
        GoTo ErrorHandler
    End If

    strOperacao = xmlPropriedades.selectSingleNode("//*/@Operacao").Text

    If InStr(1, UCase(strOperacao), "INCLUIR") <> 0 Or _
       InStr(1, UCase(strOperacao), "ALTERAR") <> 0 Or _
       InStr(1, UCase(strOperacao), "EXCLUIR") <> 0 Then
            
        Set objTransacao = CreateObject(App.EXEName & ".clsTransacao")
        Executar = objTransacao.Executar(xmlPropriedades)
        Set objTransacao = Nothing
            
    ElseIf InStr(1, UCase(strOperacao), "LER") <> 0 Or _
           InStr(1, UCase(strOperacao), "OBTER") <> 0 Then
            
        Set objConsulta = CreateObject(App.EXEName & ".clsConsulta")
        Executar = objConsulta.Executar(xmlPropriedades)
        Set objConsulta = Nothing
            
    ElseIf InStr(1, UCase(strOperacao), "LOTE") <> 0 Then
            
        Executar = flProcessarLote(xmlPropriedades, pvntCodErro, pvntMensagemErro)
            
    Else
            
        ' 4 - Operação Inválida
        lngCodigoErroNegocio = 4
        GoTo ErrorHandler
    
    End If

    Set xmlPropriedades = Nothing

    Exit Function

ErrorHandler:
    Set objConsulta = Nothing
    Set objTransacao = Nothing
    
    If xmlPropriedades.parseError.errorCode <> 0 Then
        Err.Number = xmlPropriedades.parseError.errorCode
        Err.Description = xmlPropriedades.parseError.reason
        Err.Source = App.EXEName & "-" & TypeName(Me) & "-" & "flMontarMapaNavegacao Function"
    End If
    
    Set xmlPropriedades = Nothing
    
    ' Comentado devido à inclusão de tratamento de erro do Soap
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarMapaNavegacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Processamento em lote genérico iniciado pelo método Executar.
Private Function flProcessarLote(ByVal pxmlProcessamento As MSXML2.DOMDocument40, _
                                 ByRef pvntCodErro As Variant, _
                                 ByRef pvntMensagemErro As Variant) As String

Dim objTransacao                            As Object 'MIU.clsTransacao

Dim xmlDocProcessar                         As MSXML2.DOMDocument40
Dim xmlItemProc                             As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim xmlControleResultado                    As MSXML2.DOMDocument40
Dim xmlControleErro                         As MSXML2.DOMDocument40

Dim strxmlErro                              As String
Dim lngCodErro                              As Long
Dim intTipoErro                             As Integer
Dim strErro                                 As String

Dim blnIgnoraGradeHorario                   As Boolean

    On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlControleErro = CreateObject("MSXML2.DOMDocument.4.0")

    Set xmlControleResultado = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlControleResultado, "", "Repeat_ControleErro", "")

    Set xmlDocProcessar = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlDocProcessar.loadXML(pxmlProcessamento.selectNodes("//Grupo_Filtros/*").Item(0).xml)

    For Each xmlDomNode In xmlDocProcessar.documentElement.childNodes
        Call fgAppendNode(xmlControleResultado, "Repeat_ControleErro", "Grupo_ControleErro", "")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TipoConfirmacao", 0, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Operacao", vbNullString, "Repeat_ControleErro")

        Set xmlItemProc = CreateObject("MSXML2.DOMDocument.4.0")
        Set objTransacao = CreateObject(App.EXEName & ".clsTransacao")
        
        Call xmlItemProc.loadXML(xmlDomNode.xml)
        strxmlErro = objTransacao.ProcessarItemLote(xmlItemProc)
        
        Set objTransacao = Nothing
        Set xmlItemProc = Nothing
            
        If strxmlErro <> vbNullString Then
            Call xmlControleErro.loadXML(strxmlErro)
            intTipoErro = xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/ErrorType").Text
            lngCodErro = Val(xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/Number").Text)

            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "ErroNegocioEspecifico", 0, "Repeat_ControleErro")

            strErro = xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/Description").Text

            'Incrementa o controle de erros com o STATUS do processo (ERRO)
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TipoErro", intTipoErro, "Repeat_ControleErro")
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CodigoErro", lngCodErro, "Repeat_ControleErro")
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Status", strErro, "Repeat_ControleErro")
        Else
            lngCodErro = 0

            'Incrementa o controle de erros com o STATUS do processo (OK)
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CodigoErro", lngCodErro, "Repeat_ControleErro")
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Status", "OK", "Repeat_ControleErro")
        End If
    Next

    flProcessarLote = xmlControleResultado.xml

    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing
    Set xmlDocProcessar = Nothing

    Exit Function

ErrorHandler:
    Set objTransacao = Nothing
    Set xmlDocProcessar = Nothing
    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessarLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

'Obter o mapa do filtro para a funcinalidade informada.
Public Function ObterMapaFiltro(ByVal penumSistemaSLCC As Integer, _
                                ByVal pstrNomeFormulario As String) As String

Dim xmlMapaFiltro                           As MSXML2.DOMDocument40
Dim objMapaFiltro                           As Object  'MIU.clsMapaFiltro

    On Error GoTo ErrorHandler

    Set xmlMapaFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMapaFiltro = CreateObject(App.EXEName & ".clsMapaFiltro")

    If objMapaFiltro.ObterMapaFiltro(xmlMapaFiltro, penumSistemaSLCC, pstrNomeFormulario) Then
        '3 - XML Inválido.
        lngCodigoErroNegocio = 3
        GoTo ErrorHandler
    End If

    ObterMapaFiltro = xmlMapaFiltro.xml

    Set objMapaFiltro = Nothing
    Set xmlMapaFiltro = Nothing

    Exit Function

ErrorHandler:
    
    Set xmlMapaFiltro = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMapaFiltro", lngCodigoErroNegocio, intNumeroSequencialErro, "Funcionalidade: " & pstrNomeFormulario)

End Function

'Obter o mapa de navegação para a funcinalidade informada.
Public Function ObterMapaNavegacao(ByVal penumSistemaSLCC As Integer, _
                                   ByVal pstrFuncionalidade As String, _
                                   ByRef pvntCodErro As Variant, _
                                   ByRef pvntMensagemErro As Variant) As String

Dim xmlMapaNavegacao                        As MSXML2.DOMDocument40
Dim objMapaNavegacao                        As Object 'MIU.clsMapaNavegacao

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlMapaNavegacao.validateOnParse = False
    xmlMapaNavegacao.resolveExternals = False
    
    Set objMapaNavegacao = CreateObject(App.EXEName & ".clsMapaNavegacao")

    If objMapaNavegacao.ObterMapaNavegacao(xmlMapaNavegacao, penumSistemaSLCC, pstrFuncionalidade, pvntCodErro, pvntMensagemErro) Then
        '3 - XML Inválido.
        lngCodigoErroNegocio = 3
        GoTo ErrorHandler
    End If

    Set objMapaNavegacao = Nothing

    flMontarMapaNavegacao xmlMapaNavegacao, pstrFuncionalidade

    ObterMapaNavegacao = xmlMapaNavegacao.xml

    Set xmlMapaNavegacao = Nothing

    Exit Function

ErrorHandler:
    
    Set xmlMapaNavegacao = Nothing

    'Comentado devido à inclusão de tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMapaNavegacao", lngCodigoErroNegocio, intNumeroSequencialErro, "Funcionalidade: " & pstrFuncionalidade)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

'Obter a data/hora do servidor e retornar no formato informado.
Public Function DataHoraServidor(ByVal pFormato As Integer, _
                                 ByRef pvntCodErro As Variant, _
                                 ByRef pvntMensagemErro As Variant) As String

Dim strDataHoraServidor                     As String

On Error GoTo ErrorHandler
  
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Select Case pFormato
         Case enumFormatoDataHora.Data
             strDataHoraServidor = CStr(Format(Date, "DD/MM/YYYY"))
         Case enumFormatoDataHora.Hora
             strDataHoraServidor = CStr(Format(Time, "HH:mm:ss"))
         Case enumFormatoDataHora.DataHora
             strDataHoraServidor = CStr(Format(Now, "DD/MM/YYYY HH:mm:ss"))
    End Select
    
    DataHoraServidor = strDataHoraServidor

Exit Function
ErrorHandler:
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "DataHoraServidor", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

