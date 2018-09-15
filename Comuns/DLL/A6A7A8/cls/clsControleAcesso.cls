VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControleAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Objetivo       : Controle de Acesso e Funções de Usuários
'                 Controle de Segração de dados de usuários

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Public Enum enumTipoErro
    Aplicacao = 1
    Negocio = 2
End Enum

Private strConstruct                        As String
Private Const lngOK                         As Long = 0

Implements COMSVCSLib.IObjectConstruct

'--------------------------------------------------------------------------------------
'strConstruct = 'OFF'
'Desabilita a chamada do componente do sistema MBS.('OFF' desabiltado)
'Caso a chamada estaje desabilitado as Funções e Grupos de Usuarios
'   são recuperados das tabelas a8.MBS_FUNCAO e A8.MBS_GRUPO respectivamente.
'--------------------------------------------------------------------------------------

Private Sub IObjectConstruct_Construct(ByVal pobjContruct As Object)
    
On Error GoTo ErrorHandler
    
    strConstruct = pobjContruct.ConstructString

ErrorHandler:

End Sub

'--------------------------------------------------------------------------------------
'Adicionar o usuario a lista de usuarios logados no sistema
'--------------------------------------------------------------------------------------

Public Function AdicionaUsuario(ByVal pstrUsuarioRede As String, _
                                ByVal pstrUsuarioSistema As String, _
                                ByVal pstrEstacaoTrabalho As String, _
                                ByVal pblnForcaIncluisao As Boolean, _
                                ByRef plngCodigoErro As Long, _
                                ByRef pintNumeroSequencial As Integer) As Boolean
'Inclui o usuário no Recordset

Dim strUsuarioSistema                       As String

On Error GoTo ErrorHandler
    
    pstrUsuarioRede = UCase(Trim(pstrUsuarioRede))
    pstrUsuarioSistema = UCase(Trim(pstrUsuarioSistema))
    pstrEstacaoTrabalho = UCase(Trim(pstrEstacaoTrabalho))
        
    'Verifica se o usuário já está logado
    If Not ValidaUsuario(pstrUsuarioRede, _
                         pstrUsuarioSistema, _
                         pstrEstacaoTrabalho, _
                         plngCodigoErro, _
                         pintNumeroSequencial) Then
                         
        'Forca a Inclusao do Usuário
        '   Esse artificio será utilizado quando o usuário já estiver logado
        '   na estação e utilizar a função de troca de usuário no front-end.
        '   Nesse caso, a função ValidaUsuario retorna "True"
        If pblnForcaIncluisao Then
            
            'O erro 240 indica usuário já logado porém,
            '   mesmo usuário de rede e mesma estação. Somente o usuário do sistema
            '   esta sendo trocado
            
            'É necessário dar logoff no usuário de sistema anterior.
            If plngCodigoErro = 15 Then
                
                plngCodigoErro = 0
                
'                'Guarda o login que estava sendo utilizado. Como não existe transação, o login
'                '   antigo será removido e caso aconteça algum erro, será adicionado novamente,
'                '   simulando um Roolback
'                If ObterUsuarioSistema(pstrUsuarioRede, _
'                                       strUsuarioSistema, _
'                                       plngCodigoErro, _
'                                       pintNumeroSequencial) Then
                            
                    'Remove o login antigo
                    If RemoveUsuario(pstrUsuarioRede, plngCodigoErro) Then
                        
                        'Valida o novo usuário a ser logado
                        If Not ValidaUsuario(pstrUsuarioRede, _
                                             pstrUsuarioSistema, _
                                             pstrEstacaoTrabalho, _
                                             plngCodigoErro, _
                                             pintNumeroSequencial) Then
                            
                            'Ocorreram erros na consistência
                            'O login antigo será efetivado
                            Call flAdicionaUsuario(pstrUsuarioRede, pstrUsuarioSistema, pstrEstacaoTrabalho)
                            
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                'Else
                '    Exit Function
                'End If
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    'Adiciona o usuário Shared Property Memory
    Call flAdicionaUsuario(pstrUsuarioRede, pstrUsuarioSistema, pstrEstacaoTrabalho)

    AdicionaUsuario = True

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AdicionaUsuario", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------
'Remover o usuario da lista de usuarios logados no sistema
'--------------------------------------------------------------------------------------

Public Function RemoveUsuario(ByVal pstrUsuarioRede As String, _
                              ByRef plngCodigoErro As Long) As Boolean
    
'Remove o usuário do controle de acesso
Dim xmlControleAcesso                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strControleAcesso                       As String

On Error GoTo ErrorHandler
    
    pstrUsuarioRede = UCase(Trim(pstrUsuarioRede))
    
    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")
           
    strControleAcesso = flGetControleAcesso
           
    If strControleAcesso = vbNullString Then
        RemoveUsuario = True
        Exit Function
    End If
               
    If Not xmlControleAcesso.loadXML(strControleAcesso) Then
        RemoveUsuario = True
        Exit Function
    End If
               
    Set xmlNode = xmlControleAcesso.selectSingleNode("Repet_Usuario/Grupo_Usuario[UsuarioRede='" & pstrUsuarioRede & "']")
           
    If Not xmlNode Is Nothing Then
        fgRemoveNode xmlControleAcesso, "Repet_Usuario/Grupo_Usuario[UsuarioRede='" & pstrUsuarioRede & "']"
        flSetControleAcesso xmlControleAcesso.xml
    End If
            
    Set xmlControleAcesso = Nothing
        
    RemoveUsuario = True

    Exit Function
    
ErrorHandler:
    
    Set xmlControleAcesso = Nothing
    
    plngCodigoErro = Err.Number
    
End Function

'--------------------------------------------------------------------------------------
'Verificar se o Usuário não esta logado em outra estação de Trabalho.
'--------------------------------------------------------------------------------------

Private Function ValidaUsuario(ByVal pstrUsuarioRede As String, _
                               ByVal pstrUsuarioSistema As String, _
                               ByVal pstrEstacaoTrabalho As String, _
                               ByRef plngCodigoErro As Long, _
                               ByRef pintNumeroSequencialErro As Integer) As Boolean



Dim xmlControleAcesso                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strControleAcesso                       As String

On Error GoTo ErrorHandler
    
    pstrUsuarioRede = UCase(Trim(pstrUsuarioRede))
    
    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")
           
    strControleAcesso = flGetControleAcesso
               
    If Not xmlControleAcesso.loadXML(strControleAcesso) Then
        'Se xmlControleAcesso is nothing retorna true para que o usuario seja incluido no controle de acesso
        ValidaUsuario = True
        Exit Function
    End If
    
    With xmlControleAcesso
        
        Set xmlNode = .selectSingleNode("Repet_Usuario/Grupo_Usuario[UsuarioRede='" & pstrUsuarioRede & "']")
        
        If xmlNode Is Nothing Then
            
            'Se não encotrou o usuário, verifica se o sistema não esta ativo na estacao de trabalho
            Set xmlNode = .selectSingleNode("Repet_Usuario/Grupo_Usuario[EstacaoTrabalho='" & pstrEstacaoTrabalho & "']")
            
            If xmlNode Is Nothing Then
                ValidaUsuario = True
                Exit Function
            Else
                'É necessário dar logoff no usuário de sistema anterior.
                plngCodigoErro = 15
                pintNumeroSequencialErro = 1
                ValidaUsuario = False
            End If
                        
        End If
    
        'Se encontrou o usuário e a estação for a mesma, e o login for o mesmo libera o usuário
        If UCase(Trim(pstrEstacaoTrabalho)) = UCase(Trim(xmlNode.selectSingleNode("EstacaoTrabalho").Text)) Then
            If UCase(Trim(pstrUsuarioSistema)) = UCase(Trim(xmlNode.selectSingleNode("UsuarioSistema").Text)) Then
                RemoveUsuario pstrUsuarioRede, plngCodigoErro
                ValidaUsuario = True
            Else
                'É necessário dar logoff no usuário de sistema anterior.
                plngCodigoErro = 15
                pintNumeroSequencialErro = 2
                ValidaUsuario = False
            End If
        
        Else 'Se a estação de trabalho for a diferente,não permitir a inclusão do usuário
            
            'Usuário logado em outra estação.
            plngCodigoErro = 17
            ValidaUsuario = False
        End If
    
    End With
        
    Set xmlControleAcesso = Nothing
    
    Exit Function
    
ErrorHandler:
    
    Set xmlControleAcesso = Nothing
    
    plngCodigoErro = Err.Number

End Function

'--------------------------------------------------------------------------------------
'Verificar se o Usuário está logado no sistema
'-------------------------------------------------------------------------------------

Public Function PesquisaUsuario(ByVal pstrUsuarioRede As String, _
                                ByRef pstrRetornoPesquisa As String, _
                                ByRef plngCodigoErro As Long, _
                                ByRef pintNumerSequenciaErro As Integer) As Boolean

Dim xmlControleAcesso                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strControleAcesso                       As String

On Error GoTo ErrorHandler
    
    pstrUsuarioRede = UCase(Trim(pstrUsuarioRede))
    
    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")
           
    strControleAcesso = flGetControleAcesso
           
    If Not xmlControleAcesso.loadXML(strControleAcesso) Then
        'Usuário não logado no sistema.
        plngCodigoErro = 18
        pintNumerSequenciaErro = 1
        PesquisaUsuario = False
        Exit Function
    End If
           
    With xmlControleAcesso
        
        Set xmlNode = .selectSingleNode("Repet_Usuario/Grupo_Usuario[UsuarioRede='" & pstrUsuarioRede & "']")
        
        If Not xmlNode Is Nothing Then
            pstrRetornoPesquisa = Format(xmlNode.selectSingleNode("UsuarioRede").Text, "@@@@@@@@") & _
                                  Format(xmlNode.selectSingleNode("UsuarioSistema").Text, "@@@@@@@@") & _
                                  xmlNode.selectSingleNode("EstacaoTrabalho").Text
            PesquisaUsuario = True
        Else
            pstrRetornoPesquisa = vbNullString
            
            'Usuário não logado no sistema.
            plngCodigoErro = 18
            pintNumerSequenciaErro = 2
            PesquisaUsuario = False
            Exit Function
        End If
    
    End With
        
    Set xmlControleAcesso = Nothing

    Exit Function
    
ErrorHandler:
    
    Set xmlControleAcesso = Nothing
    
    plngCodigoErro = Err.Number

End Function

'--------------------------------------------------------------------------------------
'Listar todos os usuários logaods no sistema
'-------------------------------------------------------------------------------------


Public Function ListarUsuarios() As String

'Lista todos os Registros de Ativos no Controle de Acesso

Dim xmlControleAcesso                       As MSXML2.DOMDocument40
Dim strControleAcesso                       As String

On Error GoTo ErrorHandler
    
    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")
     
    strControleAcesso = flGetControleAcesso
        
    If xmlControleAcesso.loadXML(strControleAcesso) Then
        ListarUsuarios = xmlControleAcesso.xml
    End If
        
    Set xmlControleAcesso = Nothing
    
    Exit Function
    
ErrorHandler:
    Set xmlControleAcesso = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Public Function ObterUsuarioSistema(ByVal pstrUsuarioRede As String, _
                                    ByRef pstrRetornoPesquisa As String, _
                                    ByRef plngCodigoErro As Long, _
                                    ByRef pintNumeroSequenciaErro As Integer) As Boolean

Dim xmlControleAcesso                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strControleAcesso                       As String

On Error GoTo ErrorHandler
    
    pstrUsuarioRede = UCase(Trim(pstrUsuarioRede))
    
    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")
           
    strControleAcesso = flGetControleAcesso
           
    If Not xmlControleAcesso.loadXML(strControleAcesso) Then
        pstrRetornoPesquisa = "SISTEMA"
        Exit Function
    End If
           
    With xmlControleAcesso
        Set xmlNode = .selectSingleNode("Repet_Usuario/Grupo_Usuario[UsuarioRede='" & pstrUsuarioRede & "']")
        
        If Not xmlNode Is Nothing Then
            pstrRetornoPesquisa = xmlNode.selectSingleNode("UsuarioSistema").Text
        Else
            pstrRetornoPesquisa = vbNullString
            
            'Usuário não logado no sistema.
            plngCodigoErro = 18
            pintNumeroSequenciaErro = 2
            Exit Function
        End If
    End With
        
    Set xmlControleAcesso = Nothing
        
    ObterUsuarioSistema = True

    Exit Function
    
ErrorHandler:
    
    Set xmlControleAcesso = Nothing
    
    plngCodigoErro = Err.Number

End Function

'--------------------------------------------------------------------------------------
'Obter o nome da etsação de trabalho do usuário logado no sistema
'-------------------------------------------------------------------------------------


Public Function ObterEstacaoTrabalhoUsuario(ByVal pstrUsuarioRede As String, _
                                            ByRef pstrRetornoEstacaoTrabalho As String, _
                                            ByRef plngCodigoErro As Long, _
                                            ByRef pintNumeroSequencialErro As Integer) As Boolean

Dim xmlControleAcesso                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strControleAcesso                       As String

On Error GoTo ErrorHandler
    
    pstrUsuarioRede = UCase(Trim(pstrUsuarioRede))
    
    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")
    
    strControleAcesso = flGetControleAcesso
    
    If Not xmlControleAcesso.loadXML(strControleAcesso) Then
        pstrRetornoEstacaoTrabalho = "SERVIDOR"
        ObterEstacaoTrabalhoUsuario = True
        Exit Function
    End If
     
    With xmlControleAcesso
  
        Set xmlNode = .selectSingleNode("Repet_Usuario/Grupo_Usuario[UsuarioRede='" & pstrUsuarioRede & "']")
  
        If Not xmlNode Is Nothing Then
            pstrRetornoEstacaoTrabalho = xmlNode.selectSingleNode("EstacaoTrabalho").Text
            ObterEstacaoTrabalhoUsuario = True
        Else
            pstrRetornoEstacaoTrabalho = vbNullString
            ObterEstacaoTrabalhoUsuario = False
      
            'Usuário não logado no sistema.
            plngCodigoErro = 18
            pintNumeroSequencialErro = 2
            Exit Function
        End If
    
    End With
  
    Set xmlControleAcesso = Nothing

    Exit Function
    
ErrorHandler:
    
    Set xmlControleAcesso = Nothing
    
    pstrRetornoEstacaoTrabalho = ""
    plngCodigoErro = Err.Number

End Function

'--------------------------------------------------------------------------------------
'Adicionar o usuário a lista de usuários logados
'
'   A lista de usuários logados é armazenado no 'Shared Property Memory' do COM+ em XML.
'
'<Repet_Usuario>
'   <Grupo_Usuario>
'      <UsuarioRede>C.FORTES</UsuarioRede>                          <-- ID do Usuario
'      <UsuarioSistema>C.FORTES</UsuarioSistema>
'      <EstacaoTrabalho>WXP-03</EstacaoTrabalho>                    <-- Nome da estação de trabalho
'      <Repeat_SegregaAcessoDados>
'         <Grupo_TipoBackOffice>
'            <TipoBackOffice>2</TipoBackOffice>                     <-- Tipo de BackOffice Usuário
'         </Grupo_TipoBackOffice>
'         <Repeat_LocalLiquidacao>
'            <Grupo_LocalLiquidacao>
'               <LocalLiquidacao>17</LocalLiquidacao>               <-- Local de Liquiudação
'               <LocalLiquidacao>11</LocalLiquidacao>
'            </Grupo_LocalLiquidacao>
'         </Repeat_LocalLiquidacao>
'         <Repeat_GrupoVeiculoLegal>
'            <Grupo_GrupoVeiculoLegal>
'               <GrupoVeiculoLegal>5</GrupoVeiculoLegal>            <-- Grupo de Veiculo Legal
'               <GrupoVeiculoLegal>25</GrupoVeiculoLegal>
'           </Grupo_GrupoVeiculoLegal>
'         </Repeat_GrupoVeiculoLegal>
'         <Repeat_GrupoUsuario>
'            <Grupo_GrupoUsuario>
'               <GrupoUsuario>1</GrupoUsuario>                      <-- Grupo de Usuário
'            </Grupo_GrupoUsuario>
'         </Repeat_GrupoUsuario>
'      </Repeat_SegregaAcessoDados>
'   </Grupo_Usuario>
'</Repet_Usuario>
'
'-------------------------------------------------------------------------------------



Public Function flAdicionaUsuario(ByVal pstrUsuarioRede As String, _
                                   ByVal pstrUsuarioSistema As String, _
                                   ByVal pstrEstacaoTrabalho As String) As Boolean

Dim xmlControleAcesso                       As MSXML2.DOMDocument40
Dim xmlControleAcessoAux                    As MSXML2.DOMDocument40
Dim strControleAcesso                       As String
Dim strGrupoAcessoDados                     As String
Dim xmlNode                                 As String

On Error GoTo ErrorHandler
                
    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlControleAcessoAux = CreateObject("MSXML2.DOMDocument.4.0")
    
    strControleAcesso = flGetControleAcesso
        
    strGrupoAcessoDados = ObterGruposAcessoDadosPorUsuario(pstrUsuarioRede, 999)
    
    If strControleAcesso = vbNullString Then
        Call fgAppendNode(xmlControleAcessoAux, "", "Repet_Usuario", "")
        Call fgAppendNode(xmlControleAcessoAux, "Repet_Usuario", "Grupo_Usuario", "")
        Call fgAppendNode(xmlControleAcessoAux, "Grupo_Usuario", "UsuarioRede", pstrUsuarioRede)
        Call fgAppendNode(xmlControleAcessoAux, "Grupo_Usuario", "UsuarioSistema", pstrUsuarioSistema)
        Call fgAppendNode(xmlControleAcessoAux, "Grupo_Usuario", "EstacaoTrabalho", pstrEstacaoTrabalho)
        
        If strGrupoAcessoDados <> vbNullString Then
            Call fgAppendXML(xmlControleAcessoAux, "Grupo_Usuario", strGrupoAcessoDados)
        End If
        
        Call flSetControleAcesso(xmlControleAcessoAux.xml)
    Else
        xmlControleAcesso.loadXML strControleAcesso
        Call fgAppendNode(xmlControleAcessoAux, "", "Grupo_Usuario", "")
        Call fgAppendNode(xmlControleAcessoAux, "Grupo_Usuario", "UsuarioRede", pstrUsuarioRede)
        Call fgAppendNode(xmlControleAcessoAux, "Grupo_Usuario", "UsuarioSistema", pstrUsuarioSistema)
        Call fgAppendNode(xmlControleAcessoAux, "Grupo_Usuario", "EstacaoTrabalho", pstrEstacaoTrabalho)
        
        If strGrupoAcessoDados <> vbNullString Then
            Call fgAppendXML(xmlControleAcessoAux, "Grupo_Usuario", strGrupoAcessoDados)
        End If
        
        Call fgAppendXML(xmlControleAcesso, "Repet_Usuario", xmlControleAcessoAux.xml)

        Call flSetControleAcesso(xmlControleAcesso.xml)
    End If
    
    Set xmlControleAcesso = Nothing
    Set xmlControleAcessoAux = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlControleAcesso = Nothing
    Set xmlControleAcessoAux = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAdicionaUsuario", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------------------------------
'Obter o XML de controle de acesso e segração de dados armazenados no 'Shared Property Memory' do COM+ em XML.
'---------------------------------------------------------------------------------------------------

Private Function flGetControleAcesso() As String

Dim objSharedPropMem                        As A6A7A8.clsSharedPropMem

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")

    flGetControleAcesso = objSharedPropMem.GetSPMProperty("SLCC", "CONTROLEACESSO")

    Set objSharedPropMem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGetControleAcesso", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------------------------------
'Gravar o XML de controle de acesso e segração no 'Shared Property Memory' do COM+ em XML.
'---------------------------------------------------------------------------------------------------

Private Function flSetControleAcesso(ByVal pxmlControleAcesso As String) As Boolean

Dim objSharedPropMem                        As A6A7A8.clsSharedPropMem

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    
    Call objSharedPropMem.SetSPMProperty("SLCC", "CONTROLEACESSO", pxmlControleAcesso)
    
    Set objSharedPropMem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flSetControleAcesso", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------------------------------
'Obter segregação de dados de um usuario:
'---------------------------------------------------------------------------------------------------

Public Function SegregaDados(ByVal strOwnerTabela As String, _
                    Optional ByVal blnRetornaSubSelect As Boolean = True, _
                    Optional ByVal strAliasTabelaFato As String = "TabFato_Sub", _
                    Optional ByVal strAliasTabelaDominio As String = "TabDominio_Sub", _
                    Optional ByVal blnPrefixoWhere As Boolean = True, _
                    Optional ByVal blnEstabeleceJoinComTabelaDominio As Boolean = True, _
                    Optional ByVal blnFiltraTipoBackOffice As Boolean = True, _
                    Optional ByVal blnFiltraLocalLiquidacao As Boolean = True, _
                    Optional ByVal blnFiltraGrupoVeiculoLegal As Boolean = True, _
                    Optional ByVal blnFiltraGrupoUsuario As Boolean = True, _
                    Optional ByVal objDomComplementoWhere As MSXML2.DOMDocument40, _
                    Optional ByVal blnConsideraNulos As Boolean = False) As String

'-------------------------------------------------------------------------------------------------
' >>> Considera segregação de dados por:
'           Tipo de BackOffice
'           Local de Liquidação
'           Grupo de Veículo Legal
'           Grupo de Usuario
'
'     A segregação é montada nas colunas da tabela FATO informada +
'                             as colunas da tabela de DOMÍNIO (TB_VEIC_LEGA) +
'                             SQL WHERE (opcional) fornecido para a tabela FATO
'
'     A rotina pode retornar um SubSelect (Default), ou apenas um Where de restrição
'
'     O ALIAS para a tabela FATO (<< strOwnerTabela >>), bem como o alias para a
'     tabela DOMÍNIO (TB_VEIC_LEGA), podem ser informados. Se não forem, será assumido o default
'
'     As combinações de ALIAS para as tabelas serão validadas conforme abaixo:
'           strAliasTabelaFato <> vbNullString E strAliasTabelaDominio <> vbNullString      VÁLIDA
'           strAliasTabelaFato =  vbNullString E strAliasTabelaDominio =  vbNullString      VÁLIDA
'           strAliasTabelaFato <> vbNullString E strAliasTabelaDominio =  vbNullString  NÃO VÁLIDA
'           strAliasTabelaFato =  vbNullString E strAliasTabelaDominio <> vbNullString  NÃO VÁLIDA
'
'     O prefixo Where também pode ser configurado, apenas se o retorno não for um SubSelect,
'     sendo seu default = True
'           TRUE        = Prefixo Where
'           FALSE       = Prefixo And
'
'     É possível estabelecer ou não o JOIN entre as tabelas Fato e Domínio para o caso
'     de retorno de um SubSelect
'
'     É possível configurar ou não a utilização dos filtros por:
'           Tipo de BackOffice;
'           Grupo Veículo Legal;
'           Local de Liquidação  e
'           Grupo de Usuário                    (o default para os dois filtros é TRUE)
'
'     É possível enviar instruções SQL WHERE complementares para as tabelas Fato e Domínio,
'     neste caso, enviar um XML de acordo com a estrutura abaixo:
'       <Repeat_Where>
'           <Grupo_Where>
'               <Where Alias="Tab_Dominio">Coluna01 operador Conteúdo</Where>
'               <Where Alias="Tab_Dominio">Coluna02 operador Conteúdo</Where>
'               <Where Alias="Tab_Dominio">Coluna0N operador Conteúdo</Where>
'           </Grupo_Where>
'       </Repeat_Where>
'
'     IMPORTANTE: para enviar uma instrução SQL WHERE para a tabela de domínio, criar um atributo
'                 ALIAS para a TAG WHERE e informar o alias da tabela de domínio
'
'     É possível informar se restrição será apenas para cada atributo ou se a restrição será
'     pelo atributo mais coluna NULA (o default é FALSE)
'-------------------------------------------------------------------------------------------------

Dim strSQL                                  As String
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim strTipoBackOffice                       As String
Dim strLocalLiquidacao                      As String
Dim strGrupoVeiculoLegal                    As String
Dim strGrupoUsuario                         As String
Dim intCont                                 As Integer
Dim xmlDocSegregacaoDados                   As MSXML2.DOMDocument40
Dim strAliasComplementoWhere                As String

    On Error GoTo ErrorHandler

'    If gstrUsuarioRede = vbNullString Then
'        gstrUsuarioRede = fgObterUsuarioRede
'        Call SaveSetting("A6A7A8", "ControleAcesso", "RegraSegregacao", ObterRegraSegregacaoUsuario(gstrUsuarioRede))
'    End If
'
    Set xmlDocSegregacaoDados = CreateObject("MSXML2.DOMDocument.4.0")
'    Call xmlDocSegregacaoDados.loadXML(GetSetting("A6A7A8", "ControleAcesso", "RegraSegregacao"))
'    If xmlDocSegregacaoDados.xml = vbNullString Then
        Call xmlDocSegregacaoDados.loadXML(ObterRegraSegregacaoUsuario(fgObterUsuarioRede))
'    End If

    'Verifica se o retorno será um SubSelect...
    If blnRetornaSubSelect Then
        strSQL = " Select " & flObterColunasDaTabela(Split(strOwnerTabela, ".")(0), Split(strOwnerTabela, ".")(1), strAliasTabelaFato) & vbNewLine & _
                 " From   " & strOwnerTabela & " " & strAliasTabelaFato
    
        'Verifica se o JOIN deve ser estabelecido entre Fato e Domínio
        If blnEstabeleceJoinComTabelaDominio Then
            strSQL = strSQL & ", " & vbNewLine & _
                 "        A8.TB_VEIC_LEGA " & strAliasTabelaDominio & "   " & vbNewLine & _
                 " Where  " & strAliasTabelaFato & ".CO_VEIC_LEGA    = " & strAliasTabelaDominio & ".CO_VEIC_LEGA " & vbNewLine & _
                 " And    " & strAliasTabelaFato & ".SG_SIST         = " & strAliasTabelaDominio & ".SG_SIST " & vbNewLine & _
                 " And"
        Else
            strSQL = strSQL & vbNewLine & _
                " Where"
        End If
    
    '...se não, será um Where
    Else
        strSQL = " " & IIf(blnPrefixoWhere, "Where", "And")
    End If

    'Verifica se filtra por Tipo BackOffice
    If blnFiltraTipoBackOffice Then
        'Captura filtro por Tipo BackOffice
        
        'Verificacao se o usuario esta asssociado a um tipo de back office
        If Not xmlDocSegregacaoDados.documentElement.selectSingleNode("Grupo_TipoBackOffice/TipoBackOffice") Is Nothing Then
        
            strTipoBackOffice = xmlDocSegregacaoDados.documentElement.selectSingleNode("Grupo_TipoBackOffice/TipoBackOffice").Text
                    
            If blnConsideraNulos Then
                strSQL = strSQL & _
                         " (" & strAliasTabelaDominio & ".TP_BKOF    IN (" & strTipoBackOffice & ")" & vbNewLine & _
                         " OR " & strAliasTabelaDominio & ".TP_BKOF     IS NULL)" & vbNewLine & _
                         " And"
            Else
                strSQL = strSQL & _
                         " " & strAliasTabelaDominio & ".TP_BKOF    IN (" & strTipoBackOffice & ")" & vbNewLine & _
                         " And"
            End If
        Else
            'Usuário não associado ao grupo. Resultado: SEM ACESSO AOS DADOS
            strSQL = strSQL & _
                    " 'Tipo BackOffice' = 'Acesso Negado' And"
        End If
    End If
    
    'Verifica se filtra por Local de Liquidação
    If blnFiltraLocalLiquidacao Then
        'Captura filtro por Local de Liquidação
        For Each objDomNode In xmlDocSegregacaoDados.selectNodes("Repeat_SegregaAcessoDados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao/*")
            strLocalLiquidacao = strLocalLiquidacao & objDomNode.Text & ","
        Next
        
        If Trim(strLocalLiquidacao) <> "" Then
            strLocalLiquidacao = Left$(strLocalLiquidacao, Len(strLocalLiquidacao) - 1)
            
            If blnConsideraNulos Then
                strSQL = strSQL & _
                         " (" & strAliasTabelaFato & ".CO_LOCA_LIQU  IN (" & strLocalLiquidacao & ")" & vbNewLine & _
                         " OR " & strAliasTabelaFato & ".CO_LOCA_LIQU       IS NULL)" & vbNewLine & _
                         " And"
            Else
                strSQL = strSQL & _
                         " " & strAliasTabelaFato & ".CO_LOCA_LIQU  IN (" & strLocalLiquidacao & ")" & vbNewLine & _
                         " And"
            End If
        Else
            'Usuário não associado ao grupo. Resultado: SEM ACESSO AOS DADOS
            strSQL = strSQL & _
                    " 'Local Liquidação' = 'Acesso Negado' And"
        End If
    End If
    
    'Verifica se filtra por Grupo de Veículo Legal
    If blnFiltraGrupoVeiculoLegal Then
        'Captura filtro por Grupo de Veículo Legal
        For Each objDomNode In xmlDocSegregacaoDados.selectNodes("Repeat_SegregaAcessoDados/Repeat_GrupoVeiculoLegal/Grupo_GrupoVeiculoLegal/*")
            strGrupoVeiculoLegal = strGrupoVeiculoLegal & objDomNode.Text & ","
        Next
    
        If Trim(strGrupoVeiculoLegal) <> "" Then
            strGrupoVeiculoLegal = Left$(strGrupoVeiculoLegal, Len(strGrupoVeiculoLegal) - 1)
            
            If blnConsideraNulos Then
                strSQL = strSQL & _
                         " (" & strAliasTabelaDominio & ".CO_GRUP_VEIC_LEGA IN (" & strGrupoVeiculoLegal & ")" & vbNewLine & _
                         " OR " & strAliasTabelaDominio & ".CO_GRUP_VEIC_LEGA       IS NULL)" & vbNewLine & _
                         " And"
            Else
                strSQL = strSQL & _
                         " " & strAliasTabelaDominio & ".CO_GRUP_VEIC_LEGA IN (" & strGrupoVeiculoLegal & ")" & vbNewLine & _
                         " And"
            End If
        Else
            'Usuário não associado ao grupo. Resultado: SEM ACESSO AOS DADOS
            strSQL = strSQL & _
                    " 'Grupo Veículo Legal' = 'Acesso Negado' And"
        End If
    End If
    
    'Verifica se filtra por Grupo de Usuário
    If blnFiltraGrupoUsuario Then
        'Captura filtro por Grupo de Usuário
        For Each objDomNode In xmlDocSegregacaoDados.selectNodes("Repeat_SegregaAcessoDados/Repeat_GrupoUsuario/Grupo_GrupoUsuario/*")
            strGrupoUsuario = strGrupoUsuario & objDomNode.Text & ","
        Next
        
        If Trim(strGrupoUsuario) <> "" Then
            strGrupoUsuario = Left$(strGrupoUsuario, Len(strGrupoUsuario) - 1)
            
            If blnConsideraNulos Then
                strSQL = strSQL & _
                         " (" & strAliasTabelaFato & ".CO_GRUP_USUA IN (" & strGrupoUsuario & ")" & vbNewLine & _
                         " OR " & strAliasTabelaFato & ".CO_GRUP_USUA       IS NULL)" & vbNewLine & _
                         " And"
            Else
                strSQL = strSQL & _
                         " " & strAliasTabelaFato & ".CO_GRUP_USUA IN (" & strGrupoUsuario & ")" & vbNewLine & _
                         " And"
            End If
        Else
            'Usuário não associado ao grupo. Resultado: SEM ACESSO AOS DADOS
            strSQL = strSQL & _
                    " 'Grupo Usuário' = 'Acesso Negado' And"
        End If
    End If
    
    'Verifica se existe filtro complementar para as tabelas Fato e Domínio...
    If Not objDomComplementoWhere Is Nothing Then
        '...se sim, captura o filtro
        For Each objDomNode In objDomComplementoWhere.selectNodes("Repeat_Where/Grupo_Where/*")
            strAliasComplementoWhere = strAliasTabelaFato
            
            If objDomNode.Attributes.Length > 0 Then
                strAliasComplementoWhere = _
                    objDomNode.Attributes.getNamedItem("Alias").Text
            End If
            
            strSQL = strSQL & _
                    " " & strAliasComplementoWhere & "." & objDomNode.Text & vbNewLine & _
                    " And"
        Next
    End If
    
    'Remove inconsistências da instrução SQL
    If strAliasTabelaFato = vbNullString And strAliasTabelaDominio = vbNullString Then
        strSQL = Replace(strSQL, ".", vbNullString)
    End If
    
    'Remove operadores de filtro SQL não utilizados
    If UCase(Right$(Trim(strSQL), 3)) = "AND" Then strSQL = Left$(strSQL, Len(strSQL) - 4)
    If UCase(Right$(Trim(strSQL), 5)) = "WHERE" Then strSQL = Left$(strSQL, Len(strSQL) - 6)
    
    'Remove separador de linha não utilizado
    If Right$(strSQL, 2) = vbNewLine Then strSQL = Left$(strSQL, Len(strSQL) - 2)
    
    SegregaDados = strSQL

Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8Liquidacao", "SegregaDados Function", lngCodigoErroNegocio, intNumeroSequencialErro)
 
End Function

'--------------------------------------------------------------------------------------------
'Obter um registro da tabela A8.TB_CTRL_ACES_DADO (Controle de acesso a dados)
'--------------------------------------------------------------------------------------------

Public Function Ler(ByVal plngCO_GRUP_GS As Long) As ADODB.Recordset

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

    On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_GRUP_GS,             " & vbNewLine & _
             "          NO_GRUP_GS,             " & vbNewLine & _
             "          CO_INFO,                " & vbNewLine & _
             "          TP_INFO,                " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ,      " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
             "          DH_ULTI_ATLZ,           " & vbNewLine & _
             "          '' AS TP_VIGE           " & vbNewLine & _
             " FROM     A8.TB_CTRL_ACES_DADO    " & vbNewLine & _
             " WHERE    CO_GRUP_GS = " & plngCO_GRUP_GS
    
    Set Ler = fgQuerySQL(strSQL)

    Exit Function
    
ErrorHandler:
    Set Ler = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------
'Obter os grupo de acesso a dados de um usuaio :
' - Tipo de BackOffice
' - Grupos de Veiculo Legal
' - Grupos de Usuario
' - Local de Liquidação
'--------------------------------------------------------------------------------------------

Public Function ObterGruposAcessoDadosPorUsuario(ByVal pstrUsuarioRede As String, _
                                        Optional ByRef plngCodigoRetornoVerificacao As Long = 999) As String

Dim objMBS                                  As Object 'mbscom.Cseguranca
Dim objRSRetMBS                             As ADODB.Recordset
Dim objRSRetSLCC                            As ADODB.Recordset
Dim xmlControleAcessoDados                  As MSXML2.DOMDocument40
Dim intAssocTipoBackOffice                  As Integer
Dim strSQL                                  As String
Dim strFuncaoMBS                            As String
Dim lngCodigoRetorno                        As Long
Dim strMsgRetorno                           As String
Dim strCompMBS                              As String

On Error GoTo ErrorHandler

    '>>> Cria XML padrão SLCC para a segregação de acesso e dados ---------------------------------------------------
    Set xmlControleAcessoDados = CreateObject("MSXML2.DOMDocument.4.0")
    Set objRSRetMBS = CreateObject("ADODB.Recordset")
    
    'Node Principal do XML
    Call fgAppendNode(xmlControleAcessoDados, "", "Repeat_SegregaAcessoDados", "")
    
    'Node Tipo BackOffice       (1)
    Call fgAppendNode(xmlControleAcessoDados, "Repeat_SegregaAcessoDados", "Grupo_TipoBackOffice", "")
    
    'Node Local Liquidação      (N)
    Call fgAppendNode(xmlControleAcessoDados, "Repeat_SegregaAcessoDados", "Repeat_LocalLiquidacao", "")
    Call fgAppendNode(xmlControleAcessoDados, "Repeat_LocalLiquidacao", "Grupo_LocalLiquidacao", "")
    
    'Node Grupo Veículo Legal   (N)
    Call fgAppendNode(xmlControleAcessoDados, "Repeat_SegregaAcessoDados", "Repeat_GrupoVeiculoLegal", "")
    Call fgAppendNode(xmlControleAcessoDados, "Repeat_GrupoVeiculoLegal", "Grupo_GrupoVeiculoLegal", "")
    
    'Node Grupo Usuário         (N)
    Call fgAppendNode(xmlControleAcessoDados, "Repeat_SegregaAcessoDados", "Repeat_GrupoUsuario", "")
    Call fgAppendNode(xmlControleAcessoDados, "Repeat_GrupoUsuario", "Grupo_GrupoUsuario", "")
    
    'Marcelo Kida - 15/10/2003
    'Caso o MBS não esteja disponivel passar OFF no ObjectConstruct
    'Utilizar as tabelas de apoio do schema A8.MBS_GRUPO

    If Len(Trim(pstrUsuarioRede)) > 8 Then
        pstrUsuarioRede = Mid(Trim(pstrUsuarioRede), 1, 8)
    End If

    #If DebugMode <> 1 Then
    
        If UCase(Trim(strConstruct)) = "ON" Or _
           Trim(strConstruct) = vbNullString Then
            
            strCompMBS = fgGetPrivateProfileString("MBS", "Componente", App.Path & "\SLCC.INI")
            
            intNumeroSequencialErro = 1
            Set objMBS = CreateObject(strCompMBS)

            intNumeroSequencialErro = 2
            
            lngCodigoRetorno = objMBS.recTodosGruposPorCdUsuario(pstrUsuarioRede, objRSRetMBS)
            
            intNumeroSequencialErro = 3
            
            If lngCodigoRetorno <> 0 Then
                'O sistema de segurança MBS retornou um código de erro inválido.
                'lngCodigoErroNegocio = 21
                'GoTo ErrorHandler
                'Alterado por Carlos 01/12/03 17:23hs,  Andréa estava santander e em conversa com Vanderlei ficou
                'decidido que se ocorrer erro no MBS é para parar por regra de cadastro do grupo de usuário
                ObterGruposAcessoDadosPorUsuario = vbNullString
                
                'Obter Erro MBS
                Call fgObterErroMBS(lngCodigoRetorno, lngCodigoErroNegocio, strMsgRetorno)
                
                Select Case lngCodigoRetorno
                    Case 1100
                    Case Else
                        Call fgGravaArquivo("Codigo Retorno MBS:" & lngCodigoRetorno & " - " & strMsgRetorno & vbCrLf & _
                                            "clsControleAcesso.ObterGruposAcessoDadosPorUsuario")
                End Select
                
                Exit Function
            End If
    
            Set objMBS = Nothing
    
        Else
    
    #End If
        
            strSQL = " SELECT * FROM A8.MBS_GRUPO WHERE UPPER(CD_USR) = '" & UCase(pstrUsuarioRede) & "'"
            Set objRSRetMBS = fgQuerySQL(strSQL)
        
    #If DebugMode <> 1 Then
        End If
    #End If
    
    'Se o usuário não estiver cadastrado no MBS,
    'então o sistema deve parar por regra de cadastro do grupo de usuário
    If objRSRetMBS.EOF Then
        ObterGruposAcessoDadosPorUsuario = vbNullString
        Exit Function
    End If
    
    Do While Not objRSRetMBS.EOF
        If Mid(UCase(objRSRetMBS("NM_GR_USR")), 1, 2) = "A8" Then
            
            Set objRSRetSLCC = Ler(objRSRetMBS("CD_GR_USR"))
            
            If Not objRSRetSLCC.EOF Then
                Select Case objRSRetSLCC("TP_INFO")
                    Case enumTipoInformacao.TipoBackOffice
                        intAssocTipoBackOffice = intAssocTipoBackOffice + 1
                        
                        Call fgAppendNode(xmlControleAcessoDados, "Grupo_TipoBackOffice", _
                                                                  "TipoBackOffice", objRSRetSLCC("CO_INFO"))
                    Case enumTipoInformacao.LocalLiquidacao
                        Call fgAppendNode(xmlControleAcessoDados, "Grupo_LocalLiquidacao", _
                                                                  "LocalLiquidacao", objRSRetSLCC("CO_INFO"))
                    Case enumTipoInformacao.GrupoVeiculoLegal
                        Call fgAppendNode(xmlControleAcessoDados, "Grupo_GrupoVeiculoLegal", _
                                                                  "GrupoVeiculoLegal", objRSRetSLCC("CO_INFO"))
                    Case enumTipoInformacao.GrupoUsuario
                        Call fgAppendNode(xmlControleAcessoDados, "Grupo_GrupoUsuario", _
                                                                  "GrupoUsuario", objRSRetSLCC("CO_INFO"))
                End Select
            End If
            
            Set objRSRetSLCC = Nothing
        Else
            'Não faz nada
        End If
        
        objRSRetMBS.MoveNext
    Loop
    
    objRSRetMBS.Close
    
    ObterGruposAcessoDadosPorUsuario = xmlControleAcessoDados.xml
    
    Set xmlControleAcessoDados = Nothing
    Set objRSRetMBS = Nothing
    
    Exit Function

ErrorHandler:
    Set xmlControleAcessoDados = Nothing
    Set objMBS = Nothing
    Set objRSRetMBS = Nothing
    Set objRSRetSLCC = Nothing
    
    If plngCodigoRetornoVerificacao = 999 Then
        Select Case intNumeroSequencialErro
            Case 1 'Componente do Sistema MBS não encontrado(" & strcomp & ")."
                strFuncaoMBS = "Componente do Sistema MBS não encontrado(" & strCompMBS & ")."
                lngCodigoErroNegocio = 41
            Case 2 'Erro ao acessar a função do MBS.
                strFuncaoMBS = "Cseguranca.recTodosGruposPorCdUsuario"
                lngCodigoErroNegocio = 42
            Case 3 '"Código Retorno MBS: "
                If lngCodigoRetorno <> 0 Then
                    strFuncaoMBS = "Código Retorno MBS: " & lngCodigoRetorno
                End If
        End Select
    Else
        Select Case intNumeroSequencialErro
            Case 1 'Componente do Sistema MBS não encontrado(" & strcomp & ")."
                plngCodigoRetornoVerificacao = 44
            Case 2 'Erro ao acessar a função do MBS.
                plngCodigoRetornoVerificacao = 45
            Case 3 'Erro MBS não identificado.
                plngCodigoRetornoVerificacao = 46
                If lngCodigoRetorno <> 0 Then
                    Call fgObterErroMBS(lngCodigoRetorno, lngCodigoErroNegocio, strMsgRetorno)
                
                    Select Case lngCodigoRetorno
                        Case 1100
                        Case Else
                            Call fgGravaArquivo("Codigo Retorno MBS:" & lngCodigoRetorno & " - " & strMsgRetorno & vbCrLf & _
                                                "clsControleAcesso.ObterGruposAcessoDadosPorUsuario2")
                    End Select
                
                End If
        End Select
        
        ObterGruposAcessoDadosPorUsuario = vbNullString
        Exit Function
        
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterGruposAcessoDadosPorUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro, strFuncaoMBS)

End Function

'--------------------------------------------------------------------------------------------
'Obter as funções disponibilizados para o usuario
'--------------------------------------------------------------------------------------------


Public Function ObterControleAcesso(ByVal pstrSiglaSistema As String) As String

Dim objMBS                                  As Object 'mbscom.Cseguranca
Dim objRSRetMBS                             As ADODB.Recordset
Dim xmlControleAcesso                       As MSXML2.DOMDocument40
Dim strSQL                                  As String
Dim strUsuario                              As String
Dim lngCodigoRetorno                        As Long
Dim strFuncaoMBS                            As String
Dim strMsgRetorno                           As String
Dim strCompMBS                              As String

On Error GoTo ErrorHandler
    
    strUsuario = UCase(fgObterUsuarioRede())
    
    Set objRSRetMBS = CreateObject("ADODB.Recordset")
   
    'Marcelo Kida - 15/10/2003
    'Caso o MBS não esteja disponivel passar OFF no ObjectConstruct
    'Utilizar as tabelas de apoio do schema A8.MBS_FUNCAO
    #If DebugMode <> 1 Then
        If UCase(Trim(strConstruct)) = "ON" Or _
           Trim(strConstruct) = vbNullString Then
            
            strCompMBS = fgGetPrivateProfileString("MBS", "Componente", App.Path & "\SLCC.INI")
            
            intNumeroSequencialErro = 1
            Set objMBS = CreateObject(strCompMBS)
            
            intNumeroSequencialErro = 2
            lngCodigoRetorno = objMBS.recuperaHierarquiaFuncao(strUsuario, "01" & Trim(pstrSiglaSistema), objRSRetMBS)
            
            intNumeroSequencialErro = 3
            If lngCodigoRetorno <> 0 Then
                
                'Obter Erro MBS
                Call fgObterErroMBS(lngCodigoRetorno, lngCodigoErroNegocio, strMsgRetorno)
                'Call fgGravaArquivo("Usuário: " & strUsuario & vbCrLf & _
                                     "Codigo Retorno MBS: " & lngCodigoRetorno & " - " & strMsgRetorno & vbCrLf)

                GoTo ErrorHandler
            End If
    
            Set objMBS = Nothing
        Else
    #End If
         strSQL = " SELECT * FROM A8.MBS_FUNCAO WHERE UPPER(CD_USR) = '" & strUsuario & "' AND NM_SIS = '" & Trim(pstrSiglaSistema) & "'"
         Set objRSRetMBS = fgQuerySQL(strSQL)
    #If DebugMode <> 1 Then
        End If
    #End If
    
    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlControleAcesso, "", "Repeat_ControleAcesso", "")
    Call fgAppendNode(xmlControleAcesso, "Repeat_ControleAcesso", "Grupo_Acesso", "")
    
    Do While Not objRSRetMBS.EOF
        'A coluna de retorno << nm_pgn >> (Nome Página), foi utilizada para armazenar
        'o nome de menu a ser habilitado. Ex.: mnuConsultaOperacoes
        Call fgAppendNode(xmlControleAcesso, "Grupo_Acesso", "Perfil", UCase(Trim(objRSRetMBS("nm_pgn"))))
        
        objRSRetMBS.MoveNext
    Loop
    
    objRSRetMBS.Close
       
    ObterControleAcesso = xmlControleAcesso.xml
    
    Set xmlControleAcesso = Nothing
    Set objRSRetMBS = Nothing
    
Exit Function

ErrorHandler:
    
    Set xmlControleAcesso = Nothing
    Set objMBS = Nothing
    Set objRSRetMBS = Nothing
    
    Select Case intNumeroSequencialErro
        Case 1 'Componente do Sistema MBS não encontrado(mbscom.Cseguranca).
            strFuncaoMBS = "Componente do Sistema MBS não encontrado(" & strCompMBS & ")."
            lngCodigoErroNegocio = 41
        Case 2 'Erro ao acessar a função do MBS.
            strFuncaoMBS = "Cseguranca.recuperaHierarquiaFuncao"
            lngCodigoErroNegocio = 42
    End Select
        
    If intNumeroSequencialErro = 3 And lngCodigoRetorno <> 0 Then
        strFuncaoMBS = "Código Retorno MBS: " & lngCodigoRetorno
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterControleAcesso Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------
'Obter o tipo de backOffice do usuario
'--------------------------------------------------------------------------------------------


Public Function ObterTipoBackOfficeUsuarioLogado() As String

On Error GoTo ErrorHandler
    ObterTipoBackOfficeUsuarioLogado = ObterTipoBackOfficeUsuario(fgObterUsuarioRede)

Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoBackOfficeUsuarioLogado Function", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function

'--------------------------------------------------------------------------------------------
'Obter o tipo de backOffice do usuario
'--------------------------------------------------------------------------------------------

Public Function ObterTipoBackOfficeUsuario(ByVal pstrUsuarioRede As String) As String

Dim xmlControleAcesso                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strControleAcesso                       As String

On Error GoTo ErrorHandler
    
    pstrUsuarioRede = UCase(Trim(pstrUsuarioRede))
    
    #If DebugMode = 1 Then
        AdicionaUsuario fgObterUsuarioRede, fgObterUsuarioRede, fgObterEstacaoTrabalhoUsuario, True, lngCodigoErroNegocio, intNumeroSequencialErro
    #End If
    
    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")
           
    strControleAcesso = flGetControleAcesso
           
    If Not xmlControleAcesso.loadXML(strControleAcesso) Then
        'Usuário não logado no sistema.
        lngCodigoErroNegocio = 18
        intNumeroSequencialErro = 1
        GoTo ErrorHandler
    End If
           
    With xmlControleAcesso
        
        Set xmlNode = .selectSingleNode("Repet_Usuario/Grupo_Usuario[UsuarioRede='" & pstrUsuarioRede & "']")
        
        If Not xmlNode Is Nothing Then
            If Not xmlNode.selectSingleNode("Repeat_SegregaAcessoDados/Grupo_TipoBackOffice/TipoBackOffice") Is Nothing Then
                If xmlNode.selectNodes("Repeat_SegregaAcessoDados/Grupo_TipoBackOffice/TipoBackOffice").Length > 1 Then
                    ObterTipoBackOfficeUsuario = vbNullString
                    
                    'Usuário associado a mais de um Tipo Back Office
                    lngCodigoErroNegocio = 38
                    GoTo ErrorHandler
                Else
                    ObterTipoBackOfficeUsuario = xmlNode.selectSingleNode("Repeat_SegregaAcessoDados/Grupo_TipoBackOffice/TipoBackOffice").Text
                End If
            Else
                ObterTipoBackOfficeUsuario = vbNullString
                
                'Tipo Back Office não cadastrado para o usuário
                lngCodigoErroNegocio = 40
                GoTo ErrorHandler
            End If
        Else
            ObterTipoBackOfficeUsuario = vbNullString
            
            'Usuário não logado no sistema.
            lngCodigoErroNegocio = 18
            intNumeroSequencialErro = 2
            GoTo ErrorHandler
        End If
    
    End With
        
    Set xmlControleAcesso = Nothing

    Exit Function
    
ErrorHandler:
    
    Set xmlControleAcesso = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoBackOfficeUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro, "xmlControleAcesso: " & strControleAcesso)

End Function

'--------------------------------------------------------------------------------------------
'Obter a descricaodo  tipo de backOffice do usuario
'--------------------------------------------------------------------------------------------


Public Function ObterDescricaoTipoBackOffice(ByVal plngTipoBackoffice As Long) As String

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset


On Error GoTo ErrorHandler
        
    strSQL = " SELECT   TP_BKOF ," & _
             "          DE_BKOF " & _
             "  FROM    A8.TB_TIPO_BKOF " & _
             " WHERE    TP_BKOF = " & plngTipoBackoffice
                 
    Set objRS = fgQuerySQL(strSQL)
    
    If Not objRS.EOF Then
        ObterDescricaoTipoBackOffice = objRS!DE_BKOF
    End If
    
    Exit Function
    
ErrorHandler:
    
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDescricaoTipoBackOffice Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------
'Obter a regra de segregação de dados do usuario
'--------------------------------------------------------------------------------------------

Public Function ObterRegraSegregacaoUsuario(ByVal pstrUsuarioRede As String) As String

Dim xmlControleAcesso                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strControleAcesso                       As String

On Error GoTo ErrorHandler
    
    pstrUsuarioRede = UCase(Trim(pstrUsuarioRede))
            
    #If DebugMode = 1 Then
        AdicionaUsuario fgObterUsuarioRede, fgObterUsuarioRede, fgObterEstacaoTrabalhoUsuario, True, lngCodigoErroNegocio, intNumeroSequencialErro
    #End If
    
    Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0")
    
    strControleAcesso = flGetControleAcesso
    
    'Adilson Gonçalves Damasceno - 03/11/2003 - Inclusão do False e devolução do xmlControleAcesso
    xmlControleAcesso.validateOnParse = False
    xmlControleAcesso.resolveExternals = False
    
    If Not xmlControleAcesso.loadXML(strControleAcesso) Then
        'Usuário não logado no sistema.
        lngCodigoErroNegocio = 18
        intNumeroSequencialErro = 1
        GoTo ErrorHandler
    End If
           
    With xmlControleAcesso
        
        Set xmlNode = .selectSingleNode("Repet_Usuario/Grupo_Usuario[UsuarioRede='" & pstrUsuarioRede & "']")
        
        If Not xmlNode Is Nothing Then
            ObterRegraSegregacaoUsuario = xmlNode.selectSingleNode("Repeat_SegregaAcessoDados").xml
        Else
            ObterRegraSegregacaoUsuario = vbNullString
            
            'Usuário não logado no sistema.
            lngCodigoErroNegocio = 18
            intNumeroSequencialErro = 2
            GoTo ErrorHandler
        End If
    
    End With
        
    Set xmlControleAcesso = Nothing

    Exit Function
    
ErrorHandler:
    
    Set xmlControleAcesso = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterRegraSegregacaoUsuario Function", _
                                                 lngCodigoErroNegocio, _
                                                 intNumeroSequencialErro, _
                                                 "UsuarioRede: " & pstrUsuarioRede & vbNewLine & _
                                                 "xmlControleAcesso: " & xmlControleAcesso.xml)

End Function

'--------------------------------------------------------------------------------------------
'Verifica se o usuario tem permissao de executar um dos modulos do sistema
'--------------------------------------------------------------------------------------------

Public Function VerificarPermissao(ByVal pstrUsuario As String, _
                                   ByVal pstrTransacao As String) As Boolean

Dim objMBS                                  As Object
Dim lngCodigoRetorno                        As Long
Dim strMsgRetorno                           As String
Dim vntPermissao                            As Variant
Dim vntAutorizado                           As Variant
Dim strFuncaoMBS                            As String
Dim strCompMBS                              As String

On Error GoTo ErrorHandler
       
    'Marcelo Kida - 06/09/2003
    'Caso o MBS não esteja disponivel passar OFF no ObjectConstruct
    
    If InStr(1, UCase(strConstruct), "OFF") > 0 Then
       VerificarPermissao = True
       Exit Function
    End If
    
    If Trim(pstrUsuario) = vbNullString Then
        'É obrigatório informar o usuário
        lngCodigoErroNegocio = 19
        GoTo ErrorHandler
    
    ElseIf Trim(pstrTransacao) = vbNullString Then
        'É obrigatório informar a transação
        lngCodigoErroNegocio = 20
        GoTo ErrorHandler
    End If
                    
    strCompMBS = fgGetPrivateProfileString("MBS", "Componente", App.Path & "\SLCC.INI")
                
    intNumeroSequencialErro = 1
    Set objMBS = CreateObject(strCompMBS)
    
    If Len(Trim(pstrUsuario)) > 8 Then
        pstrUsuario = Mid(Trim(pstrUsuario), 1, 8)
    End If

    intNumeroSequencialErro = 2
    
    lngCodigoRetorno = objMBS.verificarAcessoFuncao(pstrUsuario, pstrTransacao, vntPermissao)
    
    intNumeroSequencialErro = 3
    
    If lngCodigoRetorno <> lngOK Then
        'Obter Erro MBS
        Call fgObterErroMBS(lngCodigoRetorno, lngCodigoErroNegocio, strMsgRetorno)
        'Call fgGravaArquivo("Usuário: " & pstrUsuario & vbCrLf & _
                             "Codigo Retorno MBS: " & lngCodigoRetorno & " - " & strMsgRetorno)
      
        GoTo ErrorHandler
    End If

    VerificarPermissao = (vntPermissao = lngOK)
        
    Set objMBS = Nothing
    
    Exit Function
ErrorHandler:
    
    Select Case intNumeroSequencialErro
        Case 1 'Componente do Sistema MBS não encontrado(mbscom.Cseguranca).
            strFuncaoMBS = "Componente do Sistema MBS não encontrado(" & strCompMBS & ")."
            lngCodigoErroNegocio = 41
        Case 2 'Erro ao acessar a função do MBS.
            strFuncaoMBS = "Cseguranca.verificarAcessoFuncao"
            lngCodigoErroNegocio = 42
    End Select
    
    If intNumeroSequencialErro = 3 And lngCodigoRetorno <> 0 Then
        strFuncaoMBS = strFuncaoMBS
    End If
    
    Set objMBS = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "VerificarPermissao Function", lngCodigoErroNegocio, intNumeroSequencialErro, strFuncaoMBS)

End Function

'--------------------------------------------------------------------------------------------
'Obter Nomes das colunas de uma tabela
'--------------------------------------------------------------------------------------------

Private Function flObterColunasDaTabela(ByVal pstrOwner As String, _
                                        ByVal pstrTabela As String, _
                               Optional ByVal pstrAlias As String) As String

Dim strSQL                                  As String
Dim Rs                                      As Recordset
Dim strRetorno                              As String

On Error GoTo ErrorHandler

    strSQL = "SELECT " & IIf(Trim$(pstrAlias) <> "", "'" & pstrAlias & ".' ||", "") & _
                     "COLUMN_NAME " & vbCrLf & _
               "FROM ALL_TAB_COLUMNS " & vbCrLf & _
              "WHERE OWNER = '" & pstrOwner & "' " & vbCrLf & _
                "AND TABLE_NAME = '" & pstrTabela & "' " & vbCrLf & _
              "ORDER BY COLUMN_NAME "

    Set Rs = fgQuerySQL(strSQL)
    
    strRetorno = Rs.GetString(adClipString, , , ", ")
    strRetorno = Left$(strRetorno, Len(strRetorno) - 2)

    Set Rs = Nothing

    flObterColunasDaTabela = strRetorno

    Exit Function
ErrorHandler:
    Set Rs = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterColunasDaTabela Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function





