Attribute VB_Name = "basA6SubReserva"
Attribute VB_Description = "Empresa        : Regerbanc\r\nComponente     : SBRSubReserva\r\nClasse         : basSBRSubReserva\r\nData Criação   : 01-05-2001 17:13\r\nObjetivo       : Funções genéricas e Atalhos para utilização de outros objetos\r\n                 dentro do mesmo Componente\r\nAnalista       : Marcelo Kida\r\n\r\nProgramador    : Marcelo Kida\r\nData           : 06/07/2003\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF9D2EF0366"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Module"

' Este componente tem como objetivo agrupar métodos utilizados na camada de negócios do sistema A6.

Option Explicit

Public Const datDataVazia                   As Date = "00:00:00"

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

' Função genérica para tratamento de erros.

Public Sub fgRaiseError(ByVal pstrComponente As String, _
                        ByVal pstrClasse As String, _
                        ByVal pstrMetodo As String, _
                        ByRef plngCodigoErroNegocio As Long, _
               Optional ByRef pintNumeroSequencialErro As Integer = 0, _
               Optional ByVal pstrComplemento As String = "")

Dim objLogErro                              As A6A7A8CA.clsLogErro
Dim ErrNumber                               As Long
Dim ErrSource                               As String
Dim ErrDescription                          As String

    Set objLogErro = CreateObject("A6A7A8CA.clsLogErro")

    objLogErro.RaiseError pstrComponente, _
                          pstrClasse, _
                          pstrMetodo, _
                          plngCodigoErroNegocio, _
                          ErrNumber, _
                          ErrSource, _
                          ErrDescription, _
                          pintNumeroSequencialErro, _
                          pstrComplemento, _
                          Err
    
    Set objLogErro = Nothing
    
    Err.Raise ErrNumber, ErrSource, ErrDescription
'**************************************************************************
End Sub

' Executa uma instrução SQL na base de dados Oracle.

Public Function fgExecuteSQL(ByVal pstrSQL As String) As Long

Dim objTransacao                            As A6A7A8CA.clsTransacao

On Error GoTo ErrHandler

    Set objTransacao = CreateObject("A6A7A8CA.clsTransacao")
    fgExecuteSQL = objTransacao.ExecuteSQL(pstrSQL)
    Set objTransacao = Nothing
    
    Exit Function
                
ErrHandler:
    
    Set objTransacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgExecuteSQL", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Executa uma sequence na base de dados Oracle.

Public Function fgExecuteSequence(ByVal pstrNomeSequence As String) As Long

Dim objTransacao                            As A6A7A8CA.clsTransacao

On Error GoTo ErrHandler
    
    Set objTransacao = CreateObject("A6A7A8CA.clsTransacao")
    fgExecuteSequence = objTransacao.ExecuteSequence(pstrNomeSequence)
    Set objTransacao = Nothing
    
    Exit Function
                
ErrHandler:
    
    Set objTransacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgExecuteSequence", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Esta função deve ser utilizada para toda Procedure
Public Function fgExecuteCMD(ByVal pstrNomeProc As String, _
                             ByVal pintPosicaoRetorno As Integer, _
                             ByRef pvntParametros() As Variant) As Variant

Dim objTransacao                            As A6A7A8CA.clsTransacao

On Error GoTo ErrHandler

    Set objTransacao = CreateObject("A6A7A8CA.clsTransacao")
    fgExecuteCMD = objTransacao.ExecuteCMD(pstrNomeProc, pintPosicaoRetorno, pvntParametros())
    Set objTransacao = Nothing

    Exit Function

ErrHandler:
    Set objTransacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgExecuteCMD", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Retorna propriedades de uma tabela.

Public Function fgPropriedades(ByVal pstrNomeXML As String, _
                               ByVal pstrSQL As String, _
                      Optional ByVal pstrNomeObjeto As String) As String

Dim objConsulta                            As A6A7A8CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    fgPropriedades = objConsulta.Propriedades(pstrNomeXML, pstrSQL, pstrNomeObjeto)
    Set objConsulta = Nothing

    Exit Function

ErrHandler:
    Set objConsulta = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgPropriedades", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Executa um select na base de dados Oracle.

Public Function fgQuerySQL(ByVal pstrSQL As String) As Object

Dim objConsulta                             As A6A7A8CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    Set fgQuerySQL = objConsulta.QuerySQL(pstrSQL)
    Set objConsulta = Nothing
    
    Exit Function

ErrHandler:
    
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgQuerySQL", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Executa um select na base de dados Oracle e retorna uma string em formato XML.

Public Function fgQueryXMLLer(ByVal pstrNomeXML As String, _
                              ByVal pstrSQL As String, _
                              ByVal pstrNomeObjeto As String) As String

Dim objConsulta                             As A6A7A8CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    fgQueryXMLLer = objConsulta.QueryXMLLer(pstrNomeXML, pstrSQL, pstrNomeObjeto)
    Set objConsulta = Nothing

    Exit Function

ErrHandler:
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgQueryXMLLer", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Executa um select na base de dados Oracle e retorna uma string em formato XML.

Public Function fgQueryXMLLerTodos(ByVal pstrNomeXML As String, _
                                   ByVal pstrSQL As String, _
                                   ByVal pstrNomeObjeto As String, _
                          Optional ByVal pblnType As Boolean = True) As String

Dim objConsulta                             As A6A7A8CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    fgQueryXMLLerTodos = objConsulta.QueryXMLLerTodos(pstrNomeXML, pstrSQL, pstrNomeObjeto, pblnType)
    Set objConsulta = Nothing

    Exit Function

ErrHandler:
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgQueryXMLLerTodos", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Trata da segregação de dados a usuários.

Public Function fgSegregaDados(ByVal strOwnerTabela As String, _
                      Optional ByVal blnRetornaSubSelect As Boolean = True, _
                      Optional ByVal strAliasTabelaFato As String = "TabFato_Sub", _
                      Optional ByVal strAliasTabelaDominio As String = "TabDominio_Sub", _
                      Optional ByVal blnPrefixoWhere As Boolean = True, _
                      Optional ByVal blnEstabeleceJoinComTabelaDominio As Boolean = True, _
                      Optional ByVal blnFiltraTipoBackOffice As Boolean = True, _
                      Optional ByVal blnFiltraLocalLiquidacao As Boolean = True, _
                      Optional ByVal blnFiltraGrupoVeiculoLegal As Boolean = True, _
                      Optional ByVal blnFiltraGrupoUsuario As Boolean = True, _
                      Optional ByVal xmlDocComplementoWhere As MSXML2.DOMDocument40) As String

Dim objControleAcesso                       As A6A7A8.clsControleAcesso

    Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso")
    fgSegregaDados = objControleAcesso.SegregaDados(strOwnerTabela, _
                                                    blnRetornaSubSelect, _
                                                    strAliasTabelaFato, _
                                                    strAliasTabelaDominio, _
                                                    blnPrefixoWhere, _
                                                    blnEstabeleceJoinComTabelaDominio, _
                                                    blnFiltraTipoBackOffice, _
                                                    blnFiltraLocalLiquidacao, _
                                                    blnFiltraGrupoVeiculoLegal, _
                                                    blnFiltraGrupoUsuario, _
                                                    xmlDocComplementoWhere)

    Set objControleAcesso = Nothing
    
    Exit Function

ErrorHandler:
    Set objControleAcesso = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8Liquidacao", "fgSegregaDados Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Retorna um objeto Date com o número de dias úteis adicionados a uma data passada como parâmetro.

Public Function fgAdicionarDiasUteis(ByVal pdatData As Date, _
                                     ByVal pintQtdeDias As Integer, _
                                     ByVal plngMovimento As enumPaginacao) As Date

Dim intQtdeDiasValidos                      As Integer
Dim intIncremento                           As Integer
Dim datRetorno                              As Date
Dim vntArray()                              As Variant

Dim objA6A7A8Funcoes                        As A6A7A8.clsA6A7A8Funcoes

On Error GoTo ErrHandler

    Set objA6A7A8Funcoes = CreateObject("A6A7A8.clsA6A7A8Funcoes")
    fgAdicionarDiasUteis = objA6A7A8Funcoes.AdicionarDiasUteis(pdatData, pintQtdeDias, plngMovimento)
    Set objA6A7A8Funcoes = Nothing


    Exit Function

ErrHandler:
    Set objA6A7A8Funcoes = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgAdicionarDiasUteis", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Retorna o código do usuário logado na rede.

Public Function fgUsuarioRede() As String

Dim objUsuario                              As A6A7A8.clsUsuario

On Error GoTo ErrorHandler

    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    fgUsuarioRede = objUsuario.ObterUsuarioRede
    Set objUsuario = Nothing

    Exit Function
ErrorHandler:

    Set objUsuario = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgUsuarioRede", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Retorna nome da estação de trabalho do usuário logado na rede.

Public Function fgEstacaoTrabalhoUsuario() As String

Dim objUsuario                              As A6A7A8.clsUsuario

On Error GoTo ErrorHandler

    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    Call objUsuario.ObterEstacaoTrabalhoUsuario(fgUsuarioRede, fgEstacaoTrabalhoUsuario, lngCodigoErroNegocio, intNumeroSequencialErro)
    Set objUsuario = Nothing

    Exit Function
ErrorHandler:

    Set objUsuario = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgEstacaoTrabalhoUsuario", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Trata inclusão de campos declarados como varchar de 4000 na base de dados Oracle.

Public Function fgInsertVarchar4000(ByVal pstrConteudoCampoVarchar As String) As Long

Dim objTransacao                            As A6A7A8CA.clsTransacao

On Error GoTo ErrorHandler
         
    Set objTransacao = CreateObject("A6A7A8CA.clsTransacao")
    
    fgInsertVarchar4000 = objTransacao.InsertVarchar4000("A6.TB_TEXT_XML", _
                                                         "CO_TEXT_XML", _
                                                         "TX_XML", _
                                                         pstrConteudoCampoVarchar, _
                                                         "NU_SEQU_TEXT_XML", _
                                                         "A6.SQ_A6_CO_TEXT_XML")

    
    Set objTransacao = Nothing
    
    
    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "InsertVarchar4000", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Trata consulta de campos declarados como varchar de 4000 na base de dados Oracle.

Public Function fgSelectVarchar4000(ByVal plngSequencial As Variant, _
                           Optional ByVal pblnConverterBase64 As Boolean = True) As String

Dim objConsulta                             As A6A7A8CA.clsConsulta
Dim strTabela                               As String
On Error GoTo ErrorHandler

    If plngSequencial < 0 Then
        strTabela = "A6HIST.TB_TEXT_XML"
    ElseIf plngSequencial = 0 Then
        strTabela = "A8.TB_TEXT_XML"
    ElseIf plngSequencial > 0 Then
        strTabela = "A6.TB_TEXT_XML"
    End If

    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")

    fgSelectVarchar4000 = objConsulta.SelectVarchar4000(strTabela, _
                                                        "CO_TEXT_XML", _
                                                        Abs(plngSequencial), _
                                                        "TX_XML", _
                                                        "NU_SEQU_TEXT_XML", _
                                                        pblnConverterBase64)

    Set objConsulta = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA7BusServer", "fgSelectVarchar4000", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Retorna o código do item caixa, sem o complemento dos zeros à direita.

Public Function fgObterCodigoItemCaixaSelect(ByVal pstrItemCaixa As String) As String
    
Dim strCodigoSelect                         As String

    If Right$(pstrItemCaixa, 15) = String(15, "0") Then
        strCodigoSelect = Left$(pstrItemCaixa, 1)
    ElseIf Right$(pstrItemCaixa, 12) = String(12, "0") Then
        strCodigoSelect = Left$(pstrItemCaixa, 4)
    ElseIf Right$(pstrItemCaixa, 9) = String(9, "0") Then
        strCodigoSelect = Left$(pstrItemCaixa, 7)
    ElseIf Right$(pstrItemCaixa, 6) = String(6, "0") Then
        strCodigoSelect = Left$(pstrItemCaixa, 10)
    ElseIf Right$(pstrItemCaixa, 3) = String(3, "0") Then
        strCodigoSelect = Left$(pstrItemCaixa, 13)
    Else
        strCodigoSelect = pstrItemCaixa
    End If
    
    fgObterCodigoItemCaixaSelect = strCodigoSelect
    
End Function

' Retorna o código de item de caixa genérico para o tipo de caixa passado como parâmetro.

Public Function fgObterItemCaixaGenerico(ByVal pintTipoCaixa As enumTipoCaixa, _
                                         ByVal pstrCodigoVeiculoLegal As String, _
                                Optional ByVal pstrSiglaSistema As String = vbNullString) As String

Dim objItemCaixa                            As A6SubReserva.clsItemCaixa

    On Error GoTo ErrorHandler

    Set objItemCaixa = CreateObject("A6SubReserva.clsItemCaixa")
    fgObterItemCaixaGenerico = objItemCaixa.fgObterItemCaixaGenerico(pintTipoCaixa, _
                                                                     pstrCodigoVeiculoLegal, _
                                                                     pstrSiglaSistema)
    Set objItemCaixa = Nothing
    
    Exit Function

ErrorHandler:
    Set objItemCaixa = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgObterItemCaixaGenerico", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Obtém grupo de veículo legal a partir do veículo legal passado como parâmetro.

Public Function fgObterGrupoVeiculoLegal(ByVal pstrCodigoVeiculoLegal As String, _
                                Optional ByVal pstrSiglaSistema As String = vbNullString) As Long

Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal

On Error GoTo ErrorHandler

    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
    fgObterGrupoVeiculoLegal = objVeiculoLegal.ObterGrupoVeiculoLegal(pstrCodigoVeiculoLegal, _
                                                                      pstrSiglaSistema)
    Set objVeiculoLegal = Nothing

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgObterGrupoVeiculoLegal", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Obtém tipo de backoffice do usuário.

Public Function fgObterTipoBackOffice(ByVal pstrCodigoVeiculoLegal As String, _
                             Optional ByVal pstrSiglaSistema As String = vbNullString) As Long

Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal

On Error GoTo ErrorHandler

    'Tipo BackOffice
    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
    fgObterTipoBackOffice = objVeiculoLegal.ObterTipoBackOffice(pstrCodigoVeiculoLegal, _
                                                                pstrSiglaSistema)
    Set objVeiculoLegal = Nothing

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgObterTipoBackOffice", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Obtém tipo de backoffice do usuário.

Public Function fgObterTipoBackOfficeUsuario() As Integer

Dim objUsuario                           As A6A7A8.clsControleAcesso

On Error GoTo ErrorHandler
    
    Set objUsuario = CreateObject("A6A7A8.clsControleAcesso")
    
    fgObterTipoBackOfficeUsuario = objUsuario.ObterTipoBackOfficeUsuario(fgUsuarioRede)
    
    Set objUsuario = Nothing
    Exit Function
ErrorHandler:
    
    Set objUsuario = Nothing
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA6SBR", "fgObterTipoBackOfficeUsuario Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Recebe como parâmetro um valor com o separador decimal em qualquer formato.
' Retorna um valor com o separador decimal no formato do Banco de Dados (Ponto)

Public Function fgVlrToDBServer(ByVal pvntNumero As Variant, _
                       Optional ByVal pbDecimalZonado As Boolean = False, _
                       Optional ByVal piQtdeDecimais As Integer = 2) As Variant

Dim strSeparador                As String
Dim strSeparadorTroca           As String
Dim intPosSeparador             As Integer
Dim intPosSeparadorTroca        As Integer

    On Error GoTo ErrorHandler
    
    If Trim(pvntNumero) = vbNullString Then
        fgVlrToDBServer = 0
        Exit Function
    End If
    
    If Int(CDbl("1.1")) = 1 Then
        strSeparador = "."
        strSeparadorTroca = ","
    Else
        strSeparador = ","
        strSeparadorTroca = "."
    End If
    
    intPosSeparador = InStr(pvntNumero, strSeparador)
    intPosSeparadorTroca = InStr(pvntNumero, strSeparadorTroca)
    
    If intPosSeparador > 0 And intPosSeparadorTroca > 0 Then
        pvntNumero = Replace(pvntNumero, _
                IIf(intPosSeparador > intPosSeparadorTroca, strSeparadorTroca, strSeparador), vbNullString)
    End If
     
    fgVlrToDBServer = Replace(pvntNumero, strSeparadorTroca, strSeparador)
    fgVlrToDBServer = Replace(fgVlrToDBServer, ",", ".")
    
    Exit Function

ErrorHandler:
    lngCodigoErroNegocio = 43
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgVlrToDBServer|" & pvntNumero, lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Obtém quantidade de dias de expurgo do XML de base histórica.

Public Function fgObterQtdDiasExpurgo() As Long
On Error GoTo ErrorHandler
Dim strXMLParmGeral                         As String
Dim xmlParmGeral                            As DOMDocument40

    strXMLParmGeral = fgSelectVarchar4000(0, False)
    Set xmlParmGeral = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlParmGeral.loadXML(strXMLParmGeral) Then
        GoTo ErrorHandler
    End If
    
    If xmlParmGeral.selectSingleNode("//BASE_HISTORICA") Is Nothing Then
        'Valor default
        fgObterQtdDiasExpurgo = 40
    Else
        fgObterQtdDiasExpurgo = CLng(xmlParmGeral.selectSingleNode("//BASE_HISTORICA").Text)
    End If
    
    Set xmlParmGeral = Nothing
    
    Exit Function
ErrorHandler:
    Set xmlParmGeral = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgObterQtdDiasExpurgo", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Verifica erros de oracle e MQSeries.

Public Function fgVerificaErroOracleMQSeries(ByVal pstrErro As String) As Boolean

Dim xmlErro                                 As MSXML2.DOMDocument40
    
On Error GoTo ErrorHandler
    
    Set xmlErro = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlErro.loadXML(pstrErro) Then
        
        If xmlErro.selectSingleNode("//ErrorType") Is Nothing Then
            fgVerificaErroOracleMQSeries = False
        Else
            If xmlErro.selectSingleNode("//ErrorType").Text = "2" Then
                fgVerificaErroOracleMQSeries = True
            End If
        End If
    Else
        fgVerificaErroOracleMQSeries = True
    End If
     
    Set xmlErro = Nothing
    
    Exit Function
    
ErrorHandler:
       
    Set xmlErro = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'Gravar arquivo de log de erro
Public Sub fgGravaArquivo(ByVal pstrNomeArquivo As String, _
                          ByVal pstrErro As String)


Dim strNomeArquivoLogErro                    As String
Dim intFile                                  As Integer
Dim strMensagem                              As String

On Error GoTo ErrorHandler

    strMensagem = String$(50, "*") & vbCrLf
    strMensagem = strMensagem & pstrErro & vbCrLf
    strMensagem = strMensagem & String$(50, "*")

    strNomeArquivoLogErro = App.Path & "\log\" & pstrNomeArquivo & "_" & Format(Now, "yyyymmddHHmmss") & ".log"
    intFile = FreeFile
    Open strNomeArquivoLogErro For Output As intFile
    Print #intFile, strMensagem
    Close intFile
  
    Exit Sub

ErrorHandler:
    
    Close intFile
        
    Err.Clear
        
End Sub



