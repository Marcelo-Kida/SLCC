Attribute VB_Name = "basA6A7A8"

' Este componente tem como objetivo agrupar métodos utilizados pelos sistema A8, A7 e A8.

Option Explicit

Public Const strGRUPO_VEIC_LEGA_PADRAO      As String = "Grupo Padrão"

'Variáveis auxiliares para otimização da rotina de segregação de dados
Public gstrUsuarioRede                      As String

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

' Retorna o Tipo de Backoffice do usuário.

Public Function fgObterTipoBackOfficeUsuario() As Integer

Dim objUsuario                           As A6A7A8.clsControleAcesso

On Error GoTo ErrorHandler
    
    Set objUsuario = CreateObject("A6A7A8.clsControleAcesso")
    
    fgObterTipoBackOfficeUsuario = objUsuario.ObterTipoBackOfficeUsuario(fgObterUsuarioRede)
    
    Set objUsuario = Nothing
    Exit Function
ErrorHandler:
    
    Set objUsuario = Nothing
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.Path, "basA6A7A8", "fgObterTipoBackOfficeUsuario Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

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

' Executa um comando SQL na base de dados Oracle.

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
    Call fgRaiseError(App.EXEName, "basA6A7A8", "fgExecuteSQL", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Executa um comando SQL na base de dados Oracle.

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
    Call fgRaiseError(App.EXEName, "basA6A7A8", "fgExecuteCMD", lngCodigoErroNegocio, intNumeroSequencialErro)
    
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
    Call fgRaiseError(App.EXEName, "basA6A7A8", "fgPropriedades", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Executa um Select na base de dados Oracle.

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

' Executa um Select na base de dados Oracle e retorna o resulta em uma string XML.

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

' Executa um Select na base de dados Oracle e retorna o resulta em uma string XML.

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

' Função genérica para tratamento de erro no carregamento de uma string XML.

Public Function fgErroLoadXML(ByRef objDOMDocument As MSXML2.DOMDocument40, _
                              ByVal pstrComponente As String, _
                              ByVal pstrClasse As String, _
                              ByVal pstrMetodo As String)
    

    Err.Raise objDOMDocument.parseError.errorCode, pstrComponente & " - " & pstrClasse & " - " & pstrMetodo, objDOMDocument.parseError.reason
    
End Function

' Executa uma Sequence declarada no banco de dados Oracle.

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
    Call fgRaiseError(App.EXEName, "basA6A7A8", "fgExecuteSequence", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Esta função deve ser utilizada para toda Procedure que acesse alguma tabela do PJ/PK,
' Ex: A8PROC.A8P_ADICIONA_DIAS_UTEIS

'Public Function fgExecuteCMD_NotSuported(ByVal pstrNomeProc As String, _
'                                         ByVal pintPosicaoRetorno As Integer, _
'                                         ByRef pvntParametros() As Variant) As Variant
'
'Dim objTransacao                            As A6A7A8CA.clsDBLConsulta
'
'On Error GoTo ErrHandler
'
'    Set objTransacao = CreateObject("A6A7A8CA.clsDBLConsulta")
'    fgExecuteCMD_NotSuported = objTransacao.ExecuteCMD(pstrNomeProc, pintPosicaoRetorno, pvntParametros())
'    Set objTransacao = Nothing
'
'    Exit Function
'
'ErrHandler:
'
'    Set objTransacao = Nothing
'
'    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
'
'    Call fgRaiseError(App.EXEName, "basA6SubReserva", "fgExecuteCMD_NotSuported", lngCodigoErroNegocio, intNumeroSequencialErro)
'
'End Function

' Obtém usuário logado na rede.

Public Function fgObterUsuarioRede() As String

Dim objUsuario                              As A6A7A8.clsUsuario

On Error GoTo ErrorHandler
    
    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    fgObterUsuarioRede = objUsuario.ObterUsuarioRede
    Set objUsuario = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objUsuario = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A7A8", "fgObterUsuarioRede", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Obtém estação de trabalho do usuário.

Public Function fgObterEstacaoTrabalhoUsuario() As String
 
Dim objUsuario                              As A6A7A8.clsUsuario
Dim strUsuarioRede                          As String
Dim strEstacaoTrabalho                      As String

On Error GoTo ErrorHandler
    
    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    
    strUsuarioRede = fgObterUsuarioRede
    Call objUsuario.ObterEstacaoTrabalhoUsuario(strUsuarioRede, _
                                                strEstacaoTrabalho, _
                                                lngCodigoErroNegocio, _
                                                intNumeroSequencialErro)
    
    If lngCodigoErroNegocio <> 0 Then
        GoTo ErrorHandler
    End If
    
    fgObterEstacaoTrabalhoUsuario = strEstacaoTrabalho

    Set objUsuario = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objUsuario = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A7A8", "fgObterEstacaoTrabalhoUsuario", lngCodigoErroNegocio, intNumeroSequencialErro)
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
    Call fgRaiseError(App.EXEName, "basA6A7A8", "fgVlrToDBServer|" & pvntNumero, lngCodigoErroNegocio, intNumeroSequencialErro)
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
                      Optional ByVal xmlDocComplementoWhere As MSXML2.DOMDocument40, _
                      Optional ByVal blnConsideraNulos As Boolean = False) As String

Dim objControleAcesso                       As A6A7A8.clsControleAcesso

On Error GoTo ErrorHandler

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
                                                    xmlDocComplementoWhere, _
                                                    blnConsideraNulos)

    Set objControleAcesso = Nothing
    
    Exit Function

ErrorHandler:
    
    Set objControleAcesso = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A7A8", "fgSegregaDados Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verifica se as contas de custodia são da memsa camara
Public Function fgVerificaContaMesmaCamara(ByVal pstrContaCustodia1 As String, _
                                           ByVal pstrContaCustodia2 As String) As Boolean
                                           
Dim lngLocalLiquidacaoConta1                As Long
Dim lngLocalLiquidacaoConta2                As Long
                                           
On Error GoTo ErrorHandler
    
    lngLocalLiquidacaoConta1 = fgObterLocalLiquidacaoConta(pstrContaCustodia1)
    
    lngLocalLiquidacaoConta2 = fgObterLocalLiquidacaoConta(pstrContaCustodia2)
    
    If lngLocalLiquidacaoConta1 = 0 And lngLocalLiquidacaoConta2 = 0 Then
        fgVerificaContaMesmaCamara = False
    Else
        fgVerificaContaMesmaCamara = (lngLocalLiquidacaoConta1 = lngLocalLiquidacaoConta2)
    End If

    Exit Function

ErrorHandler:
    
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A7A8", "fgVerificaContaMesmaCamara Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter a Camara da conta custodia
Public Function fgObterLocalLiquidacaoConta(ByVal pstrContaCustodia As String) As Long
                                           
Dim strContaAux                             As String
                                           
On Error GoTo ErrorHandler
    
    strContaAux = fgCompletaString(pstrContaCustodia, "0", 9, True)
    
    strContaAux = Mid$(strContaAux, 5, 4)
    
    Select Case strContaAux
        
        Case "7190", "8190"
            fgObterLocalLiquidacaoConta = enumLocalLiquidacao.CLBCTPub
        
        Case "7290", "8290"
            fgObterLocalLiquidacaoConta = enumLocalLiquidacao.BMA
            
        Case "7390", "8390"
            fgObterLocalLiquidacaoConta = enumLocalLiquidacao.BMD
            
        Case "7490", "8490"
            fgObterLocalLiquidacaoConta = enumLocalLiquidacao.BMC
            
        Case "7590", "8590"
            fgObterLocalLiquidacaoConta = enumLocalLiquidacao.CETIP
            
        Case "7690", "8690"
            fgObterLocalLiquidacaoConta = enumLocalLiquidacao.CIP
            
    End Select

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A7A8", "fgObterLocalLiquidacaoConta Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Gravar arquivo de Log

Public Sub fgGravaArquivo(ByVal pstrErro As String)

Dim strNomeArquivoLogErro                    As String
Dim intFile                                  As Integer
Dim strMensagem                              As String

On Error GoTo ErrorHandler

    strMensagem = String(50, "*") & vbCrLf
    strMensagem = strMensagem & pstrErro & vbCrLf
    strMensagem = strMensagem & String(50, "*")

    strNomeArquivoLogErro = App.Path & "\log\ERRO_MBS" & Format(Now, "yyyymmddHHmmss") & ".log"
    intFile = FreeFile
    Open strNomeArquivoLogErro For Output As intFile
    Print #intFile, strMensagem
    Close intFile
  
    Exit Sub

ErrorHandler:
    
    Close intFile
        
    Err.Clear
        
End Sub

'Obter erro sistema MBS

Public Function fgObterErroMBS(ByVal plngCodErroRetMBS As Long, _
                               ByRef plngCodErroNegocio As Long, _
                               ByRef pstrDescicaoErro As String)
                               

Dim objErro                     As A6A7A8CA.clsLogErro

On Error GoTo ErrorHandler

    Select Case plngCodErroRetMBS
        Case enumCodigoErroMBS.NADA_ENCONTRADO
            plngCodErroNegocio = 45
        Case enumCodigoErroMBS.SEQUENCIA_VAZIA
            plngCodErroNegocio = 46
        Case enumCodigoErroMBS.FALHA_ABERTURA_ARQUIVO
            plngCodErroNegocio = 47
        Case enumCodigoErroMBS.PARAMETRO_NAO_ENCONTRADO
            plngCodErroNegocio = 48
        Case enumCodigoErroMBS.ASSOC_CADASTRADA
            plngCodErroNegocio = 49
        Case enumCodigoErroMBS.PARAMETROS_ENTRADA_INCONSISTENTE
            plngCodErroNegocio = 50
        Case enumCodigoErroMBS.ERRO_CONEXAO
            plngCodErroNegocio = 51
        Case enumCodigoErroMBS.SEQUENCIA_DE_TRANSACOES_NAO_CADASTRADA
            plngCodErroNegocio = 52
        Case enumCodigoErroMBS.TRANSACAO_NAO_CADASTRADA
            plngCodErroNegocio = 53
        Case enumCodigoErroMBS.HIERARQUIA_CADASTRADA
            plngCodErroNegocio = 54
        Case enumCodigoErroMBS.TRANSACAO_JA_E_FILHA
            plngCodErroNegocio = 55
        Case enumCodigoErroMBS.TRANSACAO_NIVEL_ZERO
            plngCodErroNegocio = 56
        Case enumCodigoErroMBS.US_TRANSACAO_AMARRADA_COM_SUBGRUPO_FAVORITOS
            plngCodErroNegocio = 57
        Case enumCodigoErroMBS.HIERARQUIA_NAO_CADASTRADA
            plngCodErroNegocio = 58
        Case enumCodigoErroMBS.SISTEMA_CADASTRADO
            plngCodErroNegocio = 59
        Case enumCodigoErroMBS.TRANSACAO_CADASTRADA
            plngCodErroNegocio = 60
        Case enumCodigoErroMBS.USUARIO_CADASTRADO
            plngCodErroNegocio = 61
        Case enumCodigoErroMBS.GRUPO_DE_USUARIOS_CADASTRADO
            plngCodErroNegocio = 62
        Case enumCodigoErroMBS.USUARIO_NAO_CADASTRADO
            plngCodErroNegocio = 63
        Case Else
            plngCodErroNegocio = 21
    End Select
                                        
    Set objErro = CreateObject("A6A7A8CA.clsLogErro")
    pstrDescicaoErro = objErro.ObterDescErroNegocio(plngCodErroNegocio)
    Set objErro = Nothing
    
    Exit Function
    

ErrorHandler:
    Set objErro = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A7A8", "fgObterErroMBS Function", lngCodigoErroNegocio, intNumeroSequencialErro)
                               
                               
End Function

' Trata consulta de campos declarados como varchar de 4000 na base de dados Oracle.
Public Function fgSelectVarchar4000(ByVal pvntSequencial As Variant, _
                           Optional ByVal pblnConverterBase64 As Boolean = True) As String

'Retorna decode base64

Dim objConsulta                             As A6A7A8CA.clsConsulta
Dim strTabelaTexto                          As String

On Error GoTo ErrorHandler
    
    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    
    '-----------------------------------------------------------------------------------
    'Pikachu - 22/06/2004'
    'OBS : O registro com CO_TEXT_XML = 0 (zero) é destinado para os parametros gerais
    '-----------------------------------------------------------------------------------
    
    If pvntSequencial >= 0 Then
        strTabelaTexto = "A8.TB_TEXT_XML"
    Else
        strTabelaTexto = "A8HIST.TB_TEXT_XML"
    End If
    
    fgSelectVarchar4000 = objConsulta.SelectVarchar4000(strTabelaTexto, _
                                                        "CO_TEXT_XML", _
                                                        Abs(pvntSequencial), _
                                                        "TX_XML", _
                                                        "NU_SEQU_TEXT_XML", _
                                                        pblnConverterBase64)
    
    Set objConsulta = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objConsulta = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgSelectVarchar4000", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Função implementada para melhoria de performance no Recebimento de Operações
Public Function fgObterTipoOperacaoPorLayout_Fixo(ByRef pxmlRemessa As MSXML2.DOMDocument40, _
                                                  ByRef plngTipoOperacao As Long, _
                                                  ByRef pstrCodigoMensagemSPB As String, _
                                                  ByRef pstrMensagemRetornoLegado As String) As Boolean

Dim intLayout                               As Integer

    intLayout = Val(pxmlRemessa.selectSingleNode("//TP_MESG").Text)
    
    Select Case intLayout
        Case 82
            plngTipoOperacao = 63
            pstrCodigoMensagemSPB = "CTP9015"
            pstrMensagemRetornoLegado = "83"
    
    End Select
    
End Function
