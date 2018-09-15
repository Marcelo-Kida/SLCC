Attribute VB_Name = "basA6A8ValidaRemessa"
'Empresa        : Regerbanc - Participações , Negócios e Serviços LTDA
'Pacote         :
'Classe         : basA6A8ValidaRemessa
'Data Criação   : 12/9/2003
'Objetivo       : Validar as mensagens recebidas pelos sistemas A6 e A8
'
'Analista       :
'
'Programador    : Eder Andrade
'Data           : 23/9/2003
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :

Option Explicit

Public glngCodigoEmpresa                    As Long
Public glngCodigoBanco                      As Long
Public glngCodigoEmpresaFusionada           As Long
Public glngCodigoBancoCustodia              As Long
Public gdatDataVigenciaEmpresa              As Date
Public glngCodigoLocalLiquidacao            As Long
Public gdatDataVigenciaLocalLiquidacao      As Date
Public gstrSiglaSistema                     As String
Public gdatDataVigenciaSistema              As Date
Public glngCodigoProduto                    As Long
Public gdatDataVigenciaProduto              As Date
Public gdatDataVigenciaVeiculo              As Date
Public gstrTipoLiquidacao                   As String
Public gstrVeiculoLegal                     As String
Public gstrGrupoVeiculoLegal                As String
Public gdatDataUtil                         As Date
Public gdatFeriado                          As Date
Public gstrDataRemessa                      As String
Public gdatDataCaixaSubReserva              As Date

Public glngTipoMesg                         As Long
Public gdatDataVigenciaTipoMesg             As Date
Public glngTipoConta                        As Long
Public gdatDataVigenciaTipoConta            As Date
Public glngCodigoSegmento                   As Long
Public gdatDataVigenciaSegmento             As Date
Public glngCodigoEventoFinanceiro           As Long
Public gdatDataVigenciaEventoFinanceiro     As Date
Public glngCodigoIndexador                  As Long
Public gdatDataVigenciaIndexador            As Date

Public glngTipoVinculo                      As Long
Public glngTipoRedesconto                   As Long
Public glngTipoTransferencia                As Long
Public glngTipoTransferenciaLDL             As Long
Public glngTipoCompromisso                  As Long
Public glngTipoCompromissoRetn              As Long
Public glngTipoLiquidacaoMens               As Long
Public gvntTipoLeilao                       As Long
Public glngTipoPgtoLDL                      As Long
Public glngOperacaoRotinaAbr                As Long
Public glngTipoPagtoRedesc                  As Long
Public gstrTipoPagto                        As String
Public glngFinalidadeCobertura              As Long
Public glngModalidadeLiquidacao             As Long
Public glngTipoContratoSwap                 As Long
Public glngCodigoMoeda                      As Long
Public glngTipoMovimento                    As Long
Public glngFinalidadeIF                     As Long
Public glngIndicadorBoletim                 As Long
Public gvntTipoTrigger                      As Variant
Public glngTipoFonte                        As Long
Public gvntIndicadorFormaLiquRebate         As Variant
Public gvntIndicadorPeriodicidade           As Variant
Public gvntTipoCliente                      As Variant
Public gstrIndicadorExecrOpacao             As String
Public gvntCodigoIndexadorCetip             As Variant
Public gvntCodigoIndexadorTermoCetip        As Variant
Public gvntTipoIndexadorTermoCetip          As Variant
Public gvntCodigoIndexadorEspecialCetip     As Variant

Public gstrTipoTitular                      As String
Public gstrSubTipoAtivo                     As String
Public glngTipoNegociacaoBMA                As Long
Public glngSubTipoNegociacaoBMA             As Long
Public glngTipoAtivo                        As Long

Public gstrTipoTitularBMA                   As String
Public glngTipoBackOffice                   As Long
Public gdatDataVigenciaBackOffice           As Date

Public rsControleAcesso                     As ADODB.Recordset
Public rsCtrlProcOperAtiv                   As ADODB.Recordset

Public lngCodigoErroNegocio                 As Long
Public intNumeroSequencialErro              As Integer

Private Const strSTR0004                    As String = "STR0004"
Private Const strSTR0007                    As String = "STR0007"

Public Function fgConsisteDatasRemessa(ByRef xmlDOMMensagem As MSXML2.DOMDocument40, ByVal strErros As String) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim xmlParametroGeral                       As MSXML2.DOMDocument40
Dim xmlErrosComplemento                     As MSXML2.DOMDocument40
Dim lngTipoMensagem                         As Long

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlErrosNegocio.loadXML(strErros)

    With xmlDOMMensagem.selectSingleNode("/MESG")
    
        lngTipoMensagem = Val(.selectSingleNode("TP_MESG").Text)
        
        'Data da Mensagem Invalida
        Set objDomNode = .selectSingleNode("DT_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            ElseIf fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            End If
        End If
        
        'Data da operação
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data de movimento inválida
                fgAdicionaErro xmlErrosNegocio, 4012
            Else
                
                Select Case lngTipoMensagem
                    Case enumTipoMensagemLQS.EventoJurosCETIP, _
                         enumTipoMensagemLQS.RetornoCompromissada
                         
                        gstrDataRemessa = vbNullString
                        
                        Set xmlParametroGeral = CreateObject("MSXML2.DOMDocument.4.0")
                        Call xmlParametroGeral.loadXML(fgSelectVarchar4000(0, False))
                    
                        If Not xmlParametroGeral.selectSingleNode("//DIAS_LIMITE_RECEBIMENTO_OPERACAO") Is Nothing Then
                            If xmlParametroGeral.selectSingleNode("//DIAS_LIMITE_RECEBIMENTO_OPERACAO").Text <> vbNullString Then
                                gstrDataRemessa = fgDt_To_Xml(flDataHoraServidor(DataAux) + Val(xmlParametroGeral.selectSingleNode("//DIAS_LIMITE_RECEBIMENTO_OPERACAO").Text))
                            End If
                        End If
            
                        If gstrDataRemessa <> vbNullString Then
                            If Val(objDomNode.Text) > Val(gstrDataRemessa) Then
                                'Data da operação é posterior à informada no Cadastro de Parâmetros Gerais do Módulo A7.
                                fgAdicionaErro xmlErrosNegocio, 4266
                                
                                Set xmlErrosComplemento = CreateObject("MSXML2.DOMDocument.4.0")
                                
                                Call fgAppendNode(xmlErrosComplemento, vbNullString, "Grupo_ErrorInfo", vbNullString)
                                Call fgAppendNode(xmlErrosComplemento, "Grupo_ErrorInfo", "Number", 4266)
                                Call fgAppendNode(xmlErrosComplemento, "Grupo_ErrorInfo", "Description", "Data Operação : " & Val(objDomNode.Text) & " / Data Comparação : " & Val(gstrDataRemessa))
                                
                                Call fgAppendXML(xmlErrosNegocio, "Erro", xmlErrosComplemento.xml)
                            
                                Set xmlErrosComplemento = Nothing
                            End If
                        End If
                        
                End Select
            
            End If
        End If

    End With
    
    fgConsisteDatasRemessa = xmlErrosNegocio.xml

    Set xmlParametroGeral = Nothing
    Set xmlErrosComplemento = Nothing
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    Exit Function

ErrorHandler:
    Set xmlParametroGeral = Nothing
    Set xmlErrosComplemento = Nothing
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteDatasRemessa", 0

End Function

Public Function fgValidaMensagemDadosComplContratoSWAP(ByRef pxmlRemessa As MSXML2.DOMDocument40, _
                                                       ByRef pxmlErro As MSXML2.DOMDocument40, _
                                                       ByVal pstrDataOperacao As String) As Long

Dim strSQL                                  As String
Dim rsQuery                                 As ADODB.Recordset
Dim lngErro                                 As Long
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strNumComando                           As String
Dim strCodigoVeiculoLegal                   As String
Dim vntIdentificadorParticipante            As Variant
Dim vntIdentificadorContraparte             As Variant
'Dim strDHInicio                             As String
'Dim strDHFim                                As String

On Error GoTo ErrorHandler
    
    With pxmlRemessa
        
        Set xmlNode = .selectSingleNode("//NU_COMD_OPER")
        If xmlNode Is Nothing Then
            'Número Operação Participante é obrigatório
            fgAdicionaErro pxmlErro, 4165
        Else
            If Trim$(xmlNode.Text) = vbNullString Then
                'Número Operação Participante obrigatório
                fgAdicionaErro pxmlErro, 4165
            Else
                strNumComando = Trim$(xmlNode.Text)
            End If
        End If
        
        Set xmlNode = .selectSingleNode("//CO_PARP_CAMR")
        If Not xmlNode Is Nothing Then
            If Not IsNumeric(xmlNode.Text) Or Trim$(xmlNode.Text) = vbNullString Then
                 'Identificador Participante Câmara inválido
                 fgAdicionaErro pxmlErro, 4189
            Else
                vntIdentificadorParticipante = Val(xmlNode.Text)
            End If
        End If
        
        Set xmlNode = .selectSingleNode("//CO_CNPT_CAMR")
        If Not xmlNode Is Nothing Then
            If Not IsNumeric(xmlNode.Text) Or Trim$(xmlNode.Text) = vbNullString Then
                 'Identificador de contraparte Câmara inválido
                 fgAdicionaErro pxmlErro, 4139
            Else
                vntIdentificadorContraparte = Val(xmlNode.Text)
            End If
        End If
       
    End With
    
    pstrDataOperacao = fgDtXML_To_Oracle(pstrDataOperacao)
    
    strSQL = "SELECT    A.CO_ULTI_SITU_PROC " & _
             "  FROM    A8.TB_OPER_ATIV A,  " & _
             "          A8.TB_VEIC_LEGA B   " & _
             " Where    A.CO_VEIC_LEGA                      = B.CO_VEIC_LEGA " & _
             "   AND    A.CO_CNPT_CAMR                      =  " & vntIdentificadorContraparte & _
             "   AND    B.ID_PART_CAMR_CETIP                =  " & vntIdentificadorParticipante & _
             "   AND    A.NU_COMD_OPER                      = '" & strNumComando & "'" & _
             "   AND    A.TP_OPER                           =  " & enumTipoOperacaoLQS.RegistroContratoSWAP & _
             "   AND    TRUNC(A.DT_OPER_ATIV)               =  " & pstrDataOperacao
       
    Set rsQuery = QuerySQL(strSQL)
    
    If rsQuery.EOF Then
        'Registro de contrato de Swap original não localizado
        fgAdicionaErro pxmlErro, 4190
    Else
        If rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.Pendencia Then
            'Registro de contrato de Swap não está pendente
            fgAdicionaErro pxmlErro, 4188
        End If
    End If
                    
    Set rsQuery = Nothing

    Exit Function
ErrorHandler:

    Set rsQuery = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgValidaMensagemDadosComplContratoSWAP", 0
    
End Function

Public Function fgExisteEmpresa(plngCodigoEmpresa As Long, pdatDataVigencia As Date) As Boolean

Dim strSQL                                  As String
Dim rsEmpresa                               As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = "select co_empr, co_banc,co_empr_fusi, dt_inic_vige, dt_fim_vige from a8.tb_empresa_ho where co_empr = " & plngCodigoEmpresa
    
    Set rsEmpresa = QuerySQL(strSQL)

    If rsEmpresa.RecordCount > 0 Then
        'Verifica a data de vigência
        If rsEmpresa!dt_inic_vige <= pdatDataVigencia And (IsNull(rsEmpresa!DT_FIM_VIGE) Or pdatDataVigencia <= rsEmpresa!DT_FIM_VIGE) Then
            glngCodigoEmpresa = plngCodigoEmpresa
            glngCodigoEmpresaFusionada = rsEmpresa!CO_EMPR_FUSI
            gdatDataVigenciaEmpresa = pdatDataVigencia
            fgExisteEmpresa = True
        End If
    End If
    
    Set rsEmpresa = Nothing
    
    Exit Function
ErrorHandler:
    Set rsEmpresa = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteEmpresa", 0
End Function

Public Function fgExisteTipoContaCorrente(ByVal pStr As String) As Boolean

On Error GoTo ErrorHandler
    
    If Trim(pStr) <> "CC" And _
       Trim(pStr) <> "PP" Then
       fgExisteTipoContaCorrente = False
    Else
       fgExisteTipoContaCorrente = True
    End If

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoContaCorrente", 0
End Function
Public Function fgExisteTipoPessoa(ByVal pStr As String) As Boolean

On Error GoTo ErrorHandler
    
    If Trim(pStr) <> "F" And _
       Trim(pStr) <> "J" Then
       fgExisteTipoPessoa = False
    Else
       fgExisteTipoPessoa = True
    End If

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoPessoa", 0
End Function

Public Function fgExisteCodigoBanco(plngCodigoBanco As Long) As Boolean

Dim strSQL                                  As String
Dim rsBanco                                 As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = "select co_banc from a8.tb_empresa_ho where co_banc = " & plngCodigoBanco

    Set rsBanco = QuerySQL(strSQL)

    If rsBanco.RecordCount > 0 Then
        glngCodigoBanco = IIf(IsNull(rsBanco!CO_BANC), 0, rsBanco!CO_BANC)
        fgExisteCodigoBanco = True
    End If

    Set rsBanco = Nothing
    Exit Function

ErrorHandler:

    Set rsBanco = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteCodigoBanco", 0

End Function
Public Function fgExisteCodigoBancoCustodia(plngCodigoBanco As Long) As Boolean

Dim strSQL                                  As String
Dim rsBancoCustodia                         As ADODB.Recordset

On Error GoTo ErrorHandler

              
    strSQL = "select co_cpen from a8.tb_instituicao_spb " & _
             " where  dt_inic_vige <= " & fgDataHoraServidor_To_Oracle & _
             " and    (dt_fim_vige >= " & fgDataHoraServidor_To_Oracle & " OR dt_fim_vige IS NULL)" & _
             " and    co_cpen = " & plngCodigoBanco
        
    Set rsBancoCustodia = QuerySQL(strSQL)

    If rsBancoCustodia.RecordCount > 0 Then
        glngCodigoBancoCustodia = IIf(IsNull(rsBancoCustodia!CO_CPEN), 0, rsBancoCustodia!CO_CPEN)
        fgExisteCodigoBancoCustodia = True
    End If
    
    Set rsBancoCustodia = Nothing
    
    Exit Function
ErrorHandler:
    Set rsBancoCustodia = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteCodigoBancoCustodia", 0

End Function



Public Function fgExisteBackOffice(ByVal plngTipoBackOffice As Long, _
                                   ByVal pdatDataVigencia As Date) As Boolean

Dim strSQL                                  As String
Dim rsTipoBackOffice                        As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = "select tp_bkof, dt_inic_vige, dt_fim_vige from a8.tb_tipo_bkof where tp_bkof = " & plngTipoBackOffice
    
    Set rsTipoBackOffice = QuerySQL(strSQL)

    If rsTipoBackOffice.RecordCount > 0 Then
        'Verifica a data de vigência
        If rsTipoBackOffice!dt_inic_vige <= pdatDataVigencia And (IsNull(rsTipoBackOffice!DT_FIM_VIGE) Or pdatDataVigencia <= rsTipoBackOffice!DT_FIM_VIGE) Then
            glngTipoBackOffice = plngTipoBackOffice
            gdatDataVigenciaBackOffice = pdatDataVigencia
            fgExisteBackOffice = True
        End If
    End If
    
    Set rsTipoBackOffice = Nothing
    
    Exit Function
ErrorHandler:
    
    Set rsTipoBackOffice = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteBackOffice", 0
End Function

Public Function fgErroLoadXML(ByRef objDomDocument As MSXML2.DOMDocument40, _
                              ByVal pstrComponente As String, _
                              ByVal pstrClasse As String, _
                              ByVal pstrMetodo As String)
    

    Err.Raise objDomDocument.parseError.errorCode, pstrComponente & " - " & pstrClasse & " - " & pstrMetodo, objDomDocument.parseError.reason
    
End Function

Public Function fgExisteLocalLiquidacao(ByVal plngCodigoEmpresaFusionada As Long, _
                                        ByVal plngCodigoLocalLiquidacao As Long, _
                                        ByVal pdatDataVigencia As Date) As Boolean

Dim strSQL                                  As String
Dim rsLocalLiquidacao                       As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = "select  co_loca_liqu, dt_inic_vige, dt_fim_vige "
    strSQL = strSQL & " from a8.tb_local_liquidacao"
    strSQL = strSQL & " Where co_empr_fusi =  " & plngCodigoEmpresaFusionada
    strSQL = strSQL & " And   co_loca_liqu =  " & plngCodigoLocalLiquidacao
    
    Set rsLocalLiquidacao = QuerySQL(strSQL)

    If rsLocalLiquidacao.RecordCount > 0 Then
        'Verifica a data de vigência
        If rsLocalLiquidacao!dt_inic_vige <= pdatDataVigencia And _
            (IsNull(rsLocalLiquidacao!DT_FIM_VIGE) Or _
            pdatDataVigencia <= rsLocalLiquidacao!DT_FIM_VIGE) Then
            glngCodigoLocalLiquidacao = plngCodigoLocalLiquidacao
            gdatDataVigenciaLocalLiquidacao = pdatDataVigencia
            fgExisteLocalLiquidacao = True
        End If
    End If
    
    Set rsLocalLiquidacao = Nothing
    
    Exit Function

ErrorHandler:
    
    Set rsLocalLiquidacao = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgExisteLocalLiquidacao", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function fgExisteSistema(ByVal pstrSiglaSistema As String, _
                                ByVal plngCodigoEmpresa As String, _
                                ByVal pdatDataVigencia As Date) As Boolean

Dim rsSistema                               As ADODB.Recordset
Dim lsSQL                                   As String

On Error GoTo ErrorHandler

    lsSQL = "select * from a7.tb_sist where sg_sist = '" & pstrSiglaSistema & "' and co_empr = " & plngCodigoEmpresa
    
    Set rsSistema = QuerySQL(lsSQL)

    If rsSistema.RecordCount > 0 Then
        'Verifica a data de vigência
        If rsSistema!DT_INIC_VIGE_SIST <= pdatDataVigencia And _
           (IsNull(rsSistema!DT_FIM_VIGE_SIST) Or _
           pdatDataVigencia <= rsSistema!DT_FIM_VIGE_SIST) Then
           
            gstrSiglaSistema = pstrSiglaSistema
            gdatDataVigenciaSistema = pdatDataVigencia
            fgExisteSistema = True
        End If
    End If
    
    Set rsSistema = Nothing
    
    Exit Function
ErrorHandler:
    
    Set rsSistema = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgExisteSistema", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function fgExisteProduto(ByVal plngCodigoEmpresa As Long, _
                                ByVal plngCodigoProduto As Long, _
                       Optional ByVal pdatDataVigencia As Date) As Boolean

Dim strSQL                                  As String
Dim strComplemento                          As String
Dim rsProduto                               As ADODB.Recordset

On Error GoTo ErrorHandler
   
   strComplemento = "Produto    :" & plngCodigoProduto & vbCrLf & _
                    "Empresa    :" & plngCodigoEmpresa & vbCrLf & _
                    "Data Atual :" & pdatDataVigencia & vbCrLf
                    
    strSQL = " SELECT " & _
             "      A.CO_EMPR_FUSI,     " & _
             "      A.CO_EMPR,          " & _
             "      B.CO_PROD,          " & _
             "      B.DT_FIM_VIGE,      " & _
             "      B.DT_INIC_VIGE      " & _
             " FROM  A8.TB_PRODUTO B,   " & _
             "       A8.TB_EMPRESA_HO A " & _
             " Where A.CO_EMPR_FUSI = B.CO_EMPR_FUSI" & _
             " and   co_empr = " & plngCodigoEmpresa & " AND co_prod = " & plngCodigoProduto
             
    Set rsProduto = QuerySQL(strSQL)

    If rsProduto.RecordCount > 0 Then
        If pdatDataVigencia = vbEmpty Then
            fgExisteProduto = True
            glngCodigoProduto = rsProduto!CO_PROD
            gdatDataVigenciaProduto = pdatDataVigencia
        Else
            
            strComplemento = strComplemento & "Data Inic Vige:" & rsProduto!dt_inic_vige & vbCrLf
            strComplemento = strComplemento & "Data Fim  Vige:" & IIf(IsNull(rsProduto!DT_FIM_VIGE), "NULL", rsProduto!DT_FIM_VIGE) & vbCrLf
            
            'Verifica a data de vigência
            If pdatDataVigencia >= rsProduto!dt_inic_vige Then
                If IsNull(rsProduto!DT_FIM_VIGE) Or pdatDataVigencia <= rsProduto!DT_FIM_VIGE Then
                    glngCodigoProduto = rsProduto!CO_PROD
                    gdatDataVigenciaProduto = pdatDataVigencia
                    fgExisteProduto = True
                End If
            End If
        End If
    End If
    
    Set rsProduto = Nothing
    
    Exit Function
ErrorHandler:

    Set rsProduto = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgExisteProduto", lngCodigoErroNegocio, intNumeroSequencialErro, strComplemento)

End Function

Public Function fgExisteVeiculoLegal(ByVal pstrSiglaSistema, _
                                     ByVal pstrVeiculoLegal As String, _
                                     ByVal pdatVigencia, _
                                     ByVal plngCodigoEmpresa As Long) As Boolean

Dim strSQL                                  As String
Dim rsVeiculoLegal                          As ADODB.Recordset
On Error GoTo ErrorHandler

    strSQL = " SELECT CO_CNPJ_VEIC_LEGA, TP_BKOF, CO_VEIC_LEGA, SG_SIST, TP_TITL_BMA,   " & vbNewLine & _
             "        DT_INIC_VIGE, DT_FIM_VIGE, CO_GRUP_VEIC_LEGA, CO_EMPR             " & vbNewLine & _
             " FROM   A8.TB_VEIC_LEGA                                                   " & vbNewLine & _
             " WHERE  CO_VEIC_LEGA = '" & Trim(pstrVeiculoLegal) & "'                   " & vbNewLine & _
             " AND    SG_SIST      = '" & Trim(pstrSiglaSistema) & "'                   " & vbNewLine & _
             " AND    CO_EMPR      =  " & plngCodigoEmpresa
             
    Set rsVeiculoLegal = QuerySQL(strSQL)

    With rsVeiculoLegal
        If .RecordCount > 0 Then
            If !dt_inic_vige <= pdatVigencia And (IsNull(!DT_FIM_VIGE) Or !DT_FIM_VIGE >= pdatVigencia) Then
                
                gstrVeiculoLegal = pstrVeiculoLegal
                gdatDataVigenciaVeiculo = pdatVigencia
                glngTipoBackOffice = !TP_BKOF
                gstrGrupoVeiculoLegal = !CO_GRUP_VEIC_LEGA
                glngCodigoEmpresa = !CO_EMPR

                If IsNull(!TP_TITL_BMA) Then
                    gstrTipoTitularBMA = vbNullString
                Else
                    gstrTipoTitularBMA = !TP_TITL_BMA
                End If
                
                fgExisteVeiculoLegal = True

            End If
        Else
            fgExisteVeiculoLegal = False
        End If
    End With

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteVeiculoLegal", 0
End Function

Public Function fgExisteParticipanteNegociacao(ByVal pstrCodigoPartNegociacao As String) As Boolean

Dim strSQL                                  As String
Dim rsPartNegociacao                        As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " SELECT TP_BKOF, CO_PARP_CAMR " & vbNewLine & _
             " FROM   A8.TB_TIPO_BKOF_PARP_CAMR " & _
             " WHERE  CO_PARP_CAMR = '" & pstrCodigoPartNegociacao & "'"
             
    Set rsPartNegociacao = QuerySQL(strSQL)

    With rsPartNegociacao
        If .RecordCount > 0 Then
            fgExisteParticipanteNegociacao = True
        Else
            fgExisteParticipanteNegociacao = False
        End If

    End With
    
    Set rsPartNegociacao = Nothing
    
    Exit Function

ErrorHandler:
    Set rsPartNegociacao = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteParticipanteNegociacao", 0
End Function

Public Function fgDiaUtil(ByVal pdatData As Date) As Boolean

Dim strSQL                                  As String
Dim rsDataFeriado                           As ADODB.Recordset

On Error GoTo ErrorHandler

    If pdatData = gdatDataUtil Then
        fgDiaUtil = True
        Exit Function
    End If

    If Weekday(pdatData) = vbSunday Or Weekday(pdatData) = vbSaturday Then
        fgDiaUtil = False
        gdatFeriado = pdatData
        Exit Function
    End If

    If rsDataFeriado Is Nothing Then
        strSQL = " SELECT * FROM A8.TB_FERIADO_HO "
        Set rsDataFeriado = QuerySQL(strSQL)
    End If

    rsDataFeriado.Filter = " DT_FERI = " & pdatData
    
    If rsDataFeriado.RecordCount < 1 Then
        gdatDataUtil = pdatData
        fgDiaUtil = True
    Else
        gdatFeriado = pdatData
        fgDiaUtil = False
    End If
    
    Set rsDataFeriado = Nothing
    
Exit Function
ErrorHandler:
    Set rsDataFeriado = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgDiaUtil", 0
End Function

Private Function fgExisteTipoLiquidacao(ByVal pstrTipoLiquidacao As String, ByVal pdatVigencia As Date) As Boolean

Dim strSQL                                  As String
Dim rsTipoLiquidacao                        As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT * FROM A8.TB_TIPO_LIQU_OPER_ATIV where  TP_LIQU_OPER_ATIV = " & pstrTipoLiquidacao
    
    Set rsTipoLiquidacao = QuerySQL(strSQL)
    
    With rsTipoLiquidacao
        If .RecordCount > 0 Then
            If !dt_inic_vige <= pdatVigencia And (IsNull(!DT_FIM_VIGE) Or !DT_FIM_VIGE >= pdatVigencia) Then
                gstrTipoLiquidacao = !NO_TIPO_LIQU_OPER_ATIV
                fgExisteTipoLiquidacao = True
            End If
        End If
    
    End With
    
    Set rsTipoLiquidacao = Nothing
    
    Exit Function

ErrorHandler:
    Set rsTipoLiquidacao = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoLiquidacao", 0
End Function

'Se pintQtdeRegistros = 0, retorna todos os registros selecionados
Public Function QuerySQL(ByVal pstrSQL As String, _
                Optional ByVal pintQtdeRegistros As Integer) As ADODB.Recordset

Dim objConsulta                             As A6A7A8CA.clsConsulta

On Error GoTo ErrHandler

    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    Set QuerySQL = objConsulta.QuerySQL(pstrSQL, pintQtdeRegistros)
    Set objConsulta = Nothing

Exit Function
ErrHandler:
    Set objConsulta = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "QuerySQL", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function ExecuteCMD(ByVal pstrNomeProc As String, _
                           ByVal pstrPosicaoRetorno As Integer, _
                           ByRef pvntParametros() As Variant) As Variant
                           
Dim objTransacao                            As A6A7A8CA.clsTransacao
                           
On Error GoTo ErrHandler

    Set objTransacao = CreateObject("A6A7A8CA.clsTransacao")
    ExecuteCMD = objTransacao.ExecuteCMD(pstrNomeProc, pstrPosicaoRetorno, pvntParametros())
    Set objTransacao = Nothing

Exit Function
ErrHandler:
    Set objTransacao = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "ExecuteCMD", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function
'**************************************************************************
'Encapsular chamada para Classe de Log de Erros
Public Sub fgRaiseError(ByVal psComponente As String, _
                        ByVal psClasse As String, _
                        ByVal psMetodo As String, _
                        ByRef plCodigoErroNegocio As Long, _
               Optional ByRef piNumeroSequencialErro As Integer = 0, _
               Optional ByVal psComplemento As String = "", _
               Optional ByRef pbGravarErro As Boolean = False)
Dim objLogErro                              As A6A7A8CA.clsLogErro
Dim ErrNumber                               As Long
Dim ErrSource                               As String
Dim ErrDescription                          As String
    
    Set objLogErro = CreateObject("A6A7A8CA.clsLogErro")

    objLogErro.RaiseError psComponente, _
                          psClasse, _
                          psMetodo, _
                          plCodigoErroNegocio, _
                          ErrNumber, _
                          ErrSource, _
                          ErrDescription, _
                          piNumeroSequencialErro, _
                          psComplemento, _
                          Err
                          
    Set objLogErro = Nothing
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

Public Function fgAdicionarDiasUteis(ByVal ptData As Date, _
                                     ByVal piQtdeDias As Integer, _
                                     ByVal plMovimento As enumPaginacao) As Date

Dim ltRetorno                               As Date
Dim lvArray()                               As Variant

On Error GoTo ErrHandler

    lvArray = Array(ptData, piQtdeDias, plMovimento)
    ltRetorno = ExecuteCMD("A8PROC.A8P_ADICIONA_DIAS_UTEIS", 3, lvArray)

    fgAdicionarDiasUteis = ltRetorno

Exit Function
ErrHandler:

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgAdicionarDiasUteis", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


Public Function fgConsisteRemessaEstatística(ByRef xmlDOMRemessa As MSXML2.DOMDocument40) As String

Dim datDataVigencia                         As Date
Dim datLimiteFech                           As Date
Dim lngCodigoEmpresaFusi                    As Long
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    With xmlDOMRemessa.selectSingleNode("/MESG")


        datLimiteFech = fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 1, enumPaginacao.Anterior)

        If Not fgDiaUtil(fgDtXML_To_Date(.selectSingleNode("DT_FECH_PROC").Text)) Then
            '4009 - Data de fechamento inválida
            fgAdicionaErro xmlErrosNegocio, 4009
        End If
        
        If fgDtXML_To_Date(.selectSingleNode("DT_FECH_PROC").Text) < datLimiteFech Or _
           fgDtXML_To_Date(.selectSingleNode("DT_FECH_PROC").Text) > flDataHoraServidor(enumFormatoDataHora.Data) Then
            '4009 - Data de fechamento inválida
            fgAdicionaErro xmlErrosNegocio, 4009
        End If
        
        datDataVigencia = fgAdicionarDiasUteis(fgDtXML_To_Date(.selectSingleNode("DT_FECH_PROC").Text), 1, enumPaginacao.proximo)
        
        If Not fgExisteEmpresa(CLng("0" & .selectSingleNode("CO_EMPR").Text), datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        End If
        
        If Not fgExisteSistema(.selectSingleNode("SG_SIST_ORIG").Text, glngCodigoEmpresa, datDataVigencia) Then
            'Sistema inexistente
            fgAdicionaErro xmlErrosNegocio, 4004
        End If
        
        If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                    .selectSingleNode("CO_VEIC_LEGA").Text, _
                                    datDataVigencia, _
                                    .selectSingleNode("CO_EMPR").Text) Then
            'Veículo legal inexistente
            fgAdicionaErro xmlErrosNegocio, 4007
        End If

        If Not .selectSingleNode("SG_SIST_DEST").Text = "A6" Then
            'Sistema de destino incorreto
            fgAdicionaErro xmlErrosNegocio, 4005
        End If
                
    End With

    fgConsisteRemessaEstatística = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteRemessaEstatística", 0

End Function

Public Function fgConsisteRemessaOperacao(ByRef xmlDOMRemessa As MSXML2.DOMDocument40) As String

Dim lngCodigoEmpresaFusi                    As Long
Dim datDataVigencia                         As Date
Dim datLimiteFech                           As Date
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    With xmlDOMRemessa.selectSingleNode("/MESG")

        datLimiteFech = fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 1, enumPaginacao.Anterior)

        If Not fgDiaUtil(fgDtXML_To_Date(.selectSingleNode("DT_FECH_PROC").Text)) Then
            'Data de fechamento inválida
            fgAdicionaErro xmlErrosNegocio, 4009
        End If
        
        If fgDtXML_To_Date(.selectSingleNode("DT_FECH_PROC").Text) < datLimiteFech Or _
           fgDtXML_To_Date(.selectSingleNode("DT_FECH_PROC").Text) > flDataHoraServidor(enumFormatoDataHora.Data) Then
            'Data de fechamento inválida
            fgAdicionaErro xmlErrosNegocio, 4009
        End If
        
        If fgDtXML_To_Date(.selectSingleNode("DT_FECH_PROC").Text) > fgDtXML_To_Date(.selectSingleNode("DT_LIQU_OPER_ATIV").Text) Then
            'Data de fechamento inválida
            fgAdicionaErro xmlErrosNegocio, 4009
        End If
        
        datDataVigencia = fgAdicionarDiasUteis(fgDtXML_To_Date(.selectSingleNode("DT_FECH_PROC").Text), 1, enumPaginacao.proximo)
        
        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        End If
        
        If Not fgExisteSistema(.selectSingleNode("SG_SIST_ORIG").Text, .selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Sistema inexistente
            fgAdicionaErro xmlErrosNegocio, 4004
        End If
        
        If Not .selectSingleNode("SG_SIST_DEST").Text = "A6" Then
            'Sistema de destino incorreto
            fgAdicionaErro xmlErrosNegocio, 4005
        End If
        
        If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                    .selectSingleNode("CO_VEIC_LEGA").Text, _
                                    datDataVigencia, _
                                    .selectSingleNode("CO_EMPR").Text) Then
            'Veículo Legal inexistente
            fgAdicionaErro xmlErrosNegocio, 4007
        End If
        
        If Not fgExisteLocalLiquidacao(glngCodigoEmpresaFusionada, .selectSingleNode("CO_LOCA_LIQU").Text, datDataVigencia) Then
            'Local de Liquidação inválido
            fgAdicionaErro xmlErrosNegocio, 4008
        End If
        
        If Not fgExisteTipoLiquidacao(.selectSingleNode("CO_TIPO_LIQU").Text, datDataVigencia) Then
            'Tipo de Liquidação inválido
            fgAdicionaErro xmlErrosNegocio, 4010
        Else
            .selectSingleNode("CO_TIPO_LIQU").Text = gstrTipoLiquidacao
        End If
        
        If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, .selectSingleNode("CO_PROD").Text, datDataVigencia) Then
            'Produto inválido
            fgAdicionaErro xmlErrosNegocio, 4011
        End If
        
        If Not fgDiaUtil(fgDtXML_To_Date(.selectSingleNode("DT_LIQU_OPER_ATIV").Text)) Then
            .selectSingleNode("DT_LIQU_OPER_ATIV").Text = fgDt_To_Xml(fgAdicionarDiasUteis(fgDtXML_To_Date(.selectSingleNode("DT_LIQU_OPER_ATIV").Text), 1, enumPaginacao.proximo))
        End If
        
        If .selectSingleNode("IN_MOVI_ENTR_SAID").Text <> "1" And _
           .selectSingleNode("IN_MOVI_ENTR_SAID").Text <> "2" Then
            'Indicador de Entrada e Saída inválido
            fgAdicionaErro xmlErrosNegocio, 4013
        End If
        
    End With

    fgConsisteRemessaOperacao = xmlErrosNegocio.xml
    
    Set xmlErrosNegocio = Nothing

    Exit Function
ErrorHandler:

    Set xmlErrosNegocio = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteRemessaOperacao", 0

End Function

Public Function fgConsisteRemessaSaldoFechamento(ByRef xmlDOMRemessa As MSXML2.DOMDocument40) As String

Dim lngCodigoEmpresaFusi                    As Long
Dim datDataVigencia                         As Date
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    With xmlDOMRemessa.selectSingleNode("/MESG")

        If Not fgDiaUtil(fgDtXML_To_Date(.selectSingleNode("DT_APUR_SALD").Text)) Then
            'Data de apuração do saldo inválida
            fgAdicionaErro xmlErrosNegocio, 4014
        End If

        datDataVigencia = fgAdicionarDiasUteis(fgDtXML_To_Date(.selectSingleNode("DT_APUR_SALD").Text), 1, enumPaginacao.proximo)

        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inexistente
            fgAdicionaErro xmlErrosNegocio, 4006
        End If

        If Not fgExisteSistema(.selectSingleNode("SG_SIST_ORIG").Text, glngCodigoEmpresa, datDataVigencia) Then
            'Sistema inexistente
            fgAdicionaErro xmlErrosNegocio, 4004
        End If
        
        If Not .selectSingleNode("SG_SIST_DEST").Text = "A6" Then
            'Sistema de destino incorreto
            fgAdicionaErro xmlErrosNegocio, 4005
        End If
        
        If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                    .selectSingleNode("CO_VEIC_LEGA").Text, _
                                    datDataVigencia, _
                                    .selectSingleNode("CO_EMPR").Text) Then
            'Veículo Legal inexistente
            fgAdicionaErro xmlErrosNegocio, 4007
        End If
        
        If .selectSingleNode("TP_SINA_SALD_FECH").Text <> "1" And _
           .selectSingleNode("TP_SINA_SALD_FECH").Text <> "2" Then
            'Sinal de saldo inválido
            fgAdicionaErro xmlErrosNegocio, 4015
        End If

    End With

    fgConsisteRemessaSaldoFechamento = xmlErrosNegocio.xml
    
    Set xmlErrosNegocio = Nothing

    Exit Function
ErrorHandler:

    Set xmlErrosNegocio = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteRemessaSaldoFechamento", 0

End Function

Public Function fgConsisteRemessaViaSistemaA8(ByRef xmlDOMRemessa As MSXML2.DOMDocument40) As String

Dim datDataVigencia                         As Date
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

'    fgGravaArquivo "LOGA6_ObtemDataServidor", vbNullString
'    Alterada a obtenção da Data de Vigência para tentativa de melhora de performance dos componentes.
'    RATS 779
'    Cas - 18/07/2008
'    =================================================================================================
'    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)
    datDataVigencia = Date

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

'    fgGravaArquivo "LOGA6_basA6A8ValidaRemessa_fgConsisteRemessaViaSistemaA8", vbNullString
    
    With xmlDOMRemessa.selectSingleNode("/MESG")

        If Not .selectSingleNode("DT_OPER_ATIV") Is Nothing Then
            If Not fgDiaUtil(fgDtXML_To_Date(.selectSingleNode("DT_OPER_ATIV").Text)) Then
                'Data de movimento inválida
                fgAdicionaErro xmlErrosNegocio, 4012
            End If
        Else
            If Not fgDiaUtil(fgDtXML_To_Date(.selectSingleNode("DT_MESG").Text)) Then
                'Data de movimento inválida
                fgAdicionaErro xmlErrosNegocio, 4012
            End If
        End If
        
        If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                    .selectSingleNode("//CO_VEIC_LEGA").Text, _
                                    datDataVigencia, _
                                    .selectSingleNode("CO_EMPR").Text) Then
            'Veículo Legal inexistente
            fgAdicionaErro xmlErrosNegocio, 4007
        End If
        
        If .selectSingleNode("IN_OPER_DEBT_CRED") Is Nothing And .selectSingleNode("IN_MOVI_ENTR_SAID") Is Nothing Then
            'Indicador de débito/crédito inválido
            fgAdicionaErro xmlErrosNegocio, 4021
        End If

    End With

    fgConsisteRemessaViaSistemaA8 = xmlErrosNegocio.xml
    
    Set xmlErrosNegocio = Nothing

    Exit Function
ErrorHandler:
    Set xmlErrosNegocio = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteRemessaViaSistemaA8", 0

End Function

Public Sub fgAdicionaErro(ByRef xmlDOMErro As MSXML2.DOMDocument40, _
                          ByVal lngCodigoErroNegocio As Long)


Dim objLogErros                             As A6A7A8CA.clsLogErro

On Error GoTo ErrorHandler

    Set objLogErros = CreateObject("A6A7A8CA.clsLogErro")
    objLogErros.AdicionaErroNegocio xmlDOMErro, lngCodigoErroNegocio
    Set objLogErros = Nothing

    Exit Sub
ErrorHandler:
    Set objLogErros = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgAdicionaErro", 0
End Sub

'A8-----------------------------------------

Public Function fgQueryDominioInternoCCR(ByVal pstrNO_ATRB, ByVal pstrCO_DOMI As String) As Boolean


Dim rsCCR                                   As ADODB.Recordset
Dim strSQL                                  As String

On Error GoTo ErrorHandler

     strSQL = " SELECT CO_DOMI FROM A8.TB_CTRL_DOMI " & vbCrLf & _
              " WHERE NO_ATRB = '" & pstrNO_ATRB & "' " & vbCrLf & _
              " AND   CO_DOMI = '" & pstrCO_DOMI & "' "
              
    Set rsCCR = QuerySQL(strSQL)
    
    If rsCCR.EOF Then
        fgQueryDominioInternoCCR = False
    Else
        fgQueryDominioInternoCCR = True
    End If
    
    Set rsCCR = Nothing
    
    Exit Function

ErrorHandler:
    Set rsCCR = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "flExisteTipoTransferenciaLDL", 0
    
End Function
Public Function fgQueryDominioInterno(ByVal pstrNomeTag As String) As ADODB.Recordset

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT CO_DOMI FROM A8.TB_CTRL_DOMI " & vbCrLf & _
             " WHERE NO_ATRB = '" & pstrNomeTag & "' "
             
    Set fgQueryDominioInterno = QuerySQL(strSQL)

Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgQueryDominioInterno", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function


Public Function fgQueryDominio(ByVal pstrNomeTag As String) As ADODB.Recordset

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT A.CO_DOMI                         " & vbCrLf & _
             " FROM   A8.TB_DOMINIO  A,                    " & vbCrLf & _
             "        A8.TB_TIPO_TAG B                     " & vbCrLf & _
             " WHERE  A.SQ_TIPO_TAG = B.SQ_TIPO_TAG     " & vbCrLf & _
             " AND    B.NO_TIPO_TAG = ( SELECT CO_DOMI FROM A8.TB_CTRL_DOMI " & vbCrLf & _
             "                            WHERE NO_ATRB = '" & pstrNomeTag & "' )"
             
            
    Set fgQueryDominio = QuerySQL(strSQL)

Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgQueryDominio", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

Public Function fgExisteDominioTagMensagemBACEN(ByVal pstrNomeTag As String, ByVal pstrConteudo As String) As Boolean

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

    On Error GoTo ErrorHandler

    strSQL = " SELECT A.CO_DOMI " & vbCrLf & _
             " FROM   A8.TB_DOMINIO  A, " & vbCrLf & _
             "        A8.TB_TIPO_TAG B " & vbCrLf & _
             " WHERE  A.SQ_TIPO_TAG = B.SQ_TIPO_TAG " & vbCrLf & _
             " AND    B.NO_TIPO_TAG = '" & pstrNomeTag & "'" & vbCrLf & _
             " AND    A.CO_DOMI     = '" & pstrConteudo & "'"
            
    Set objRS = QuerySQL(strSQL)
    
    If Not objRS.EOF Then
        fgExisteDominioTagMensagemBACEN = objRS.RecordCount > 0
    Else
        fgExisteDominioTagMensagemBACEN = False
    End If

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgExisteDominioTagMensagemBACEN", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

Public Function fgExisteTipoTransferencia(ByVal plngTipoTransferencia As Long) As Boolean

Dim rsTipoTransferencia                     As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsTipoTransferencia = fgQueryDominio("TP_TRAF")
    
    With rsTipoTransferencia
        .Filter = " CO_DOMI = " & Format$(plngTipoTransferencia, "00")
        If .RecordCount > 0 Then
            glngTipoTransferencia = plngTipoTransferencia
            fgExisteTipoTransferencia = True
        End If
    End With
    
    Set rsTipoTransferencia = Nothing
    
    Exit Function
ErrorHandler:
    
    Set rsTipoTransferencia = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoTransferencia", 0

End Function

Public Function fgExisteTipoTransferenciaLDL(ByVal plngTipoTransferenciaLDL As Long) As Boolean

Dim rsTipoTransferenciaLDL                  As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsTipoTransferenciaLDL = fgQueryDominio("TP_TRAF_LDL")
    
    With rsTipoTransferenciaLDL
    
        .Filter = " CO_DOMI = " & Format$(plngTipoTransferenciaLDL, "00")
    
        If .RecordCount > 0 Then
            glngTipoTransferenciaLDL = plngTipoTransferenciaLDL
            fgExisteTipoTransferenciaLDL = True
        End If
    End With
    
    Set rsTipoTransferenciaLDL = Nothing
    
    Exit Function

ErrorHandler:
    Set rsTipoTransferenciaLDL = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "flExisteTipoTransferenciaLDL", 0

End Function

Public Function fgExisteTipoCompromisso(ByVal plngTipoCompromisso As Long) As Boolean

Dim rsTipoCompromisso                       As ADODB.Recordset

On Error GoTo ErrorHandler
    
    Set rsTipoCompromisso = fgQueryDominio("TP_CPRO_OPER_ATIV")
    
    With rsTipoCompromisso
    
        .Filter = " CO_DOMI = " & Format$(plngTipoCompromisso, "00")
    
        If .RecordCount > 0 Then
            glngTipoCompromisso = plngTipoCompromisso
            fgExisteTipoCompromisso = True
        End If
    End With
    
    Set rsTipoCompromisso = Nothing
    
Exit Function
ErrorHandler:
    
    Set rsTipoCompromisso = Nothing
    
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "flExisteTipoCompromisso", 0

End Function

Public Function fgExisteTipoLiquidacaoMensageria(ByVal plngTipoLiquidacaoMens As Long) As Boolean

Dim rsTipoLiquidacaoMens                    As ADODB.Recordset

On Error GoTo ErrorHandler
    
    Set rsTipoLiquidacaoMens = fgQueryDominio("TP_LIQU")
    
    With rsTipoLiquidacaoMens
    
        .Filter = " CO_DOMI = " & Format$(plngTipoLiquidacaoMens, "00")
    
        If .RecordCount > 0 Then
            glngTipoLiquidacaoMens = plngTipoLiquidacaoMens
            fgExisteTipoLiquidacaoMensageria = True
        End If
    End With
    
    Set rsTipoLiquidacaoMens = Nothing
    Exit Function
    
ErrorHandler:
    
    Set rsTipoLiquidacaoMens = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoLiquidacaoMensageria", 0

End Function

Public Function fgExisteOperacaoRotinaAbertura(ByVal plngOperacaoRotinaAbr As Long) As Boolean

Dim strSQL                                  As String
Dim rsControleDominio                       As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " SELECT   NO_ATRB, CO_DOMI, TP_CNTR_DOMI, DE_DOMI FROM A8.TB_CTRL_DOMI " & _
             " where    NO_ATRB = 'TP_OPER_ROTI_ABER' and CO_DOMI = '" & plngOperacaoRotinaAbr & "'"
    
    Set rsControleDominio = QuerySQL(strSQL)
    
    If rsControleDominio.RecordCount > 0 Then
        glngOperacaoRotinaAbr = plngOperacaoRotinaAbr
        fgExisteOperacaoRotinaAbertura = True
    End If
    
    Set rsControleDominio = Nothing
    
    Exit Function
ErrorHandler:
    
    Set rsControleDominio = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgExisteOperacaoRotinaAbertura", lngCodigoErroNegocio)

End Function

Public Function fgExisteTipoPagamento(ByVal pstrTipoPagto As String) As Boolean

Dim strSQL                                  As String
Dim rsControleDominio                       As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " SELECT   NO_ATRB, CO_DOMI, TP_CNTR_DOMI, DE_DOMI FROM A8.TB_CTRL_DOMI " & _
             " where    NO_ATRB = 'TP_PAGTO' AND CO_DOMI = '" & pstrTipoPagto & "'"
    
    Set rsControleDominio = QuerySQL(strSQL)
    
    With rsControleDominio
        If .RecordCount > 0 Then
            gstrTipoPagto = pstrTipoPagto
            fgExisteTipoPagamento = True
        End If
    End With
    Set rsControleDominio = Nothing
Exit Function
ErrorHandler:
    Set rsControleDominio = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoPagamento", 0

End Function

Public Function fgExisteTipoPagamentoRedesc(ByVal plngTipoPagtoRedesc As Long) As Boolean

Dim strSQL                                  As String
Dim rsControleDominio                       As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " SELECT   NO_ATRB, CO_DOMI, TP_CNTR_DOMI, DE_DOMI FROM A8.TB_CTRL_DOMI " & _
             " where    NO_ATRB = 'TP_PAGTO_RDSC' AND CO_DOMI = '" & plngTipoPagtoRedesc & "'"
    
    Set rsControleDominio = QuerySQL(strSQL)
    
    With rsControleDominio
        If .RecordCount > 0 Then
            glngTipoPagtoRedesc = plngTipoPagtoRedesc
            fgExisteTipoPagamentoRedesc = True
        End If
    End With
    Set rsControleDominio = Nothing
    
    Exit Function
ErrorHandler:
    Set rsControleDominio = Nothing
   fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoPagamentoRedesc", 0

End Function

Public Function fgExisteTipoLeilao(ByVal pvntTipoLeilao As Variant) As Boolean

Dim rsTipoLeilao                            As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsTipoLeilao = fgQueryDominio("TP_LEIL")
    
    With rsTipoLeilao
    
        .Filter = " CO_DOMI = " & Format$(pvntTipoLeilao, "00")
    
        If .RecordCount > 0 Then
            gvntTipoLeilao = pvntTipoLeilao
            fgExisteTipoLeilao = True
        End If
    End With
    
    Set rsTipoLeilao = Nothing
Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoLeilao", 0

End Function

Public Function fgExisteTipoPagtoLDL(ByVal plngTipoPagtoLDL As Variant) As Boolean

Dim rsTipoPgtoLDL                           As ADODB.Recordset

On Error GoTo ErrorHandler
   
    Set rsTipoPgtoLDL = fgQueryDominio("TP_PAGTO_LDL")
    
    With rsTipoPgtoLDL
    
        .Filter = " CO_DOMI = " & Format$(plngTipoPagtoLDL, "00")
    
        If .RecordCount > 0 Then
            glngTipoPgtoLDL = plngTipoPagtoLDL
            fgExisteTipoPagtoLDL = True
        End If
    End With
    
    Set rsTipoPgtoLDL = Nothing

Exit Function
ErrorHandler:
    Set rsTipoPgtoLDL = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoPagtoLDL", 0

End Function

Public Function fgExisteTipoCompromissoRetn(ByVal plngTipoCompromissoRetn As Long) As Boolean

Dim rsTipoCompromissoRetn                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    Set rsTipoCompromissoRetn = fgQueryDominio("TP_CPRO_RETN_OPER_ATIV")
    
    With rsTipoCompromissoRetn
    
        .Filter = " CO_DOMI = " & Format$(plngTipoCompromissoRetn, "00")
    
        If .RecordCount > 0 Then
            glngTipoCompromissoRetn = plngTipoCompromissoRetn
            fgExisteTipoCompromissoRetn = True
        End If
    End With
    
    Set rsTipoCompromissoRetn = Nothing
    
    Exit Function
ErrorHandler:
    
    Set rsTipoCompromissoRetn = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoCompromissoRetn", 0

End Function

Public Function fgExisteTipoMensagem(ByVal plngTipoMesg As Long, _
                                     ByVal pdatDataVigencia As Date) As Boolean
                                     
Dim strSQL                                  As String
Dim rsTipoMensagem                          As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   TP_MESG,           " & vbCrLf & _
             "          DT_INIC_VIGE_MESG, " & vbCrLf & _
             "          DT_FIM_VIGE_MESG   " & vbCrLf & _
             " FROM     A7.TB_TIPO_MESG      " & vbCrLf & _
             " WHERE    TP_MESG = '" & plngTipoMesg & "'"
    
    Set rsTipoMensagem = QuerySQL(strSQL)

    If rsTipoMensagem.RecordCount > 0 Then
        'Verifica a data de vigência
        If rsTipoMensagem!DT_INIC_VIGE_MESG <= pdatDataVigencia And _
            (IsNull(rsTipoMensagem!DT_FIM_VIGE_MESG) Or _
            pdatDataVigencia <= rsTipoMensagem!DT_FIM_VIGE_MESG) Then
            glngTipoMesg = plngTipoMesg
            gdatDataVigenciaTipoMesg = pdatDataVigencia
            fgExisteTipoMensagem = True
        End If
    End If
    
    Set rsTipoMensagem = Nothing
    
    Exit Function
ErrorHandler:

    Set rsTipoMensagem = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoMensagem", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
'
Public Function fgObterItemCaixaGenerico(ByVal pintTipoCaixa As enumTipoCaixa, ByVal plngTipoBackOffice As Long) As String

Dim strSQL                                  As String
Dim rsItemCaixaGenerico                     As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " SELECT CO_ITEM_CAIX, TP_CAIX, TP_BKOF " & vbNewLine & _
             " FROM   A6.TB_ITEM_CAIX_OPER_ATIV      " & vbNewLine & _
             " WHERE  DE_ITEM_CAIX = '" & gstrItemGenerico & "'" & _
             " AND    TP_CAIX = " & pintTipoCaixa & " AND TP_BKOF = " & plngTipoBackOffice
             
    Set rsItemCaixaGenerico = QuerySQL(strSQL)

    With rsItemCaixaGenerico
        If .RecordCount > 0 Then
            fgObterItemCaixaGenerico = !CO_ITEM_CAIX
        End If
    End With
    
    Set rsItemCaixaGenerico = Nothing
    
    Exit Function
ErrorHandler:
    
    Set rsItemCaixaGenerico = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgObterItemCaixaGenerico", 0

End Function

Public Function fgTipoSolicitacaoValido(ByVal plngTipoMensagem As Long, _
                                        ByVal plngTipoSolicitacao As Long) As Boolean

On Error GoTo ErrorHandler

    Select Case plngTipoMensagem
        Case enumTipoMensagemLQS.Definitiva, _
             enumTipoMensagemLQS.TermoD0, _
             enumTipoMensagemLQS.Leilao, _
             enumTipoMensagemLQS.VinculoDesvinculoTransf

             fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Alteracao

        Case enumTipoMensagemLQS.ConversaoRedesconto
        
             fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento
                                       
        Case enumTipoMensagemLQS.RegistroOperacaoBMA, _
             enumTipoMensagemLQS.LiquidacaoOperacoesBMA, _
             enumTipoMensagemLQS.IntermediacaoOperacoesInternasBMA, _
             enumTipoMensagemLQS.DespesasBMC

             fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Alteracao

        Case enumTipoMensagemLQS.Redesconto

             fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoPorLastro

        Case enumTipoMensagemLQS.Compromissada

             fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoPorLastro Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Alteracao

        Case enumTipoMensagemLQS.PgtoRedesconto
             
             fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoPorLastro

        Case enumTipoMensagemLQS.RetornoCompromissada

             fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Alteracao
             
        Case enumTipoMensagemLQS.TransferenciaLDL_BMA, _
             enumTipoMensagemLQS.EspecificacaoOperacoesBMA

             fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento
                                       
        Case enumTipoMensagemLQS.LiquidacaoEventosBMA
        
             fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Alteracao

        Case enumTipoMensagemLQS.TermoDataLiquidacao

             fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                       plngTipoSolicitacao = enumTipoSolicitacao.Alteracao

        Case enumTipoMensagemLQS.EventosSelic, _
             enumTipoMensagemLQS.DespesasSelic

            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Alteracao

        Case enumTipoMensagemLQS.ExercicioOpcaoContratoSwapCETIP
            
            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento
                                      
        Case enumTipoMensagemLQS.OperacoesComCorretorasCETIP
        
            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Alteracao
            
        Case enumTipoMensagemLQS.OperacaoCompromissadaCETIP
             
            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Alteracao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoPorLastro

        Case enumTipoMensagemLQS.OperacaoRetornoAntecipacaoCETIP, _
             enumTipoMensagemLQS.AntecipacaoResgateContratoDerivativoCETIP
             
            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Alteracao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem

        Case enumTipoMensagemLQS.MovimentacoesCustodiaCETIP, _
             enumTipoMensagemLQS.ExercicioDesistenciaCETIP, _
             enumTipoMensagemLQS.ConversaoPermutaValorImobCETIP, _
             enumTipoMensagemLQS.ConversaoPermutaValorImobCETIP, _
             enumTipoMensagemLQS.EspecificacaoQuantidadesCotasCETIP, _
             enumTipoMensagemLQS.RegistroContratoSWAP, _
             enumTipoMensagemLQS.MovimentacoesContratoDerivativo
             
            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem
                                      
        Case enumTipoMensagemLQS.MovimentacoesInstFinancCETIP, _
             enumTipoMensagemLQS.ResgateFundoInvestimentoCETIP, _
             enumTipoMensagemLQS.OperacaoDefinitivaCETIP, _
             enumTipoMensagemLQS.RegistroContratoTermoCETIP, _
             enumTipoMensagemLQS.RegistroContratoSWAPCetip21
        
            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Alteracao

        Case enumTipoMensagemLQS.EspecificacaoQuantidadesCotasCETIP
             
            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem

        Case enumTipoMensagemLQS.OperacaoRetencaoIRF_CETIP, _
             enumTipoMensagemLQS.RegistroOperacaoesCETIP, _
             enumTipoMensagemLQS.LancamentoPU_CETIP, _
             enumTipoMensagemLQS.EventoJurosCETIP

            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Alteracao

        Case enumTipoMensagemLQS.DespesasCETIP

            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Alteracao

        Case enumTipoMensagemLQS.AlteracaoDadosContaCorrente

            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Alteracao

        Case enumTipoMensagemLQS.RegistroLiquidacaoMultilateralCBLC, _
             enumTipoMensagemLQS.RegistroLiquidacaoBrutaCBLC, _
             enumTipoMensagemLQS.RegistroLiquidacaoEventoCBLC, _
             enumTipoMensagemLQS.RegistroLiquidacaoMultilateralBMF
            
            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Alteracao

        Case enumTipoMensagemLQS.TransferenciasBMC, enumTipoMensagemLQS.RegistroOperacoesRodaDolar

            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento

        Case enumTipoMensagemLQS.RegistroOperacoesBMC

            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Reativacao

        Case enumTipoMensagemLQS.LiquidacaoMultilateralBMC

            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Alteracao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento

        Case enumTipoMensagemLQS.EnvioTEDClientes, _
             enumTipoMensagemLQS.EnvioPagDespesas, _
             enumTipoMensagemLQS.LancamentoContaCorrenteBG

            fgTipoSolicitacaoValido = plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
                                      plngTipoSolicitacao = enumTipoSolicitacao.Alteracao
       
        Case Else
            
            fgTipoSolicitacaoValido = True

    End Select

Exit Function
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgTipoSolicitacaoValido", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Public Function fgExisteTipoConta(ByVal plngCodigoEmpresaFusionada As Long, _
                                  ByVal plngTipoConta As Long, _
                                  ByVal pdatDataVigencia As Date) As Boolean

Dim strSQL                                  As String
Dim rsTipoConta                             As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_TIPO_CNTA, CO_EMPR_FUSI, DT_INIC_VIGE, DT_FIM_VIGE " & _
             " FROM     A8.TB_TIPO_CONTA " & _
             " WHERE    CO_EMPR_FUSI = " & plngCodigoEmpresaFusionada & " AND CO_TIPO_CNTA = " & plngTipoConta
    
    Set rsTipoConta = QuerySQL(strSQL)

    If rsTipoConta.RecordCount > 0 Then
        If rsTipoConta!dt_inic_vige <= pdatDataVigencia And (IsNull(rsTipoConta!DT_FIM_VIGE) Or pdatDataVigencia <= rsTipoConta!DT_FIM_VIGE) Then
            glngTipoConta = plngTipoConta
            gdatDataVigenciaTipoConta = pdatDataVigencia
            fgExisteTipoConta = True
        End If
    End If
    
    Set rsTipoConta = Nothing
    
    Exit Function

ErrorHandler:
    
    Set rsTipoConta = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoConta", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
                                  
Public Function fgExisteSegmento(ByVal plngCodigoEmpresaFusionada As Long, _
                                 ByVal plngCodigoSegmento As Long, _
                                 ByVal pdatDataVigencia As Date) As Boolean


Dim strSQL                                  As String
Dim rsSegmento                              As ADODB.Recordset
    
On Error GoTo ErrorHandler

    strSQL = "select    co_empr_fusi,co_segm, dt_inic_vige, dt_fim_vige " & vbCrLf & _
             " from     a8.tb_segmento" & _
             " where    co_empr_fusi = " & plngCodigoEmpresaFusionada & " AND co_segm = " & plngCodigoSegmento
    
    Set rsSegmento = QuerySQL(strSQL)

    If rsSegmento.RecordCount > 0 Then
        If rsSegmento!dt_inic_vige <= pdatDataVigencia And (IsNull(rsSegmento!DT_FIM_VIGE) Or pdatDataVigencia <= rsSegmento!DT_FIM_VIGE) Then
            glngCodigoSegmento = plngCodigoSegmento
            gdatDataVigenciaSegmento = pdatDataVigencia
            fgExisteSegmento = True
        End If
    End If
    Set rsSegmento = Nothing
Exit Function
ErrorHandler:
    Set rsSegmento = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgExisteSegmento", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

Public Function fgExisteEventoFinanceiro(ByVal plngCodigoEmpresaFusionada As Long, _
                                         ByVal plngCodigoEventoFinaceiro As Long, _
                                         ByVal pdatDataVigencia As Date) As Boolean

Dim strSQL                                  As String
Dim rsEventoFinanceiro                      As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = "select co_empr_fusi, co_even_finc, dt_inic_vige, dt_fim_vige " & vbCrLf & _
             " from a8.tb_evento_financeiro " & _
             " where co_empr_fusi = " & plngCodigoEmpresaFusionada & " AND co_even_finc = " & plngCodigoEventoFinaceiro
    
    Set rsEventoFinanceiro = QuerySQL(strSQL)

    If rsEventoFinanceiro.RecordCount > 0 Then
        If rsEventoFinanceiro!dt_inic_vige <= pdatDataVigencia And (IsNull(rsEventoFinanceiro!DT_FIM_VIGE) Or pdatDataVigencia <= rsEventoFinanceiro!DT_FIM_VIGE) Then
            glngCodigoEventoFinanceiro = plngCodigoEventoFinaceiro
            gdatDataVigenciaEventoFinanceiro = pdatDataVigencia
            fgExisteEventoFinanceiro = True
        End If
    End If
    
    Set rsEventoFinanceiro = Nothing
    
    Exit Function
    
ErrorHandler:
    
    Set rsEventoFinanceiro = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgExisteEventoFinanceiro", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

Public Function fgExisteIndexador(ByVal plngCodigoEmpresaFusionada As Long, _
                                         ByVal plngCodigoIndexador As Long, _
                                         ByVal pdatDataVigencia As Date) As Boolean

Dim strSQL                                  As String
Dim rsIndexador                             As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = "select co_empr_fusi, co_indx, dt_inic_vige, dt_fim_vige " & vbCrLf & _
             " from a8.tb_indexador " & _
             " where co_empr_fusi = " & plngCodigoEmpresaFusionada & " AND co_indx = " & plngCodigoIndexador
    
    Set rsIndexador = QuerySQL(strSQL)

    If rsIndexador.RecordCount > 0 Then
        If rsIndexador!dt_inic_vige <= pdatDataVigencia And (IsNull(rsIndexador!DT_FIM_VIGE) Or pdatDataVigencia <= rsIndexador!DT_FIM_VIGE) Then
            glngCodigoIndexador = plngCodigoIndexador
            gdatDataVigenciaIndexador = pdatDataVigencia
            fgExisteIndexador = True
        End If
    End If
    Set rsIndexador = Nothing
Exit Function
ErrorHandler:
    Set rsIndexador = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgExisteIndexador", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

Public Function fgExisteTipoVinculo(ByVal plngTipoVinculo As Long) As Boolean

Dim strSQL                                  As String
Dim rsControleDominio                       As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " SELECT   NO_ATRB, CO_DOMI, TP_CNTR_DOMI, DE_DOMI FROM A8.TB_CTRL_DOMI " & _
             " where    NO_ATRB = 'TP_VINC_DSVN_TRAF' and CO_DOMI = '" & plngTipoVinculo & "'"
    
    Set rsControleDominio = QuerySQL(strSQL)
    
    If rsControleDominio.RecordCount > 0 Then
        glngTipoVinculo = plngTipoVinculo
        fgExisteTipoVinculo = True
    End If
    
    Set rsControleDominio = Nothing
    
    Exit Function

ErrorHandler:
    Set rsControleDominio = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoVinculo", lngCodigoErroNegocio)

End Function

Public Function fgExisteTipoRedesconto(ByVal plngTipoRedesconto As Long) As Boolean

Dim strSQL                                  As String
Dim rsControleDominio                       As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " SELECT   NO_ATRB, CO_DOMI, TP_CNTR_DOMI, DE_DOMI FROM A8.TB_CTRL_DOMI " & _
             " where    NO_ATRB = 'TP_RDSC' and CO_DOMI = '" & plngTipoRedesconto & "'"
    
    Set rsControleDominio = QuerySQL(strSQL)
    
    If rsControleDominio.RecordCount > 0 Then
        glngTipoRedesconto = plngTipoRedesconto
        fgExisteTipoRedesconto = True
    End If
    Set rsControleDominio = Nothing
Exit Function
ErrorHandler:
    Set rsControleDominio = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoRedesconto", lngCodigoErroNegocio)

End Function


Public Function fgConsisteLiquidacaoOperacao_BMA(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim lngCodigoEmpresaFusi                    As Long
Dim datDataVigencia                         As Date
Dim blnEntradaManual                        As Boolean
Dim lngTipoSolicitacao                      As Long
Dim lngTipoMensagem                         As Long
Dim lngTemp                                 As Long

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")
    
        lngTipoMensagem = CLng("0" & .selectSingleNode("TP_MESG").Text)

        'Empresa
        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        End If
        lngCodigoEmpresaFusi = glngCodigoEmpresaFusionada

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
        
            If Not blnEntradaManual Then
                If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                               objDomNode.Text) Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If

        'Data da Operação
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da operação inválida
                fgAdicionaErro xmlErrosNegocio, 4024
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                    'Data da operação não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4059
                Else
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da operação inválida
                        fgAdicionaErro xmlErrosNegocio, 4024
                    End If
                End If
            End If
        End If
        
        'Data da mensagem
        Set objDomNode = .selectSingleNode("DT_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(.selectSingleNode("DT_MESG").Text)) Then
                    'Data da Mensagem inválida
                    fgAdicionaErro xmlErrosNegocio, 4060
                Else
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Mensagem inválida
                        fgAdicionaErro xmlErrosNegocio, 4060
                    End If
                
                End If
            End If
        End If

        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.BMA Then
                'Código do local de liquidação inválido
                fgAdicionaErro xmlErrosNegocio, 4008
            End If
        End If

        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            Else
                'Produto inválido
                fgAdicionaErro xmlErrosNegocio, 4011
            End If
        End If

        'Tipo de Conta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(lngCodigoEmpresaFusi, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If

        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(lngCodigoEmpresaFusi, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If
        
        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(lngCodigoEmpresaFusi, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If

        'Código do Índexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If
        
        'Tipo de Operação
        Set objDomNode = .selectSingleNode("TP_OPER_BMA")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumTipoOperacaoBMA.TermoNaDataLiquidacao And _
               Val("0" & objDomNode.Text) <> enumTipoOperacaoBMA.RetornoCompromissada Then
                'Tipo de operação BMA inválido
                fgAdicionaErro xmlErrosNegocio, 4077
            End If
        End If
        
        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                            objDomNode.Text, _
                                            datDataVigencia, _
                                            .selectSingleNode("CO_EMPR").Text) Then
                    'Veiculo legal inválido
                    fgAdicionaErro xmlErrosNegocio, 4007
                End If
            End If
        End If

        'Indicador débito/crédito
        Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                lngTemp = Val(objDomNode.Text)
                If lngTemp <> enumTipoDebitoCredito.Credito And _
                   lngTemp <> enumTipoDebitoCredito.Debito Then
                    'Indicador de débito/crédito inválido
                    fgAdicionaErro xmlErrosNegocio, 4021
                End If
            End If
        End If

        'Data de vencimento
        Set objDomNode = .selectSingleNode("DT_VENC_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data de vencimento inválida
                fgAdicionaErro xmlErrosNegocio, 4063
            End If
        End If

        'Quantidade do título
        Set objDomNode = .selectSingleNode("QT_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) = 0 Then
                'Quantidade de Títulos deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4050
            End If
        End If

        'Preço unitário
        Set objDomNode = .selectSingleNode("PU_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                'Preço Unitário deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4053
            End If
        End If
    
        'Forma de Liquidação
        Set objDomNode = .selectSingleNode("CO_FORM_LIQU")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumFormaLiquidacao.ContaCorrente And _
               Val("0" & objDomNode.Text) <> enumFormaLiquidacao.Contabil Then
                'Forma de Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4026
            End If
        End If

        #If ValidaCC = 1 Then
        
            If .selectSingleNode("CO_FORM_LIQU").Text = enumFormaLiquidacao.ContaCorrente And _
              (lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               lngTipoSolicitacao = enumTipoSolicitacao.Alteracao) Then
        
                'Código do banco
                Set objDomNode = .selectSingleNode("CO_BANC")
                If Not objDomNode Is Nothing Then
                    If Not fgExisteCodigoBanco(Val(objDomNode.Text)) Then
                        'Código do Banco inválido
                        fgAdicionaErro xmlErrosNegocio, 4027
                    End If
                End If
    
                'Código da Agência
                Set objDomNode = .selectSingleNode("CO_AGEN")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da Agência inválido
                        fgAdicionaErro xmlErrosNegocio, 4028
                    End If
                End If
    
                'Número da Conta Corrente
                Set objDomNode = .selectSingleNode("NU_CC")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4029
                    End If
                End If
    
                'Valor do Lançamento Conta Corrente
                Set objDomNode = .selectSingleNode("VA_LANC_CC")
                If Not objDomNode Is Nothing Then
                    If fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                        'Valor do lançamento na conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4030
                    End If
                End If
            End If
        #End If
    End With

    fgConsisteLiquidacaoOperacao_BMA = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteLiquidacaoOperacao_BMA", 0

End Function

Public Function fgConsisteOperacaoTransfCamaraLDL_BMA(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim lngCodigoEmpresaFusi                    As Long
Dim datDataVigencia                         As Date
Dim blnEntradaManual                        As Boolean
Dim lngTipoSolicitacao                      As Long
Dim lngTipoTransferencia                    As Long
Dim lngTipoMensagem                         As Long
Dim lngTemp                                 As Long

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")

        lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)

        'Empresa
        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        End If
        lngCodigoEmpresaFusi = glngCodigoEmpresaFusionada

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
        
            If Not blnEntradaManual Then
                If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                               objDomNode.Text) Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If
        
        'Hora de agendamento
        Set objDomNode = .selectSingleNode("HO_AGND")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 And _
            Not flValidaHoraNumerica(objDomNode.Text) Then
                'Horário de Agendamento inválido
                fgAdicionaErro xmlErrosNegocio, 4062
            End If
        End If

        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.BMA And _
               Val("0" & objDomNode.Text) <> enumLocalLiquidacao.BMC And _
               Val("0" & objDomNode.Text) <> enumLocalLiquidacao.BMD And _
               Val("0" & objDomNode.Text) <> enumLocalLiquidacao.CETIP Then
                'Código do local de liquidação inválido
                fgAdicionaErro xmlErrosNegocio, 4008
            End If
        End If

        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                            objDomNode.Text, _
                                            datDataVigencia, _
                                            .selectSingleNode("CO_EMPR").Text) Then
                    'Veiculo legal inválido
                    fgAdicionaErro xmlErrosNegocio, 4007
                End If
            End If
        End If

        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            Else
                If glngTipoBackOffice = enumTipoBackOffice.Tesouraria Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            End If
        End If

        'Tipo de Conta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(lngCodigoEmpresaFusi, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If

        'Data da Operação
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da operação inválida
                fgAdicionaErro xmlErrosNegocio, 4024
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                    'Data da operação não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4059
                Else
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da operação inválida
                        fgAdicionaErro xmlErrosNegocio, 4024
                    End If
                End If
            End If
        End If
        
        'Data da mensagem
        Set objDomNode = .selectSingleNode("DT_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(.selectSingleNode("DT_MESG").Text)) Then
                    'Data da Mensagem inválida
                    fgAdicionaErro xmlErrosNegocio, 4060
                Else
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Mensagem inválida
                        fgAdicionaErro xmlErrosNegocio, 4060
                    End If
                
                End If
            End If
        End If

        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(lngCodigoEmpresaFusi, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If
        
        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(lngCodigoEmpresaFusi, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If

        'Código do Índexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If

        'Tipo Transferência LDL
        Set objDomNode = .selectSingleNode("TP_TRAF_LDL")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumTipoTransferenciaLDL.Depósito And _
               Val("0" & objDomNode.Text) <> enumTipoTransferenciaLDL.Retirada And _
               Val("0" & objDomNode.Text) <> enumTipoTransferenciaLDL.TrasnfEntreClearings Then
                'Tipo de Transferência LDL inválido
                fgAdicionaErro xmlErrosNegocio, 4040
            End If
            lngTipoTransferencia = Val("0" & objDomNode.Text)
        End If

        'Indicador débito/crédito
        Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                lngTemp = Val(objDomNode.Text)
                If lngTemp <> enumTipoDebitoCredito.Credito And _
                   lngTemp <> enumTipoDebitoCredito.Debito Then
                    'Indicador de débito/crédito inválido
                    fgAdicionaErro xmlErrosNegocio, 4021
                End If
            End If
        End If

        'Data de vencimento
        Set objDomNode = .selectSingleNode("DT_VENC_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data de vencimento inválida
                fgAdicionaErro xmlErrosNegocio, 4063
            End If
        End If

        'Quantidade do título
        Set objDomNode = .selectSingleNode("QT_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) = 0 Then
                'Quantidade de Títulos deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4050
            End If
        End If

        'Preço unitário
        Set objDomNode = .selectSingleNode("PU_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                'Preço Unitário deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4053
            End If
        End If
    
    
        'Número Operação Selic
        Set objDomNode = .selectSingleNode("NU_COMD_OPER")
        If Not objDomNode Is Nothing Then
            If Len(objDomNode.Text) > 6 Or (Not IsNumeric(objDomNode.Text)) Then
                'Número do comando inválido
                fgAdicionaErro xmlErrosNegocio, 4100
            End If
        End If
    
        'Valor Financeiro
        Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                'Valor Financeiro deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4049
            End If
        End If
        
        'Quantidade mínima de título
        Set objDomNode = .selectSingleNode("QT_MINI_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 And _
                lngTipoTransferencia = enumTipoTransferenciaLDL.Retirada Then
                'Quantidade minima de título deve ser mair que zero
                fgAdicionaErro xmlErrosNegocio, 4081
            End If
        End If

        'Tipo Requisição
        Set objDomNode = .selectSingleNode("TP_REQU")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal(objDomNode.Text) <> enumTipoRequisicaoBMA.Requisitado And _
               fgVlrXml_To_Decimal(objDomNode.Text) <> enumTipoRequisicaoBMA.Simulacao And _
               fgVlrXml_To_Decimal(objDomNode.Text) <> enumTipoRequisicaoBMA.Voluntario Then
                'Tipo Requisição inválido
                fgAdicionaErro xmlErrosNegocio, 4082
            End If
        End If
        
        'Tipo de Ativo
        Set objDomNode = .selectSingleNode("TP_ATIV")
        If Not fgExisteTipoAtivo(Val("0" & objDomNode.Text)) Then
            'Tipo de Ativo inexistente
            fgAdicionaErro xmlErrosNegocio, 4128
        End If

        'Tipo Titular de origem
        Set objDomNode = .selectSingleNode("TP_TITL_ORIG")
        If Not objDomNode Is Nothing Then
            If Not fgExisteTipoTitular(objDomNode.Text) Then
                'Tipo Titular de origem inválido
                fgAdicionaErro xmlErrosNegocio, 4083
            End If
        End If

        'Tipo Subtitular de origem
        Set objDomNode = .selectSingleNode("TP_SUB_TITL_ORIG")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                If Not fgExisteTipoTitular(objDomNode.Text) Then
                    'Tipo SubTitular de origem inválido
                    fgAdicionaErro xmlErrosNegocio, 4084
                End If
            End If
        End If

        'Finalidade Cobertura Conta
        Set objDomNode = .selectSingleNode("CO_FIND_COBE")
        If Not objDomNode Is Nothing Then
            If Not fgExisteFinalidadeCoberturaConta(Val("0" & objDomNode.Text)) Then
                'Finalidade Cobertura Conta inválida
                fgAdicionaErro xmlErrosNegocio, 4085
            End If
        End If

        'Tipo titular destino
        Set objDomNode = .selectSingleNode("TP_TITL_DEST")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                If Not fgExisteTipoTitular(objDomNode.Text) Then
                    'Tipo Titular de destino inválido
                    fgAdicionaErro xmlErrosNegocio, 4086
                End If
            End If
        End If

        'Tipo subtitular destino
        Set objDomNode = .selectSingleNode("TP_SUB_TITL_DEST")
        
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                If Not fgExisteTipoTitular(objDomNode.Text) Then
                    'Tipo subtitular destino inválido
                    fgAdicionaErro xmlErrosNegocio, 4087
                End If
            End If
        End If
        
        'Finalidade Cobertura destino
        Set objDomNode = .selectSingleNode("CO_FIND_COBE_DEST")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteFinalidadeCoberturaConta(Val("0" & objDomNode.Text)) Then
                    'Finalidade Cobertura destino inválida
                    fgAdicionaErro xmlErrosNegocio, 4085
                End If
            End If
        End If
    
        If lngTipoTransferencia = enumTipoTransferenciaLDL.TrasnfEntreClearings Then
        
            'Conta Custodia Local Liquidacao Origem
            Set objDomNode = .selectSingleNode("//CO_CNTA_CUTD_SELIC_LOCA_LIQU")
            
            If objDomNode Is Nothing Then
                'Conta Custodia Local Liquidacao Origem inválida
                fgAdicionaErro xmlErrosNegocio, 4152
            Else
                If IsNumeric(objDomNode.Text) Then
                    If Val(objDomNode.Text) = 0 Then
                        'Conta Custodia Local Liquidacao Origem inválida
                        fgAdicionaErro xmlErrosNegocio, 4152
                    End If
                Else
                    'Conta Custodia Local Liquidacao Origem inválida
                    fgAdicionaErro xmlErrosNegocio, 4152
                End If
            End If
        End If
        
        If lngTipoTransferencia = enumTipoTransferenciaLDL.TrasnfEntreClearings Then
        
            'Conta Custodia Local Liquidacao Destino
            Set objDomNode = .selectSingleNode("//CO_CNTA_CUTD_SELIC_LOCA_LIQU_DEST")
            If objDomNode Is Nothing Then
                'Conta Custodia Local Liquidacao Destino inválida
                fgAdicionaErro xmlErrosNegocio, 4153
            Else
                If IsNumeric(objDomNode.Text) Then
                    If Val(objDomNode.Text) = 0 Then
                        'Conta Custodia Local Liquidacao Destino inválida
                        fgAdicionaErro xmlErrosNegocio, 4153
                    End If
                Else
                    'Conta Custodia Local Liquidacao Destino inválida
                    fgAdicionaErro xmlErrosNegocio, 4153
                End If
            End If
        End If
    
    End With
    
    fgConsisteOperacaoTransfCamaraLDL_BMA = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteOperacaoTransfCamaraLDL_BMA", 0


End Function

Public Function fgExisteTipoAtivo(ByVal plngTipoAtivo As Long) As Boolean

Dim rsTipoAtivo                             As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsTipoAtivo = fgQueryDominioInterno("TP_ATIV")
    
    With rsTipoAtivo
    
        .Filter = " CO_DOMI = '" & plngTipoAtivo & "' "
    
        If .RecordCount > 0 Then
            glngTipoAtivo = plngTipoAtivo
            fgExisteTipoAtivo = True
        End If
    End With
    Set rsTipoAtivo = Nothing
Exit Function
ErrorHandler:
    Set rsTipoAtivo = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoAtivo", 0

End Function

Public Function fgConsisteEspecificacaoOperacoes_BMA(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim lngCodigoEmpresa                        As Long
Dim lngCodigoEmpresaFusi                    As Long
Dim datDataVigencia                         As Date
Dim blnEntradaManual                        As Boolean
Dim lngTipoSolicitacao                      As Long
Dim objNodeRepet                            As MSXML2.IXMLDOMNode
Dim objNodeRepetTitulo                      As MSXML2.IXMLDOMNode
Dim lngTipoMensagem                         As Long
Dim lngTemp                                 As Long
Dim lngTipoNegociacao                       As Long
Dim blnQuantidade                           As Boolean
Dim strNumComdOper                          As String
Dim lngTipoEspec                            As Long
Dim strNumCtrlEspec                         As String
Dim strCodigoVeiculoLegal                   As String

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")
    
    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")

        lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)
        
        'Empresa
        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        Else
            lngCodigoEmpresa = Val(.selectSingleNode("CO_EMPR").Text)
        End If
        
        lngCodigoEmpresaFusi = glngCodigoEmpresaFusionada

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
        
            If Not blnEntradaManual Then
                If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                               objDomNode.Text) Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If

        'Data da Operação
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da operação inválida
                fgAdicionaErro xmlErrosNegocio, 4024
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                    'Data da operação não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4059
                Else
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da operação inválida
                        fgAdicionaErro xmlErrosNegocio, 4024
                    End If
                End If
            End If
        End If
        
        'Data da mensagem
        Set objDomNode = .selectSingleNode("DT_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(.selectSingleNode("DT_MESG").Text)) Then
                    'Data da Mensagem inválida
                    fgAdicionaErro xmlErrosNegocio, 4060
                Else
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Mensagem inválida
                        fgAdicionaErro xmlErrosNegocio, 4060
                    End If
                
                End If
            End If
        End If
        
        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("//CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                            objDomNode.Text, _
                                            datDataVigencia, _
                                            .selectSingleNode("CO_EMPR").Text) Then
                    'Veiculo legal inválido
                    fgAdicionaErro xmlErrosNegocio, 4007
                Else
                    strCodigoVeiculoLegal = objDomNode.Text
                End If
            End If
        End If
        
        'Número Operação Negociação BMA inválido.
        Set objDomNode = .selectSingleNode("NU_COMD_OPER")
        If Not objDomNode Is Nothing Then
            If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao And _
               Trim$(objDomNode.Text) = vbNullString Then
                'Número Operação Negociação BMA inválido.
                fgAdicionaErro xmlErrosNegocio, 4157
            Else
                strNumComdOper = Trim(objDomNode.Text)
            End If
        End If
        
        'Tipo de Negociação BMA
        Set objDomNode = .selectSingleNode("TP_NEGO_BMA")
        
        If Not objDomNode Is Nothing Then
            If Not fgExisteTipoNegocicaoBMA(Val("0" & objDomNode.Text)) Then
                'Tipo de Negociação BMA inválido
                fgAdicionaErro xmlErrosNegocio, 4092
            Else
                lngTipoNegociacao = Val("0" & objDomNode.Text)
            End If
        End If
        
        'Tipo especificação
        Set objDomNode = .selectSingleNode("TP_ESFC")
        If Val("0" & objDomNode.Text) <> enumTipoEspecificacao.Cancelamento And _
           Val("0" & objDomNode.Text) <> enumTipoEspecificacao.Cobertura And _
           Val("0" & objDomNode.Text) <> enumTipoEspecificacao.Intermediacao Then
            'Tipo de especificação inválida
            fgAdicionaErro xmlErrosNegocio, 4097
        Else
            lngTipoEspec = objDomNode.Text
        End If
        
        'Número controle BMA especificação original
        Set objDomNode = .selectSingleNode("NU_CTRL_MESG_SPB_ORIG")
        
        If Not objDomNode Is Nothing Then
            If lngTipoEspec = enumTipoEspecificacao.Cancelamento Then
                If Trim(objDomNode.Text) = vbNullString Then
                    'Número controle BMA especificação original
                    fgAdicionaErro xmlErrosNegocio, 4159
                Else
                    strNumCtrlEspec = objDomNode.Text
                End If
            End If
        End If

        'Número de controle mensagem SPB original
        blnQuantidade = False
        
        If lngTipoEspec = enumTipoEspecificacao.Cancelamento Then
            strNumComdOper = strNumCtrlEspec
        End If
        
        If strNumComdOper <> "" Then
            Call fgValidaMensagemEspecificacao(strNumComdOper, _
                                              lngTipoEspec, _
                                              .selectSingleNode("//QT_ATIV_MERC").Text, _
                                              lngTipoNegociacao, _
                                              xmlErrosNegocio, _
                                              xmlDOMMensagem, _
                                              lngCodigoEmpresa, _
                                              strCodigoVeiculoLegal)
        End If

        'Hora de envio
        Set objDomNode = .selectSingleNode("HO_REME")
        If Not objDomNode Is Nothing Then
            If Not flValidaHoraNumerica(objDomNode.Text) Then
                'Hora da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4061
            End If
        End If

        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.BMA Then
                'Código do local de liquidação inválido
                fgAdicionaErro xmlErrosNegocio, 4008
            End If
        End If

        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            Else
                If lngTipoEspec = enumTipoEspecificacao.Intermediacao Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            End If
        End If

        'Tipo de Conta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(lngCodigoEmpresaFusi, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If

        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(lngCodigoEmpresaFusi, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If
        
        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(lngCodigoEmpresaFusi, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If

        'Código do Índexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If

        'Número Operação negociação BMA
        Set objDomNode = .selectSingleNode("NU_COMD_OPER_BMA")
        If Not objDomNode Is Nothing Then
            If lngTipoSolicitacao = enumTipoSolicitacao.Cancelamento And _
               Trim$(objDomNode.Text) = vbNullString Then
                'Número do comando da especifição obrigatório
                fgAdicionaErro xmlErrosNegocio, 4089
            End If
        End If

        'Tipo Coberto/Descoberto
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo Coberto/Descoberto inválido
                    fgAdicionaErro xmlErrosNegocio, 4090
                End If
            End If
        End If

        'Indicador débito/crédito
        Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                lngTemp = Val(objDomNode.Text)
                If lngTemp <> enumTipoDebitoCredito.Credito And _
                   lngTemp <> enumTipoDebitoCredito.Debito Then
                    'Indicador de débito/crédito inválido
                    fgAdicionaErro xmlErrosNegocio, 4021
                End If
            End If
        End If
        
        'Valor Financeiro Total
        Set objDomNode = .selectSingleNode("QT_TOTL_ATIV_MERC")
        If Val("0" & objDomNode.Text) <= 0 Then
            'Valor Financeiro Total deve ser maior que zero
            fgAdicionaErro xmlErrosNegocio, 4129
        End If
        
        'Quantidade total título
        Set objDomNode = .selectSingleNode("QT_TOTL_ATIV_MERC")
        If Val("0" & objDomNode.Text) <= 0 Then
            'Quantidade Total Título inválida
            fgAdicionaErro xmlErrosNegocio, 4130
        End If

        Set objDomNode = .selectSingleNode("REPE_CNTA")
        If Not objDomNode Is Nothing Then
            For Each objNodeRepet In .selectSingleNode("REPE_CNTA").childNodes
                With objNodeRepet
                    
                    'Finalidade Cobertura Conta
                    Set objDomNode = .selectSingleNode("CO_FIND_COBE")
                    If Not objDomNode Is Nothing Then
                        If Not fgExisteFinalidadeCoberturaConta(Val("0" & objDomNode.Text)) Then
                            'Finalidade Cobertura Conta inválida
                            fgAdicionaErro xmlErrosNegocio, 4085
                        End If
                    End If
                    
                    'Tipo Titular da Conta
                    Set objDomNode = .selectSingleNode("TP_TITL_CONTA")
                    If Not objDomNode Is Nothing Then
                        If Not fgExisteTipoTitular(objDomNode.Text) Then
                            'Tipo Titular conta inválido
                            fgAdicionaErro xmlErrosNegocio, 4091
                        End If
                    End If

                    'Valor Financeiro
                    Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
                    If Not objDomNode Is Nothing Then
                        If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                            'Valor Financeiro deve ser maior que zero
                            fgAdicionaErro xmlErrosNegocio, 4049
                        End If
                    End If

                    Set objNodeRepetTitulo = .selectSingleNode("Repeat_Titulo")
                    If Not objNodeRepet Is Nothing Then
                        If lngTipoNegociacao = enumTipoNegociacaoBMA.Compromissada Then
        
                            'Identificador de título SELIC
                            Set objDomNode = .selectSingleNode("DE_ATIV_MERC")
                            If Not objDomNode Is Nothing Then
                                If Trim$(objDomNode.Text) = vbNullString Then
                                    'Descrição Título Selic inválido
                                    fgAdicionaErro xmlErrosNegocio, 4075
                                End If
                            End If
                            
                            'Data de vencimento
                            Set objDomNode = .selectSingleNode("DT_VENC_ATIV")
                            If Not objDomNode Is Nothing Then
                                If Not flValidaDataNumerica(objDomNode.Text) Then
                                    'Data de vencimento inválida
                                    fgAdicionaErro xmlErrosNegocio, 4063
                                End If
                            End If
    
                            'Quantidade do título
                            Set objDomNode = .selectSingleNode("QT_ATIV_MERC")
                            If Not objDomNode Is Nothing Then
                                If Val("0" & objDomNode.Text) = 0 Then
                                    'Quantidade de Títulos deve ser maior que zero
                                    fgAdicionaErro xmlErrosNegocio, 4050
                                End If
                            End If
                
                            'Preço unitário
                            Set objDomNode = .selectSingleNode("PU_ATIV_MERC")
                            If Not objDomNode Is Nothing Then
                                If fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                                    'Preço Unitário deve ser maior que zero
                                    fgAdicionaErro xmlErrosNegocio, 4053
                                End If
                            End If
                        End If
                    End If
                End With
            Next objNodeRepet
        End If

    End With

    fgConsisteEspecificacaoOperacoes_BMA = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteEspecificacaoOperacoes_BMA", 0

End Function

Public Function fgConsisteLiquidacaoFisicaOperacao(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim lngCodigoEmpresaFusi                    As Long
Dim datDataVigencia                         As Date
Dim blnEntradaManual                        As Boolean
Dim lngTipoSolicitacao                      As Long
Dim objDomRepeat                            As MSXML2.IXMLDOMNode
Dim objDomRepeatConta                       As MSXML2.IXMLDOMNode
Dim lngTipoMensagem                         As Long
Dim lngTemp                                 As Long

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")
    
        lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)

        'Empresa
        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        End If
        lngCodigoEmpresaFusi = glngCodigoEmpresaFusionada

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
        
            If Not blnEntradaManual Then
                If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                               objDomNode.Text) Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If

        'Tipo de Negociação BMA
        Set objDomNode = .selectSingleNode("TP_NEGO_BMA")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoNegocicaoBMA(objDomNode.Text) Then
                    'Tipo de Negociação BMA inválido
                    fgAdicionaErro xmlErrosNegocio, 4092
                End If
            End If
        End If
        
        'Data da Operação
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da operação inválida
                fgAdicionaErro xmlErrosNegocio, 4024
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                    'Data da operação não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4059
                Else
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da operação inválida
                        fgAdicionaErro xmlErrosNegocio, 4024
                    End If
                End If
            End If
        End If
        
        'Data da mensagem
        Set objDomNode = .selectSingleNode("DT_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(.selectSingleNode("DT_MESG").Text)) Then
                    'Data da Mensagem inválida
                    fgAdicionaErro xmlErrosNegocio, 4060
                Else
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Mensagem inválida
                        fgAdicionaErro xmlErrosNegocio, 4060
                    End If
                
                End If
            End If
        End If

        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
              If Not Trim$(objDomNode.Text) = vbNullString Then
                If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                            objDomNode.Text, _
                                            datDataVigencia, _
                                            .selectSingleNode("CO_EMPR").Text) Then
                    'Veiculo legal inválido
                    fgAdicionaErro xmlErrosNegocio, 4007
                End If
            End If
        End If

        'Repetição Título
        If Not .selectSingleNode("REPE_TITL") Is Nothing Then
            For Each objDomRepeat In .selectSingleNode("REPE_TITL").childNodes
                With objDomRepeat
                    
                    'Data de vencimento
                    Set objDomNode = .selectSingleNode("DT_VENC_ATIV")
                    If Not objDomNode Is Nothing Then
                        If Not flValidaDataNumerica(objDomNode.Text) Then
                            'Data de vencimento inválida
                            fgAdicionaErro xmlErrosNegocio, 4063
                        End If
                    End If

                    For Each objDomRepeatConta In .selectSingleNode("REPE_CNTA").childNodes
                        With objDomRepeatConta
                            'Indicador débito/crédito
                            Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
                            If Not objDomNode Is Nothing Then
                                If Not Trim$(objDomNode.Text) = vbNullString Then
                                    lngTemp = Val(objDomNode.Text)
                                    If lngTemp <> enumTipoDebitoCredito.Credito And _
                                       lngTemp <> enumTipoDebitoCredito.Debito Then
                                        'Indicador de débito/crédito inválido
                                        fgAdicionaErro xmlErrosNegocio, 4021
                                    End If
                                End If
                            End If

                            'Quantidade do título
                            Set objDomNode = .selectSingleNode("QT_ATIV_MERC")
                            If Not objDomNode Is Nothing Then
                                If Val("0" & objDomNode.Text) = 0 Then
                                    'Quantidade de Títulos deve ser maior que zero
                                    fgAdicionaErro xmlErrosNegocio, 4050
                                End If
                            End If
                    
                            'Preço unitário
                            Set objDomNode = .selectSingleNode("PU_ATIV_MERC")
                            If Not objDomNode Is Nothing Then
                                If fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                                    'Preço Unitário deve ser maior que zero
                                    fgAdicionaErro xmlErrosNegocio, 4053
                                End If
                            End If
                        
                            'Valor Financeiro
                            Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
                            If Not objDomNode Is Nothing Then
                                If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                                    'Valor Financeiro deve ser maior que zero
                                    fgAdicionaErro xmlErrosNegocio, 4049
                                End If
                            End If
                        End With
                    Next objDomRepeatConta
                End With
            Next objDomRepeat
        End If
    End With

    fgConsisteLiquidacaoFisicaOperacao = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteLiquidacaoFisicaOperacao", 0

End Function

Public Function fgConsisteLiquidacaoEventos_BMA(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim lngCodigoEmpresaFusi                    As Long
Dim datDataVigencia                         As Date
Dim blnEntradaManual                        As Boolean
Dim lngTipoSolicitacao                      As Long
Dim lngTipoMensagem                         As Long
Dim lngTemp                                 As Long

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")
    
        lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)

        'Empresa
        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        End If
        lngCodigoEmpresaFusi = glngCodigoEmpresaFusionada

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
        
            If Not blnEntradaManual Then
                If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                               objDomNode.Text) Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If

        'Data da Operação
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da operação inválida
                fgAdicionaErro xmlErrosNegocio, 4024
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                    'Data da operação não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4059
                Else
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da operação inválida
                        fgAdicionaErro xmlErrosNegocio, 4024
                    End If
                End If
            End If
        End If
        
        'Data da mensagem
        Set objDomNode = .selectSingleNode("DT_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(.selectSingleNode("DT_MESG").Text)) Then
                    'Data da Mensagem inválida
                    fgAdicionaErro xmlErrosNegocio, 4060
                Else
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Mensagem inválida
                        fgAdicionaErro xmlErrosNegocio, 4060
                    End If
                
                End If
            End If
        End If

        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.BMA Then
                'Código do local de liquidação inválido
                fgAdicionaErro xmlErrosNegocio, 4008
            End If
        End If

        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            Else
                'Produto inválido
                fgAdicionaErro xmlErrosNegocio, 4011
            End If
        End If

        'Tipo de Conta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(lngCodigoEmpresaFusi, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If

        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(lngCodigoEmpresaFusi, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If
        
        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(lngCodigoEmpresaFusi, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If

        'Código do Índexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If

        'Tipo Pagamento LDL
        Set objDomNode = .selectSingleNode("TP_PAGTO_LDL")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumTipoPagamentoLDL.PagamentoJuros And _
               Val("0" & objDomNode.Text) <> enumTipoPagamentoLDL.Resgates And _
               Val("0" & objDomNode.Text) <> enumTipoPagamentoLDL.Amortizacoes Then
                'Tipo de Pagamento LDL
                fgAdicionaErro xmlErrosNegocio, 4093
            End If
        End If
        
        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                            objDomNode.Text, _
                                            datDataVigencia, _
                                            .selectSingleNode("CO_EMPR").Text) Then
                    'Veiculo legal inválido
                    fgAdicionaErro xmlErrosNegocio, 4007
                End If
            End If
        End If

        'Indicador débito/crédito
        Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                lngTemp = Val(objDomNode.Text)
                If lngTemp <> enumTipoDebitoCredito.Credito And _
                   lngTemp <> enumTipoDebitoCredito.Debito Then
                    'Indicador de débito/crédito inválido
                    fgAdicionaErro xmlErrosNegocio, 4021
                End If
            End If
        End If

        'Data de vencimento
        Set objDomNode = .selectSingleNode("DT_VENC_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data de vencimento inválida
                fgAdicionaErro xmlErrosNegocio, 4063
            End If
        End If
    
        'Quantidade do título
        Set objDomNode = .selectSingleNode("QT_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) = 0 Then
                'Quantidade de Títulos deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4050
            End If
        End If

        'Preço unitário
        Set objDomNode = .selectSingleNode("PU_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                'Preço Unitário deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4053
            End If
        End If
    
        'Valor Financeiro
        Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                'Valor Financeiro deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4049
            End If
        End If
    
        'Forma de Liquidação
        Set objDomNode = .selectSingleNode("CO_FORM_LIQU")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumFormaLiquidacao.ContaCorrente And _
               Val("0" & objDomNode.Text) <> enumFormaLiquidacao.Contabil Then
                'Forma de Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4026
            End If
        End If

        #If ValidaCC = 1 Then
        
            If Val("0" & .selectSingleNode("CO_FORM_LIQU").Text) = enumFormaLiquidacao.ContaCorrente And _
              (lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               lngTipoSolicitacao = enumTipoSolicitacao.Alteracao) Then
       
                'Código do banco
                Set objDomNode = .selectSingleNode("CO_BANC")
                If Not objDomNode Is Nothing Then
                    If Not fgExisteCodigoBanco(Val(objDomNode.Text)) Then
                        'Código do Banco inválido
                        fgAdicionaErro xmlErrosNegocio, 4027
                    End If
                End If
    
                'Código da Agência
                Set objDomNode = .selectSingleNode("CO_AGEN")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da Agência inválido
                        fgAdicionaErro xmlErrosNegocio, 4028
                    End If
                End If
    
                'Número da Conta Corrente
                Set objDomNode = .selectSingleNode("NU_CC")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4029
                    End If
                End If
    
                'Valor do Lançamento Conta Corrente
                Set objDomNode = .selectSingleNode("VA_LANC_CC")
                If Not objDomNode Is Nothing Then
                    If fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                        'Valor do lançamento na conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4030
                    End If
                End If
            End If
        #End If
    End With
    
    fgConsisteLiquidacaoEventos_BMA = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteLiquidacaoEventos_BMA", 0

End Function

Public Function fgConsisteIntermediacaoOperInterna_BMA(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim lngCodigoEmpresaFusi                    As Long
Dim datDataVigencia                         As Date
Dim blnEntradaManual                        As Boolean
Dim lngTipoSolicitacao                      As Long
Dim lngTipoMensagem                         As Long
Dim lngTemp                                 As Long
Dim lngQuantidadeTitulo                     As Long

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)
    
    With xmlDOMMensagem.selectSingleNode("/MESG")
    
        lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)

        'Empresa
        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        End If
        lngCodigoEmpresaFusi = glngCodigoEmpresaFusionada

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
        
            If Not blnEntradaManual Then
                If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                               objDomNode.Text) Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If

        'Data da Operação
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da operação inválida
                fgAdicionaErro xmlErrosNegocio, 4024
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                    'Data da operação não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4059
                Else
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da operação inválida
                        fgAdicionaErro xmlErrosNegocio, 4024
                    End If
                End If
            End If
        End If
        
        'Data da mensagem
        Set objDomNode = .selectSingleNode("DT_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(.selectSingleNode("DT_MESG").Text)) Then
                    'Data da Mensagem inválida
                    fgAdicionaErro xmlErrosNegocio, 4060
                Else
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Mensagem inválida
                        fgAdicionaErro xmlErrosNegocio, 4060
                    End If
                
                End If
            End If
        End If

        'Quantidade de Títulos
        If lngTipoMensagem <> enumTipoMensagemLQS.Redesconto And _
           lngTipoMensagem <> enumTipoMensagemLQS.Compromissada Then
            Set objDomNode = .selectSingleNode("QT_ATIV_MERC")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) <= 0 Then
                    'Quantidade de Títulos deve ser maior que zero
                    fgAdicionaErro xmlErrosNegocio, 4050
                Else
                    lngQuantidadeTitulo = objDomNode.Text
                End If
            End If
        End If

        'Valida Intermediação Operação Interna
        Set objDomNode = .selectSingleNode("NU_CTRL_MESG_SPB_ORIG")
        If Not objDomNode Is Nothing Then
            Call fgValidaMensagemEspecificacaoInterna(xmlErrosNegocio, objDomNode.Text, lngTipoSolicitacao, lngQuantidadeTitulo)
        End If
        
        
        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Not fgExisteLocalLiquidacao(lngCodigoEmpresaFusi, _
                                           objDomNode.Text, _
                                           datDataVigencia) Then
                'Código do local de liquidação inválido
                fgAdicionaErro xmlErrosNegocio, 4008
            Else
                If Val(objDomNode.Text) <> enumLocalLiquidacao.BMA Then
                    'Código do local de liquidação inválido
                    fgAdicionaErro xmlErrosNegocio, 4008
                End If
            End If
        End If

        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            Else
                'Produto inválido
                fgAdicionaErro xmlErrosNegocio, 4011
            End If
        End If

        'Tipo de Conta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(lngCodigoEmpresaFusi, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If

        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(lngCodigoEmpresaFusi, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If
        
        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(lngCodigoEmpresaFusi, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If

        'Código do Índexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If
        
        'Tipo de Negociação BMA
        Set objDomNode = .selectSingleNode("TP_NEGO_BMA")
        If Not objDomNode Is Nothing Then
            If Not fgExisteTipoNegocicaoBMA(objDomNode.Text) Then
                'Tipo de Negociação BMA inválido
                fgAdicionaErro xmlErrosNegocio, 4092
            End If
        End If

        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                            objDomNode.Text, _
                                            datDataVigencia, _
                                            .selectSingleNode("CO_EMPR").Text) Then
                    'Veiculo legal inválido
                    fgAdicionaErro xmlErrosNegocio, 4007
                End If
            End If
        End If

        'Indicador débito/crédito
        Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                lngTemp = Val(objDomNode.Text)
                If lngTemp <> enumTipoDebitoCredito.Credito And _
                   lngTemp <> enumTipoDebitoCredito.Debito Then
                    'Indicador de débito/crédito inválido
                    fgAdicionaErro xmlErrosNegocio, 4021
                End If
            End If
        End If

        'Data de vencimento
        Set objDomNode = .selectSingleNode("DT_VENC_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data de vencimento inválida
                fgAdicionaErro xmlErrosNegocio, 4063
            End If
        End If

        'Forma de Liquidação
        Set objDomNode = .selectSingleNode("CO_FORM_LIQU")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumFormaLiquidacao.ContaCorrente And _
               Val("0" & objDomNode.Text) <> enumFormaLiquidacao.Contabil Then
                'Forma de Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4026
            End If
        End If
        
        'Quantidade de Títulos
        If lngTipoMensagem <> enumTipoMensagemLQS.Redesconto And _
           lngTipoMensagem <> enumTipoMensagemLQS.Compromissada Then
            Set objDomNode = .selectSingleNode("QT_ATIV_MERC")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) <= 0 Then
                    'Quantidade de Títulos deve ser maior que zero
                    fgAdicionaErro xmlErrosNegocio, 4050
                End If
            End If
        End If

        'Preço unitário
        Set objDomNode = .selectSingleNode("PU_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal("0" & objDomNode.Text) <= 0 Then
                'Preço Unitário deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4053
            End If
        End If
        
        'Valor Financeiro
        Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                'Valor Financeiro deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4049
            End If
        End If
        
        '#If ValidaCC = 1 Then
            If .selectSingleNode("CO_FORM_LIQU").Text = enumFormaLiquidacao.ContaCorrente And _
              (lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               lngTipoSolicitacao = enumTipoSolicitacao.Alteracao) Then
        
                'Código do banco
                Set objDomNode = .selectSingleNode("CO_BANC")
                If Not objDomNode Is Nothing Then
                    If Not fgExisteCodigoBanco(Val(objDomNode.Text)) Then
                        'Código do Banco inválido
                        fgAdicionaErro xmlErrosNegocio, 4027
                    End If
                End If
    
                'Código da Agência
                Set objDomNode = .selectSingleNode("CO_AGEN")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da Agência inválido
                        fgAdicionaErro xmlErrosNegocio, 4028
                    End If
                End If
    
                'Número da Conta Corrente
                Set objDomNode = .selectSingleNode("NU_CC")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4029
                    End If
                End If
    
                'Valor do Lançamento Conta Corrente
                Set objDomNode = .selectSingleNode("VA_LANC_CC")
                If Not objDomNode Is Nothing Then
                    If fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                        'Valor do lançamento na conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4030
                    End If
                End If
            End If
        '#End If
        
    End With

    fgConsisteIntermediacaoOperInterna_BMA = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteIntemediacaoOperInterna_BMA", 0

End Function

Public Function fgConsisteRegistroOperacao_BMA(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim datDataVigencia                         As Date
Dim blnEntradaManual                        As Boolean
Dim lngCodigoEmpresaFusi                    As Long
Dim lngTipoMensagem                         As Long
Dim lngTipoNegociacao                       As Long
Dim lngTipoSolicitacao                      As Long
Dim lngTemp                                 As Long
Dim lngSubTipoNegociacao                    As Long
Dim lngTipoCobertura                        As Long
Dim strDataOperacaoRetorno                  As String

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")
    
    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)
    
    With xmlDOMMensagem.selectSingleNode("/MESG")
    
        If Not fgExisteTipoMensagem(.selectSingleNode("TP_MESG").Text, datDataVigencia) Then
            'Tipo de mensagem inválida
            fgAdicionaErro xmlErrosNegocio, 4003
        Else
            lngTipoMensagem = glngTipoMesg
        End If
    
        'Empresa
        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        End If
        lngCodigoEmpresaFusi = glngCodigoEmpresaFusionada
    
        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
        
            If Not blnEntradaManual Then
                If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                               objDomNode.Text) Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If
        
        'Tipo de Negociação BMA
        Set objDomNode = .selectSingleNode("TP_NEGO_BMA")
        If Not objDomNode Is Nothing Then
            
            If Not IsNumeric(objDomNode.Text) Then
                'Tipo de Negociação BMA inválido
                fgAdicionaErro xmlErrosNegocio, 4092
            Else
            
                If Not fgExisteTipoNegocicaoBMA(Val("0" & objDomNode.Text)) Then
                    'Tipo de Negociação BMA inválido
                    fgAdicionaErro xmlErrosNegocio, 4092
                Else
                    lngTipoNegociacao = Val("0" & objDomNode.Text)
                    objDomNode.Text = lngTipoNegociacao
                End If
            End If
        End If
        
        If (lngTipoNegociacao = enumTipoNegociacaoBMA.Compromissada Or _
           lngTipoNegociacao = enumTipoNegociacaoBMA.MigracaoIdaCamara Or _
           lngTipoNegociacao = enumTipoNegociacaoBMA.MigracaoVoltaCamara) And _
           lngTipoNegociacao <> 0 Then
            
            'Sub Tipo de Negociação BMA
            Set objDomNode = .selectSingleNode("CO_SUB_TIPO_NEGO")
            
            If Not objDomNode Is Nothing Then
                If Not fgExisteSubTipoNegocicaoBMA(Val("0" & objDomNode.Text)) Then
                    'Sub Tipo de Negociação BMA inválido
                    fgAdicionaErro xmlErrosNegocio, 4156
                Else
                    lngSubTipoNegociacao = Val("0" & objDomNode.Text)
                End If
            End If
            
        End If
                
        'Tipo de Cobertura
        Set objDomNode = .selectSingleNode("TP_COBE")
        If Not objDomNode Is Nothing Then
            If lngTipoNegociacao <> enumTipoNegociacaoBMA.TermoLeilao And _
               lngTipoNegociacao <> enumTipoNegociacaoBMA.TermoPapelDecorridoComCorrecao And _
               lngTipoNegociacao <> enumTipoNegociacaoBMA.TermoPapelDecorridoSemCorrecao Then
            
                If Val("0" & objDomNode.Text) <> enumCobertaDescobertaBMA.Coberta And _
                   Val("0" & objDomNode.Text) <> enumCobertaDescobertaBMA.Descoberta Then
                    'Tipo de Cobertura inválida
                    fgAdicionaErro xmlErrosNegocio, 4176
                Else
                    lngTipoCobertura = Val(objDomNode.Text)
                End If
            End If
        End If
                
        If lngTipoNegociacao = enumTipoNegociacaoBMA.DefinitivaD0 And _
           lngTipoCobertura = enumCobertaDescobertaBMA.Coberta Then
        
            'Conta Específica Deposito
            Set objDomNode = .selectSingleNode("CO_CNTA_CUTD_DEPO")
            
            If Not objDomNode Is Nothing Then
                If Val(objDomNode.Text) = 0 Then
                    'Conta Específica Depósito não informado
                    fgAdicionaErro xmlErrosNegocio, 4192
                End If
            Else
                'Conta Específica Depósito não informado
                fgAdicionaErro xmlErrosNegocio, 4192
            End If
            
        ElseIf lngTipoNegociacao = enumTipoNegociacaoBMA.Compromissada And _
               lngSubTipoNegociacao = enumSubTipoNegociacaoBMA.EspecificaAVista And _
               lngTipoCobertura = enumCobertaDescobertaBMA.Coberta Then
            
            'Conta Específica Deposito
            Set objDomNode = .selectSingleNode("CO_CNTA_CUTD_DEPO")
            
            If Not objDomNode Is Nothing Then
                If Val(objDomNode.Text) = 0 Then
                    'Conta Específica Depósito não informado
                    fgAdicionaErro xmlErrosNegocio, 4192
                End If
            Else
                'Conta Específica Depósito não informado
                fgAdicionaErro xmlErrosNegocio, 4192
            End If
            
        End If
                
        'Consiste dados do Titulo para regisdtro de operações BMA
        'de acordo com o Tipo de Negociação - enumTipoNegociacaoBMA
        If lngTipoNegociacao <> 0 Then
            Call fgConsisteTituloRegistroOperacao_BMA(xmlDOMMensagem, _
                                                      xmlErrosNegocio, _
                                                      lngTipoNegociacao, _
                                                      lngSubTipoNegociacao, _
                                                      lngTipoSolicitacao)
        End If
        
        'Número operação negociação BMA
        Set objDomNode = .selectSingleNode("NU_COMD_OPER")
        If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
            If Trim$(objDomNode.Text) = vbNullString Then
                'Número do comando da especifição obrigatório
                fgAdicionaErro xmlErrosNegocio, 4089
            End If
        End If

        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Not fgExisteLocalLiquidacao(lngCodigoEmpresaFusi, _
                                           objDomNode.Text, _
                                           datDataVigencia) Then
                'Código do local de liquidação inválido
                fgAdicionaErro xmlErrosNegocio, 4008
            Else
                If Val(objDomNode.Text) <> enumLocalLiquidacao.BMA Then
                    'Código do local de liquidação inválido
                    fgAdicionaErro xmlErrosNegocio, 4008
                End If
            End If
        End If
        
        'Preço unitário
        Set objDomNode = .selectSingleNode("PU_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                'Preço Unitário deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4053
            End If
        End If
        
        'Valor Financeiro Retorno
        Set objDomNode = .selectSingleNode("VA_OPER_ATIV_RETN")
        If Not objDomNode Is Nothing Then
            If lngTipoNegociacao = enumTipoNegociacaoBMA.Compromissada Or _
               lngTipoNegociacao = enumTipoNegociacaoBMA.MigracaoIdaCamara Or _
               lngTipoNegociacao = enumTipoNegociacaoBMA.MigracaoVoltaCamara Then
                
                If Trim$(objDomNode.Text) = vbNullString Then
                    'Valor Financeiro Retorno deve ser maior que zero
                    fgAdicionaErro xmlErrosNegocio, 4051
                ElseIf fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                    'Valor Financeiro Retorno deve ser maior que zero
                    fgAdicionaErro xmlErrosNegocio, 4051
                End If
            
            End If
        End If

        'Data Operação Retorno
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV_RETN")
        If Not objDomNode Is Nothing Then
            If lngTipoNegociacao = enumTipoNegociacaoBMA.Compromissada Or _
               lngTipoNegociacao = enumTipoNegociacaoBMA.MigracaoIdaCamara Or _
               lngTipoNegociacao = enumTipoNegociacaoBMA.MigracaoVoltaCamara Then
            
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data da Operação de Retorno inválida
                    fgAdicionaErro xmlErrosNegocio, 4064
                ElseIf Val("0" & objDomNode.Text) = 0 Then
                    'Data da Operação de Retorno inválida
                    fgAdicionaErro xmlErrosNegocio, 4064
                ElseIf fgDtXML_To_Date(objDomNode.Text) <= flDataHoraServidor(enumFormatoDataHora.Data) Then
                    'Data da operação de retorno deve ser maior que hoje
                    fgAdicionaErro xmlErrosNegocio, 4180
                End If
            End If
            strDataOperacaoRetorno = objDomNode.Text
        End If

        'Quantidade do título
        Set objDomNode = .selectSingleNode("QT_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) = 0 Then
                'Quantidade de Títulos deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4050
            End If
        End If
        
        'Hora de agendamento
        Set objDomNode = .selectSingleNode("HO_AGND")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 And _
            Not flValidaHoraNumerica(objDomNode.Text) Then
                'Horário de Agendamento inválido
                fgAdicionaErro xmlErrosNegocio, 4062
            End If
        End If
        
        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            Else
                If lngTipoNegociacao <> enumTipoNegociacaoBMA.MigracaoVoltaCamara Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            End If
        End If
    
        'Tipo de Conxta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(lngCodigoEmpresaFusi, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If
    
        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(lngCodigoEmpresaFusi, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If
        
        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(lngCodigoEmpresaFusi, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If

        'Código do indexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If

        'Data da Operação
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da operação inválida
                fgAdicionaErro xmlErrosNegocio, 4024
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da operação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4059
            Else
                If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                    'Data da operação inválida
                    fgAdicionaErro xmlErrosNegocio, 4024
                End If
            End If
        End If

        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                            objDomNode.Text, _
                                            datDataVigencia, _
                                            .selectSingleNode("CO_EMPR").Text) Then
                    'Veiculo legal inválido
                    fgAdicionaErro xmlErrosNegocio, 4007
                End If
            End If
        End If
        
        'Não permitir forma de liquidação = C/C quando o tipo de titular = PNA
        If gstrTipoTitularBMA = "PNA" Then
            Set objDomNode = .selectSingleNode("CO_FORM_LIQU")
            If Not objDomNode Is Nothing Then
                If IsNumeric(Trim$(objDomNode.Text)) Then
                    If CLng(objDomNode.Text) = enumFormaLiquidacao.ContaCorrente Then
                        'Operações do PNA não podem ser liquidadas em conta corrente
                        fgAdicionaErro xmlErrosNegocio, 4179
                    End If
                End If
            End If
        End If
        
        'Tipo de Titular
        Set objDomNode = .selectSingleNode("TP_TITL")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text <> vbNullString Then
                If Not fgExisteTipoTitular(objDomNode.Text) Then
                    'Tipo de titular inválido
                    fgAdicionaErro xmlErrosNegocio, 4094
                End If
            End If
        End If

        'Identificador do sistema de negociação
        Set objDomNode = .selectSingleNode("SG_SIST_NEGO")
        If Not objDomNode Is Nothing Then
            If Val(objDomNode.Text) <> 0 Then
                If Val(objDomNode.Text) <> enumSistemaNegociacaoBMA.NegociacaoEletronica And _
                   Val(objDomNode.Text) <> enumSistemaNegociacaoBMA.OperacoesBalcao Then
                    'Sistema de Negociação inválido
                    fgAdicionaErro xmlErrosNegocio, 4096
                End If
            End If
        End If
        
        'Indicador débito/crédito
        Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                lngTemp = Val(objDomNode.Text)
                If lngTemp <> enumTipoDebitoCredito.Credito And _
                   lngTemp <> enumTipoDebitoCredito.Debito Then
                    'Indicador de débito/crédito inválido
                    fgAdicionaErro xmlErrosNegocio, 4021
                End If
            End If
        End If

        'Tipo titular ofertante
        Set objDomNode = .selectSingleNode("TP_TITL_OFER")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text <> vbNullString Then
                If Not fgExisteTipoTitular(objDomNode.Text) Then
                    'Tipo de titular ofertante inválido
                    fgAdicionaErro xmlErrosNegocio, 4099
                End If
            End If
        End If
       
        'Data da Liquidação
        Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4068
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da Liquidação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4069
            ElseIf fgDtXML_To_Date(objDomNode.Text) <= flDataHoraServidor(enumFormatoDataHora.Data) And _
                (lngTipoNegociacao = enumTipoNegociacaoBMA.TermoLeilao Or _
                lngTipoNegociacao = enumTipoNegociacaoBMA.TermoPapelDecorridoComCorrecao Or _
                lngTipoNegociacao = enumTipoNegociacaoBMA.TermoPapelDecorridoSemCorrecao) Then
                'Data da liquidação deve ser maior que hoje
                fgAdicionaErro xmlErrosNegocio, 4070
            ElseIf fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) And _
                lngTipoNegociacao = enumTipoNegociacaoBMA.DefinitivaD0 Then
                'Data da Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4068
            ElseIf fgDtXML_To_Date(objDomNode.Text) <= flDataHoraServidor(enumFormatoDataHora.Data) And _
                lngSubTipoNegociacao = enumSubTipoNegociacaoBMA.GenericaATermo Then
                'Data da liquidação deve ser maior que hoje
                fgAdicionaErro xmlErrosNegocio, 4070
            ElseIf lngTipoNegociacao = enumTipoNegociacaoBMA.Compromissada Then
                If lngSubTipoNegociacao = enumSubTipoNegociacaoBMA.EspecificaAVista Then
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Liquidação inválida
                        fgAdicionaErro xmlErrosNegocio, 4068
                    ElseIf fgDtXML_To_Date(objDomNode.Text) >= fgDtXML_To_Date(strDataOperacaoRetorno) Then
                        'Data de Retorno deve ser maior que a Data da Liquidação
                        fgAdicionaErro xmlErrosNegocio, 4197
                    End If
                ElseIf lngSubTipoNegociacao = enumSubTipoNegociacaoBMA.EspecificaATermo Then
                    If fgDtXML_To_Date(objDomNode.Text) <= flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Liquidação inválida
                        fgAdicionaErro xmlErrosNegocio, 4068
                    ElseIf fgDtXML_To_Date(objDomNode.Text) >= fgDtXML_To_Date(strDataOperacaoRetorno) Then
                        'Data de Retorno deve ser maior que a Data da Liquidação
                        fgAdicionaErro xmlErrosNegocio, 4197
                    End If
                ElseIf lngSubTipoNegociacao = enumSubTipoNegociacaoBMA.GenericaAVista Then
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Liquidação inválida
                        fgAdicionaErro xmlErrosNegocio, 4068
                ElseIf fgDtXML_To_Date(objDomNode.Text) >= fgDtXML_To_Date(strDataOperacaoRetorno) Then
                        'Data de Retorno deve ser maior que a Data da Liquidação
                        fgAdicionaErro xmlErrosNegocio, 4197
                    End If
                ElseIf lngSubTipoNegociacao = enumSubTipoNegociacaoBMA.GenericaATermo Then
                    If fgDtXML_To_Date(objDomNode.Text) <= flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Liquidação inválida
                        fgAdicionaErro xmlErrosNegocio, 4068
                ElseIf fgDtXML_To_Date(objDomNode.Text) >= fgDtXML_To_Date(strDataOperacaoRetorno) Then
                        'Data de Retorno deve ser maior que a Data da Liquidação
                        fgAdicionaErro xmlErrosNegocio, 4197
                    End If
                End If
            End If
        End If

        'Data de Vencimento
        Set objDomNode = .selectSingleNode("DT_VENC_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data de vencimento inválida
                fgAdicionaErro xmlErrosNegocio, 4063
            Else
                If fgDtXML_To_Date(objDomNode.Text) < flDataHoraServidor(enumFormatoDataHora.Data) Then
                    'Data de vencimento inválida
                    fgAdicionaErro xmlErrosNegocio, 4063
                End If
            End If
        End If

        'Código da negociação
        Set objDomNode = .selectSingleNode("CO_NEGO")
        If Not objDomNode Is Nothing Then
            If Replace$(Trim$(objDomNode.Text), "0", vbNullString) = vbNullString Then
                'Código da negociação inválido
                fgAdicionaErro xmlErrosNegocio, 4098
            End If
        End If

        'Indicador de Participante Intermediário
        Set objDomNode = .selectSingleNode("IN_PARP_INTM")
        
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) <> vbNullString Then
                If Val("0" & objDomNode.Text) <> enumIndicadorSimNao.Sim And _
                   Val("0" & objDomNode.Text) <> enumIndicadorSimNao.Nao Then
                    'Indicador de Participante Intermediário
                    fgAdicionaErro xmlErrosNegocio, 4181
                End If
            End If
        End If

        'Finalidade de Cobertura de conta
        Set objDomNode = .selectSingleNode("CO_FIND_COBE_CNTA")
        If Not objDomNode Is Nothing Then
            If Val("0" & .selectSingleNode("TP_COBE").Text) = enumCobertaDescobertaBMA.Coberta Then
                If Not fgExisteFinalidadeCoberturaConta(Val("0" & objDomNode.Text)) Then
                    'Finalidade de Cobertura inválida
                    fgAdicionaErro xmlErrosNegocio, 4085
                End If
            End If
        End If

        'Tipo titular conta
        Set objDomNode = .selectSingleNode("TP_TITL_CNTA")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text <> vbNullString Then
                If Not fgExisteTipoTitular(objDomNode.Text) Then
                    'Tipo Titular conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4091
                End If
            End If
        End If

        'Valor Financeiro
        Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                'Valor Financeiro deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4049
            End If
        End If

        'Forma de Liquidação
        Set objDomNode = .selectSingleNode("CO_FORM_LIQU")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumFormaLiquidacao.ContaCorrente And _
               Val("0" & objDomNode.Text) <> enumFormaLiquidacao.Contabil Then
                'Forma de Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4026
            End If
        End If
        
        If Val("0" & .selectSingleNode("CO_FORM_LIQU").Text) = enumFormaLiquidacao.ContaCorrente And _
          (lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
           lngTipoSolicitacao = enumTipoSolicitacao.Alteracao) Then
    
            'Código do banco
            Set objDomNode = .selectSingleNode("CO_BANC")
            If Not objDomNode Is Nothing Then
                If Not fgExisteCodigoBanco(Val(objDomNode.Text)) Then
                    'Código do Banco inválido
                    fgAdicionaErro xmlErrosNegocio, 4027
                End If
            End If

            'Código da Agência
            Set objDomNode = .selectSingleNode("CO_AGEN")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) = 0 Then
                    'Código da Agência inválido
                    fgAdicionaErro xmlErrosNegocio, 4028
                End If
            End If

            'Número da Conta Corrente
            Set objDomNode = .selectSingleNode("NU_CC")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) = 0 Then
                    'Código da conta corrente inválido
                    fgAdicionaErro xmlErrosNegocio, 4029
                End If
            End If

            'Valor do Lançamento Conta Corrente
            Set objDomNode = .selectSingleNode("VA_LANC_CC")
            If Not objDomNode Is Nothing Then
                If fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                    'Valor do lançamento na conta corrente inválido
                    fgAdicionaErro xmlErrosNegocio, 4030
                End If
            End If
            
        End If

    End With

    fgConsisteRegistroOperacao_BMA = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteRegistroOperacao_BMA", 0


End Function

Public Function fgExisteFinalidadeCoberturaConta(ByVal plngFinalidadeCobertura As Long) As Boolean

Dim rsFinalidadeAberturaConta               As ADODB.Recordset

On Error GoTo ErrorHandler
    
    Set rsFinalidadeAberturaConta = fgQueryDominioInterno("CO_FIND_COBE_CNTA")
    
    With rsFinalidadeAberturaConta
    
        .Filter = " CO_DOMI = " & plngFinalidadeCobertura
    
        If .RecordCount > 0 Then
            glngFinalidadeCobertura = plngFinalidadeCobertura
            fgExisteFinalidadeCoberturaConta = True
        End If
    End With
    Set rsFinalidadeAberturaConta = Nothing
    
Exit Function
ErrorHandler:
    Set rsFinalidadeAberturaConta = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteFinalidadeCoberturaConta", 0

End Function

Public Function fgExisteTipoNegocicaoBMA(ByVal plngTipoNegociacaoBMA As Long) As Boolean

Dim rsTipoNegociacaoBMA                     As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsTipoNegociacaoBMA = fgQueryDominioInterno("TP_NEGO_BMA")
    
    With rsTipoNegociacaoBMA
    
        .Filter = " CO_DOMI = " & plngTipoNegociacaoBMA
    
        If .RecordCount > 0 Then
            plngTipoNegociacaoBMA = plngTipoNegociacaoBMA
            fgExisteTipoNegocicaoBMA = True
        End If
    End With
    Set rsTipoNegociacaoBMA = Nothing
Exit Function
ErrorHandler:
    Set rsTipoNegociacaoBMA = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoNegocicaoBMA", 0

End Function

Public Function fgExisteSubTipoNegocicaoBMA(ByVal plngSubTipoNegociacaoBMA As Long) As Boolean

Dim rsSubTipoNegociacaoBMA                  As ADODB.Recordset

On Error GoTo ErrorHandler
    
    If plngSubTipoNegociacaoBMA = 0 Then
        Exit Function
    End If
    
    Set rsSubTipoNegociacaoBMA = fgQueryDominioInterno("CO_SUB_TIPO_NEGO")
    
    With rsSubTipoNegociacaoBMA
    
        .Filter = " CO_DOMI = " & plngSubTipoNegociacaoBMA
    
        If .RecordCount > 0 Then
            plngSubTipoNegociacaoBMA = plngSubTipoNegociacaoBMA
            fgExisteSubTipoNegocicaoBMA = True
        End If
    End With
    Set rsSubTipoNegociacaoBMA = Nothing
Exit Function
ErrorHandler:
    Set rsSubTipoNegociacaoBMA = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteSubTipoNegocicaoBMA", 0

End Function

Public Function fgExisteTipoTitular(ByVal pstrTipoTitular As String) As Boolean

Dim rsTipoTitular                           As ADODB.Recordset

On Error GoTo ErrorHandler
    
    Set rsTipoTitular = fgQueryDominio("TP_TITL")
    
    With rsTipoTitular
    
        .Filter = " CO_DOMI = '" & pstrTipoTitular & "' "
    
        If .RecordCount > 0 Then
            gstrTipoTitular = pstrTipoTitular
            fgExisteTipoTitular = True
        End If
    End With
    Set rsTipoTitular = Nothing
Exit Function
ErrorHandler:
    Set rsTipoTitular = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoTitular", 0

End Function

'Concluir
Public Function fgExisteIndicadorFormaLiquRebate(ByVal pvntIndicadorFormaLiquidacaoRebate As Variant) As Boolean

Dim rsIndicadorFormaLiqu                    As ADODB.Recordset

On Error GoTo ErrorHandler
    
    Set rsIndicadorFormaLiqu = fgQueryDominio("CO_FORM_LIQU_REBA")
    
    With rsIndicadorFormaLiqu
    
        .Filter = " CO_DOMI = '" & pvntIndicadorFormaLiquidacaoRebate & "'"
    
        If .RecordCount > 0 Then
            gvntIndicadorFormaLiquRebate = pvntIndicadorFormaLiquidacaoRebate
            fgExisteIndicadorFormaLiquRebate = True
        End If
    End With
    Set rsIndicadorFormaLiqu = Nothing
Exit Function
ErrorHandler:
    Set rsIndicadorFormaLiqu = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteIndicadorFormaLiquRebate", 0

End Function

Public Function fgConsisteMensagemCETIP(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim objDomNodeRepet1                        As MSXML2.IXMLDOMNode
Dim objDomNodeRepet2                        As MSXML2.IXMLDOMNode

Dim datDataVigencia                         As Date
Dim lngCodigoEmpresaFusi                    As Long
Dim lngEmpresa                              As Long
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim lngTipoMensagem                         As Long
Dim lngTemp                                 As Long
Dim blnEntradaManual                        As Boolean
Dim lngTipoSolicitacao                      As Long
Dim lngTipoInstituicaoCreditadaDebitada     As Long
Dim lngCodigoOperacaoCETIP                  As Long
Dim lngTipoLiquidacao                       As Long
Dim lngTipoContraparte                      As Long
Dim lngDebitoCredito                        As Long
Dim strSiglaSistemaCETIP                    As String
Dim strSiglaSistemaOrigem                   As String
Dim lngTipoRentabilidade                    As Long
Dim strTipoTabelaResgate                    As String
Dim strDataOperacao                         As String
Dim strDataInicio                           As String
Dim strCodFormaPagtoCTP                     As String
Dim strCritCalcJuros                        As String
Dim strCodIndxCTP                           As String
Dim blnAtivosImobiliarios                   As Boolean
Dim strSubTipoAtivo                         As String
Dim strCodigoMensagemSPB                    As String
Dim blnValidarCC                            As Boolean

Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    With xmlDOMMensagem.selectSingleNode("/MESG")
    
        blnValidarCC = True
        datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)
    
        lngTipoMensagem = fgVlrXml_To_Decimal(.selectSingleNode("TP_MESG").Text)
        
        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        Else
            lngEmpresa = glngCodigoEmpresa
            lngCodigoEmpresaFusi = glngCodigoEmpresaFusionada
        End If
        
        If Not fgExisteSistema(.selectSingleNode("SG_SIST_ORIG").Text, glngCodigoEmpresa, datDataVigencia) Then
            'Sistema inexistente
            fgAdicionaErro xmlErrosNegocio, 4004
        Else
            strSiglaSistemaOrigem = .selectSingleNode("SG_SIST_ORIG").Text
        End If

        If Not .selectSingleNode("SG_SIST_DEST").Text = "A8" Then
            'Sistema destino incorreto
            fgAdicionaErro xmlErrosNegocio, 4005
        End If

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
        
            If Not blnEntradaManual Then
                If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                               objDomNode.Text) Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If

        'Data de envio
        Set objDomNode = .selectSingleNode("DT_REME")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            End If
        End If
        
        'Data da Mensagem Invalida
        Set objDomNode = .selectSingleNode("DT_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            ElseIf fgDtXML_To_Date(objDomNode.Text) >= fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 2, enumPaginacao.proximo) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            End If
        End If
        
        'Hora de envio
        Set objDomNode = .selectSingleNode("HO_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaHoraNumerica(fgCompletaString(objDomNode.Text, "0", 4, True)) Then
                'Hora da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4061
            End If
        End If

        'Hora de envio
        Set objDomNode = .selectSingleNode("HO_REME")
        If Not objDomNode Is Nothing Then
            If Not flValidaHoraNumerica(objDomNode.Text) Then
                'Hora da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4061
            End If
        End If

        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If lngTipoMensagem = enumTipoMensagemLQS.OperacoesComCorretorasCETIP Then
                If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.SSTR Then
                    'Código do local de liquidação inválido
                    fgAdicionaErro xmlErrosNegocio, 4008
                End If
            Else
                If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.CETIP Then
                    'Código do local de liquidação inválido
                    fgAdicionaErro xmlErrosNegocio, 4008
                End If
            End If
        End If

        'Código Operação CETIP
        Set objDomNode = .selectSingleNode("CO_OPER_CETIP")
        If Not objDomNode Is Nothing Then
            If Not fgExisteCodigoOperacaoCETIP(.selectSingleNode("CO_OPER_CETIP").Text, lngTipoMensagem, lngTipoSolicitacao) Then
                'Código de Operação CETIP inválido
                fgAdicionaErro xmlErrosNegocio, 4101
            Else
                lngCodigoOperacaoCETIP = Val(Trim$(objDomNode.Text))
            End If
        End If
        
        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If lngTipoMensagem = enumTipoMensagemLQS.OperacoesComCorretorasCETIP Then
                
                If strSiglaSistemaOrigem = "YS" Then
                    If Val("0" & objDomNode.Text) <> 186 And Val("0" & objDomNode.Text) <> 254 Then
                        'Produto inválido
                        fgAdicionaErro xmlErrosNegocio, 4011
                    End If
                End If
                If strSiglaSistemaOrigem = "LQC" Then
                    If Val("0" & objDomNode.Text) <> 470 And Val("0" & objDomNode.Text) <> 467 Then
                        'Produto inválido
                        fgAdicionaErro xmlErrosNegocio, 4011
                    End If
                End If
            
            ElseIf lngTipoMensagem <> enumTipoMensagemLQS.OperacaoRetencaoIRF_CETIP And _
                   lngTipoMensagem <> enumTipoMensagemLQS.RegistroContratoSWAP And _
                   lngTipoMensagem <> enumTipoMensagemLQS.RegistroContratoTermoCETIP And _
                   lngTipoMensagem <> enumTipoMensagemLQS.RegistroContratoSWAPCetip21 Then
                   
                If Val("0" & objDomNode.Text) <> 0 Then
                    If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                           objDomNode.Text, _
                                           datDataVigencia) Then
                        'Produto inválido
                        fgAdicionaErro xmlErrosNegocio, 4011
                    End If
                Else
                    If lngTipoMensagem <> enumTipoMensagemLQS.OperacaoRetencaoIRF_CETIP And _
                        lngTipoMensagem <> enumTipoMensagemLQS.RegistroContratoSWAP And _
                        lngTipoMensagem <> enumTipoMensagemLQS.RegistroContratoSWAPCetip21 Then
                        'Produto inválido
                        fgAdicionaErro xmlErrosNegocio, 4011
                    End If
                End If
                
            End If
        End If

        'Tipo Instituição Creditada / Debitada
        Set objDomNode = .selectSingleNode("TP_IF_CRED_DEB")
        If Not objDomNode Is Nothing Then
            If Val(objDomNode.Text) <> enumTipoInstituicaoCreditadaDebitada.CorretoraExterna And _
               Val(objDomNode.Text) <> enumTipoInstituicaoCreditadaDebitada.AgenteCompensacao And _
               Val(objDomNode.Text) <> enumTipoInstituicaoCreditadaDebitada.CorretoraInterna Then
                'Tipo Instituição Creditada / Debitada inválido
                fgAdicionaErro xmlErrosNegocio, 4161
            Else
                lngTipoInstituicaoCreditadaDebitada = Val(objDomNode.Text)
            End If
        End If

        'Código da Agência Creditada ou Debitada
        Set objDomNode = .selectSingleNode("CO_AGEN_CRED_DEB")
        If Not objDomNode Is Nothing Then
            If Val(objDomNode.Text) = 0 And lngTipoInstituicaoCreditadaDebitada = enumTipoInstituicaoCreditadaDebitada.CorretoraExterna Then
                'Código da Agência Creditada ou Debitada é obrigatório para esse Tipo Instituição Creditada ou Debitada
                fgAdicionaErro xmlErrosNegocio, 4162
            End If
        End If

        'Número da Conta Corrente Creditada ou Debitada
        Set objDomNode = .selectSingleNode("NU_CC_AGEN_CRED_DEB")
        If Not objDomNode Is Nothing Then
            If Val(objDomNode.Text) = 0 And lngTipoInstituicaoCreditadaDebitada = enumTipoInstituicaoCreditadaDebitada.CorretoraExterna Then
                'Número da Conta Corrente Creditada ou Debitada é obrigatório para esse Tipo Instituição Creditada ou Debitada
                fgAdicionaErro xmlErrosNegocio, 4163
            End If
        End If

        'Tipo Conta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If

        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(glngCodigoEmpresaFusionada, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If

        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(glngCodigoEmpresaFusionada, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If
        
        'Código do indexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If
        
        'Veículo Legal
        If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                    .selectSingleNode("CO_VEIC_LEGA").Text, _
                                    datDataVigencia, _
                                    .selectSingleNode("CO_EMPR").Text) Then
            'Veículo legal inexistente
            fgAdicionaErro xmlErrosNegocio, 4007
        End If
        
        'Modalidade Liquidacao
        lngTipoLiquidacao = 0
        Set objDomNode = .selectSingleNode("TP_LIQU_OPER_ATIV")
        If Not objDomNode Is Nothing Then

            If Trim$(.selectSingleNode("TP_LIQU_OPER_ATIV").Text) <> vbNullString Then
            
                If Not fgExisteTipoLiquidacao(.selectSingleNode("TP_LIQU_OPER_ATIV").Text, datDataVigencia) Then
                    'Tipo de Liquidação inválido
                    fgAdicionaErro xmlErrosNegocio, 4010
                Else
                    lngTipoLiquidacao = .selectSingleNode("TP_LIQU_OPER_ATIV").Text
                End If
    
                'Carlos - 30/08/2004 - Solicitação Pedro
                If lngTipoMensagem = enumTipoMensagemLQS.RegistroOperacaoesCETIP Or _
                   lngTipoMensagem = enumTipoMensagemLQS.AntecipacaoResgateContratoDerivativoCETIP Then
                   
                    If lngTipoLiquidacao <> enumTipoLiquidacao.Bruta And _
                       lngTipoLiquidacao <> enumTipoLiquidacao.SemModalidade Then
                        'Tipo de Liquidação inválido
                        fgAdicionaErro xmlErrosNegocio, 4010
                    End If
    
                ElseIf lngTipoMensagem = enumTipoMensagemLQS.EventoJurosCETIP Then
                    
                    If lngTipoLiquidacao <> enumTipoLiquidacao.Bruta And _
                       lngTipoLiquidacao <> enumTipoLiquidacao.SemModalidade And _
                       lngTipoLiquidacao <> enumTipoLiquidacao.Bilateral And _
                       lngTipoLiquidacao <> enumTipoLiquidacao.Multilateral Then
                        'Tipo de Liquidação inválido
                        fgAdicionaErro xmlErrosNegocio, 4010
                    End If
                    
                ElseIf lngTipoMensagem = enumTipoMensagemLQS.OperacaoDefinitivaCETIP Or _
                       lngTipoMensagem = enumTipoMensagemLQS.ResgateFundoInvestimentoCETIP Or _
                       lngTipoMensagem = enumTipoMensagemLQS.OperacaoCompromissadaCETIP Or _
                       lngTipoMensagem = enumTipoMensagemLQS.OperacaoRetornoAntecipacaoCETIP Then
                
                    If lngTipoLiquidacao <> enumTipoLiquidacao.Bruta And _
                       lngTipoLiquidacao <> enumTipoLiquidacao.SemModalidade And _
                       lngTipoLiquidacao <> enumTipoLiquidacao.Multilateral Then
                        'Tipo de Liquidação inválido
                        fgAdicionaErro xmlErrosNegocio, 4010
                    End If
                
                ElseIf lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesInstFinancCETIP Then
                
                    If .selectSingleNode("CO_SUB_TIPO_ATIV").Text = "CFA" And _
                        .selectSingleNode("SG_SIST_CETIP").Text = "SCF" And _
                        (lngCodigoOperacaoCETIP = 1 Or _
                         lngCodigoOperacaoCETIP = 101) Then
    
                        If lngTipoLiquidacao <> enumTipoLiquidacao.SemModalidade Then
                            'Tipo de Liquidação inválido
                            fgAdicionaErro xmlErrosNegocio, 4010
                        End If
    
                    Else
                        If lngTipoLiquidacao <> enumTipoLiquidacao.Bruta And _
                           lngTipoLiquidacao <> enumTipoLiquidacao.SemModalidade And _
                           lngTipoLiquidacao <> enumTipoLiquidacao.Multilateral Then
                            'Tipo de Liquidação inválido
                            fgAdicionaErro xmlErrosNegocio, 4010
                        End If
                    End If
                
                ElseIf lngTipoMensagem = enumTipoMensagemLQS.OperacaoRetencaoIRF_CETIP Then
    
                    If lngTipoLiquidacao <> enumTipoLiquidacao.SemModalidade And _
                       lngTipoLiquidacao <> enumTipoLiquidacao.Multilateral Then
                        'Tipo de Liquidação inválido
                        fgAdicionaErro xmlErrosNegocio, 4010
                    End If
    
                ElseIf lngTipoMensagem = enumTipoMensagemLQS.DespesasCETIP Then
    
                    If lngTipoLiquidacao <> enumTipoLiquidacao.Bruta Then
                        'Tipo de Liquidação inválido
                        fgAdicionaErro xmlErrosNegocio, 4010
                    End If
                    
                ElseIf lngTipoMensagem = enumTipoMensagemLQS.RegistroContratoTermoCETIP Then
                
                    If lngTipoLiquidacao <> enumTipoLiquidacao.SemModalidade _
                    And lngTipoLiquidacao <> enumTipoLiquidacao.Multilateral _
                    And lngTipoLiquidacao <> enumTipoLiquidacao.Bruta _
                    And lngTipoLiquidacao <> enumTipoLiquidacao.BilateralSTR _
                    And lngTipoLiquidacao <> enumTipoLiquidacao.BilateralBT _
                    And lngTipoLiquidacao <> enumTipoLiquidacao.BrutaSTR _
                    And lngTipoLiquidacao <> enumTipoLiquidacao.BrutaBT _
                    And lngTipoLiquidacao <> enumTipoLiquidacao.BrutaBTAUT _
                    And lngTipoLiquidacao <> enumTipoLiquidacao.Bilateral Then
                        'Tipo de Liquidação inválido
                        fgAdicionaErro xmlErrosNegocio, 4010
                    End If

                Else
    
                    If lngTipoLiquidacao <> enumTipoLiquidacao.SemModalidade Then
                        'Tipo de Liquidação inválido
                        fgAdicionaErro xmlErrosNegocio, 4010
                    End If
    
                End If
        
            End If
        
        End If

        'Tipo Indexador Termo CETIP
        Set objDomNode = .selectSingleNode("//TP_INDX_TERM_CETIP")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteTipoIndexadorTermoCetip(objDomNode.Text, objDomNode.nodeName) Then
                    'Tipo Indexador Termo CETIP inválido
                    fgAdicionaErro xmlErrosNegocio, 4213
                End If
            End If
        End If
        
        'Código indexador Termo Cetip
        Set objDomNode = .selectSingleNode("//CO_INDX_TERM_CETIP")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteCodigoIndexadorTermoCetip(objDomNode.Text, objDomNode.nodeName) Then
                    'Código indexador Termo Cetip Inválido
                    fgAdicionaErro xmlErrosNegocio, 4214
                End If
            End If
        End If

        'Indicador de Titular
        Set objDomNode = .selectSingleNode("IN_TITL")
        If Not objDomNode Is Nothing Then

'            'Carlos - 28/09/2004 - Solicitação Pedro
'            If lngTipoMensagem = enumTipoMensagemLQS.RegistroOperacaoesCETIP Then

'                If objDomNode.Text = "S" Then
'                    If lngDebitoCredito <> enumTipoDebitoCredito.Credito Then
'                        'Identificador do Titular Inválido.
'                        fgAdicionaErro xmlErrosNegocio, 4191
'                    End If
'                ElseIf objDomNode.Text = "N" Then
'                    If lngDebitoCredito <> enumTipoDebitoCredito.Debito Then
'                        'Identificador do Titular Inválido.
'                        fgAdicionaErro xmlErrosNegocio, 4191
'                    End If
'                Else
'                    'Identificador do Titular Inválido.
'                    fgAdicionaErro xmlErrosNegocio, 4191
'                End If
'            End If
        End If

        'Tipo Contraparte
        lngTipoContraparte = 0
        Set objDomNode = .selectSingleNode("TP_CNPT")
        
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumTipoContraparte.Cliente1 And _
               Val("0" & objDomNode.Text) <> enumTipoContraparte.Externo And _
               Val("0" & objDomNode.Text) <> enumTipoContraparte.Interno Then
                
                'Tipo Contraparte Inválido.
                fgAdicionaErro xmlErrosNegocio, 4182
            Else
                lngTipoContraparte = objDomNode.Text
                
                If lngTipoMensagem = enumTipoMensagemLQS.RegistroOperacaoesCETIP And _
                   lngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem And _
                   lngTipoContraparte <> enumTipoContraparte.Cliente1 Then
            
                    'Tipo Contraparte Inválido.
                    fgAdicionaErro xmlErrosNegocio, 4182
                End If
            End If
        End If

        'pikachu  - 02/04/2005
        'RATS 227
        If lngTipoContraparte = enumTipoContraparte.Cliente1 Then
            If lngTipoLiquidacao <> enumTipoLiquidacao.SemModalidade Then
                'Tipo de Liquidação inválido
                fgAdicionaErro xmlErrosNegocio, 4010
            End If
        End If
        
        'Data da operação
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            strDataOperacao = objDomNode.Text
        
            'Data da Operação não pode ser maior que hoje
            If fgDtXML_To_Date(objDomNode.Text) > flDataHoraServidor(enumFormatoDataHora.Data) And _
               lngTipoMensagem <> enumTipoMensagemLQS.EventoJurosCETIP Then
                'Data de movimento inválida
                fgAdicionaErro xmlErrosNegocio, 4012
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da operação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4059
            ElseIf fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                If lngTipoMensagem = enumTipoMensagemLQS.OperacaoRetencaoIRF_CETIP Or _
                   lngTipoMensagem = enumTipoMensagemLQS.RegistroContratoSWAP Or _
                   lngTipoMensagem = enumTipoMensagemLQS.RegistroOperacaoesCETIP Or _
                   lngTipoMensagem = enumTipoMensagemLQS.RegistroContratoTermoCETIP Or _
                   lngTipoMensagem = enumTipoMensagemLQS.AntecipacaoResgateContratoDerivativoCETIP Or _
                   lngTipoMensagem = enumTipoMensagemLQS.RegistroContratoSWAPCetip21 Then
                    
                    If lngTipoContraparte = enumTipoContraparte.Cliente1 Then
                        If lngTipoLiquidacao = enumTipoLiquidacao.SemModalidade Then
                            If fgDtXML_To_Date(objDomNode.Text) < fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 1, enumPaginacao.Anterior) Then
                                'Data de movimento inválida
                                fgAdicionaErro xmlErrosNegocio, 4012
                            End If
                        Else
                            'Data de movimento inválida
                            fgAdicionaErro xmlErrosNegocio, 4012
                        End If
                    Else
                    '    If fgDtXML_To_Date(objDomNode.Text) < fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 1, enumPaginacao.Anterior) Then
                            'Data de movimento inválida
                            fgAdicionaErro xmlErrosNegocio, 4012
                    '    End If
                    End If
                
                ElseIf lngTipoMensagem = enumTipoMensagemLQS.EventoJurosCETIP Then
                    If fgDtXML_To_Date(objDomNode.Text) > fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 1, enumPaginacao.proximo) Then
                        'Data de movimento inválida
                        fgAdicionaErro xmlErrosNegocio, 4012
                    End If
                
                ElseIf lngTipoMensagem = enumTipoMensagemLQS.OperacaoDefinitivaCETIP Then
                    If lngTipoContraparte = enumTipoContraparte.Cliente1 Then
                        If fgDtXML_To_Date(objDomNode.Text) < fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 1, enumPaginacao.Anterior) Then
                            'Data de movimento inválida
                            fgAdicionaErro xmlErrosNegocio, 4012
                        End If
                    Else
                        'Data de movimento inválida
                        fgAdicionaErro xmlErrosNegocio, 4012
                    End If
                    
                ElseIf lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesInstFinancCETIP Then
                    If lngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
                       lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                        If fgDtXML_To_Date(objDomNode.Text) < fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 1, enumPaginacao.Anterior) Then
                            'Data de movimento inválida
                            fgAdicionaErro xmlErrosNegocio, 4012
                        Else
                            If lngTipoContraparte <> enumTipoContraparte.Cliente1 Then
                                'Tipo Contraparte Inválido.
                                fgAdicionaErro xmlErrosNegocio, 4182
                            Else
                                Set objDomNode = .selectSingleNode("SG_SIST_CETIP")
                                If Not objDomNode Is Nothing Then
                                    If objDomNode.Text <> "CETIP" Then
                                        'Sigla de Sistema CETIP inválida para este tipo de mensagem
                                        fgAdicionaErro xmlErrosNegocio, 4196
                                    Else
                                        Set objDomNode = .selectSingleNode("TP_LIQU_OPER_ATIV")
                                        If Not objDomNode Is Nothing Then
                                            If Val("0" & objDomNode.Text) <> enumTipoLiquidacao.SemModalidade Then
                                                'Tipo de Liquidação inválido
                                                fgAdicionaErro xmlErrosNegocio, 4010
                                            End If
                                        Else
                                            'Tipo de Liquidação inválido
                                            fgAdicionaErro xmlErrosNegocio, 4010
                                        End If
                                    End If
                                Else
                                    'Sigla de Sistema CETIP inválida para este tipo de mensagem
                                    fgAdicionaErro xmlErrosNegocio, 4196
                                End If
                            End If
                        End If
                    Else
                        'Data de movimento inválida
                        fgAdicionaErro xmlErrosNegocio, 4012
                    End If
                ElseIf lngTipoMensagem = enumTipoMensagemLQS.OperacaoCompromissadaCETIP Or _
                       lngTipoMensagem = enumTipoMensagemLQS.OperacaoRetornoAntecipacaoCETIP Then
                    If fgDtXML_To_Date(objDomNode.Text) <= fgDtXML_To_Date("20060901") Then
                        'Data de movimento inválida
                        fgAdicionaErro xmlErrosNegocio, 4012
                    End If
                Else
                    'Data de movimento inválida
                    fgAdicionaErro xmlErrosNegocio, 4012
                End If
            End If
        End If

        'Data Vencimento Operação Original
        Set objDomNode = .selectSingleNode("//DT_VENC_OPER_ATIV_ORIG")
        If Not objDomNode Is Nothing Then
            If lngTipoMensagem = enumTipoMensagemBUS.OperacaoDefinitivaCETIP Then
                If lngCodigoOperacaoCETIP <> enumOperacaoCETIP_CTP0052.ResgateAntecipado Then
                    If Val(objDomNode.Text) <> 0 Then
                        If Not flValidaDataNumerica(Val(objDomNode.Text)) Then
                            'Data Vencimento Operação Original inválida
                            fgAdicionaErro xmlErrosNegocio, 4216
                        ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                            'Data Vencimento Operação Original inválida
                            fgAdicionaErro xmlErrosNegocio, 4216
                        End If
                    Else
                        'Data Vencimento Operação Original inválida
                        fgAdicionaErro xmlErrosNegocio, 4216
                    End If
                End If
            End If
        End If
        
        'ISPBIF Banco liquidante
        Set objDomNode = .selectSingleNode("CO_ISPB_BANC_LIQU_CNPT")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("CO_MODA_LIQU_FINC") Is Nothing Then
                If Val("0" & .selectSingleNode("CO_MODA_LIQU_FINC").Text) = enumTipoLiquidacao.BilateralSTR And _
                   Val("0" & objDomNode.Text) = 0 Then
                    'ISPBIF Banco liquidante não pode ser estar em branco para essa modalidade
                    fgAdicionaErro xmlErrosNegocio, 4150
                   End If
            End If
        End If
        
        'Identificador do Participante Câmara
        Set objDomNode = .selectSingleNode("CO_PARP_CAMR")
        If Not objDomNode Is Nothing Then
            If Not IsNumeric(objDomNode.Text) And Trim$(objDomNode.Text) <> vbNullString Then
                 'Identificador Participante Câmara inválido
                 fgAdicionaErro xmlErrosNegocio, 4189
            End If
        End If
        
        'Identificador de contraparte Câmara inválido
        Set objDomNode = .selectSingleNode("CO_CNPT_CAMR")
        If Not objDomNode Is Nothing Then
            If Not IsNumeric(objDomNode.Text) And Trim$(objDomNode.Text) <> vbNullString Then
                 'Identificador de contraparte Câmara inválido
                 fgAdicionaErro xmlErrosNegocio, 4139
            End If
        End If
        
        'Identificador de contraparte Câmara
        Set objDomNode = .selectSingleNode("CO_CNPT_CAMR")
        If Not objDomNode Is Nothing Then
            
            If lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesCustodiaCETIP Then
                
                If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0011.RetiradaCustodia Then
                    If Val("0" & Trim$(objDomNode.Text)) = 0 Then
                         'Identificador de contraparte Câmara inválido
                         fgAdicionaErro xmlErrosNegocio, 4139
                    End If
                End If
                
            ElseIf lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesContratoDerivativo Then
                
                If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9020.CessaoContrato Or _
                   lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9082.PagamentoIntermediacaoContrato Or _
                   lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9082.RegistroIntermediacaoContratoSwap Then
                     If Val("0" & Trim$(objDomNode.Text)) = 0 Then
                          'Identificador de contraparte Câmara inválido
                          fgAdicionaErro xmlErrosNegocio, 4139
                     End If
                End If
            End If
        End If
        
        'Número Operação Participante
        If lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesContratoDerivativo Then
            If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9020.CessaoContrato Or _
               lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9082.PagamentoIntermediacaoContrato Or _
               lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9082.RegistroIntermediacaoContratoSwap Then
                Set objDomNode = .selectSingleNode("NU_COMD_OPER")
                If Not objDomNode Is Nothing Then
                    If Val(objDomNode.Text) <= 0 Then
                        'Número Operação Participante é obrigatório
                        fgAdicionaErro xmlErrosNegocio, 4165
                    End If
                End If
            End If
        ElseIf lngTipoMensagem = enumTipoMensagemLQS.OperacaoDefinitivaCETIP Or _
               lngTipoMensagem = enumTipoMensagemLQS.OperacaoCompromissadaCETIP Then
            If lngTipoSolicitacao <> enumTipoSolicitacao.Inclusao And _
               lngTipoSolicitacao <> enumTipoSolicitacao.Cancelamento Then
                Set objDomNode = .selectSingleNode("NU_COMD_OPER")
                If Not objDomNode Is Nothing Then
                    If Val(objDomNode.Text) <= 0 Then
                        'Número Operação Participante é obrigatório
                        fgAdicionaErro xmlErrosNegocio, 4165
                    End If
                End If
            End If
        End If

        'Número Operação CTP Original
        If lngTipoMensagem = enumTipoMensagemLQS.OperacaoDefinitivaCETIP Or _
           lngTipoMensagem = enumTipoMensagemLQS.OperacaoRetencaoIRF_CETIP Or _
           lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesInstFinancCETIP Or _
           lngTipoMensagem = enumTipoMensagemLQS.ResgateFundoInvestimentoCETIP Or _
           lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesCustodiaCETIP Then
            'Validar somente se tipo solicitação for de cancelamento com mensagem
            If lngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then
                Set objDomNode = .selectSingleNode("NU_COMD_OPER_ORIG")
                If Not objDomNode Is Nothing Then
                    If Val(objDomNode.Text) <= 0 Then
                        'Número Operação Participante Original é obrigatório
                        fgAdicionaErro xmlErrosNegocio, 4195
                    End If
                End If
            End If
        End If

        If lngTipoMensagem = enumTipoMensagemLQS.EspecificacaoQuantidadesCotasCETIP Then
            Set objDomNode = .selectSingleNode("NU_COMD_OPER_ORIG")
            If Not objDomNode Is Nothing Then
                fgValidaMensagemEspecificacaoCETIP objDomNode.Text, _
                                                   .selectSingleNode("//CO_EMPR").Text, _
                                                   .selectSingleNode("//CO_VEIC_LEGA").Text, _
                                                   xmlErrosNegocio
            End If
        End If

        'Quantidade CETIP
        If lngTipoSolicitacao <> enumTipoSolicitacao.Inclusao And _
           lngTipoSolicitacao <> enumTipoSolicitacao.Cancelamento Then
            If lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesInstFinancCETIP Then
                Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
                'Verifica se a operação é Com Ordem de Lançamento
                If objTipoOperacao.flVerificarConciliacaoCETIP(xmlDOMMensagem, 0, strCodigoMensagemSPB) Then
                    If strCodigoMensagemSPB <> "CTP4001" Then
                        Set objDomNode = .selectSingleNode("QT_TITU_CETIP")
                        If Not objDomNode Is Nothing Then
                            If fgVlrXml_To_Decimal(objDomNode.Text) = 0 Then
                                'Quantidade CETIP inválida
                                fgAdicionaErro xmlErrosNegocio, 4166
                            End If
                        End If
                    End If
                End If
                Set objTipoOperacao = Nothing
            ElseIf lngTipoMensagem <> enumTipoMensagemLQS.EventoJurosCETIP And _
                   lngTipoMensagem <> enumTipoMensagemLQS.OperacaoCompromissadaCETIP Then
                Set objDomNode = .selectSingleNode("QT_TITU_CETIP")
                If Not objDomNode Is Nothing Then
                    If fgVlrXml_To_Decimal(objDomNode.Text) = 0 Then
                       'Quantidade CETIP inválida
                       fgAdicionaErro xmlErrosNegocio, 4166
                    End If
                End If
            End If
        End If

        'Sub Tipo Ativo
        strSubTipoAtivo = vbNullString
        
        Set objDomNode = .selectSingleNode("CO_SUB_TIPO_ATIV")
        If Not objDomNode Is Nothing Then
        
            strSubTipoAtivo = UCase$(Trim$(objDomNode.Text))
            
            'Pikachu - 09/03/2005
            'Alteração temporária para o SAC e SIGOM , pois estão com valor FIXO - QF
            'Aguardar SAC corrigir
            If strSubTipoAtivo = "QF" Or strSubTipoAtivo = "QFF" Then
                objDomNode.Text = "CFA"
                strSubTipoAtivo = "CFA"
            End If
            
            If strSubTipoAtivo = vbNullString Then
                If lngTipoMensagem <> enumTipoMensagemLQS.OperacaoRetencaoIRF_CETIP Then
                    If lngTipoMensagem = enumTipoMensagemLQS.OperacaoCompromissadaCETIP Then
                        If lngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Or _
                           lngTipoSolicitacao = enumTipoSolicitacao.CancelamentoPorLastro Or _
                           lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                            'SubTipo Ativo não pode ser nulo
                            fgAdicionaErro xmlErrosNegocio, 4108
                        End If
                    Else
                        'SubTipo Ativo não pode ser nulo
                        fgAdicionaErro xmlErrosNegocio, 4108
                    End If
                End If
            Else
                If lngTipoMensagem = enumTipoMensagemLQS.EspecificacaoQuantidadesCotasCETIP Then
                    If strSubTipoAtivo <> "CFA" Then
                        'SubTipo Ativo inválido
                        fgAdicionaErro xmlErrosNegocio, 4109
                    End If
                Else
                    If Not fgExisteSubTipoAtivo(strSubTipoAtivo) Then
                        'SubTipo Ativo inválido
                        fgAdicionaErro xmlErrosNegocio, 4109
                    End If
                End If
            End If
        End If

        'Valor financeiro
        Set objDomNode = .selectSingleNode("//VA_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                If lngTipoMensagem <> enumTipoMensagemLQS.ConversaoPermutaValorImobCETIP And _
                   lngTipoMensagem <> enumTipoMensagemLQS.AntecipacaoResgateContratoDerivativoCETIP And _
                   lngTipoMensagem <> enumTipoMensagemLQS.EventoJurosCETIP And _
                   lngTipoMensagem <> enumTipoMensagemLQS.OperacaoDefinitivaCETIP Then
                    'Valor Financeiro deve ser maior que zero
                    fgAdicionaErro xmlErrosNegocio, 4049
                'Ivan 31/01/2011 - Codigo removido, pois o VA_OPER_ATIV nao e obrigatorio para o Layout 64
'                Else
'                    'Validacao Layout 64 para novo Fluxo CTP0052
'                    'Ivan 25/05/2010
'                    If lngTipoMensagem = enumTipoMensagemLQS.OperacaoDefinitivaCETIP Then
'                        If strSubTipoAtivo = "DI" _
'                        Or strSubTipoAtivo = "DII" _
'                        Or strSubTipoAtivo = "DIM" _
'                        Or strSubTipoAtivo = "DIR" _
'                        Or strSubTipoAtivo = "DIRP" _
'                        Or strSubTipoAtivo = "DIRR" _
'                        Or strSubTipoAtivo = "DIRS" _
'                        Or strSubTipoAtivo = "DIRG" Then
'                            'Valor Financeiro deve ser maior que zero
'                            fgAdicionaErro xmlErrosNegocio, 4049
'                        Else
'                            If lngCodigoOperacaoCETIP <> 14 Then
'                                'Valor Financeiro deve ser maior que zero
'                                fgAdicionaErro xmlErrosNegocio, 4049
'                            End If
'                        End If
'                    End If
'                    'Fim
                End If
            End If
        End If

        'Indicador de direito caucionante
        Set objDomNode = .selectSingleNode("IN_DIRE_CAUC")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) <> vbNullString And _
               UCase$(Trim$(objDomNode.Text)) <> "S" And _
               UCase$(Trim$(objDomNode.Text)) <> "N" Then
                'Indicador de direito caucionante inválido
                fgAdicionaErro xmlErrosNegocio, 4102
            End If
        End If

        'Indicador de Corretora interna
        Set objDomNode = .selectSingleNode("IN_CCVM_INTE")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumIndicadorSimNao.Sim And _
               Val("0" & objDomNode.Text) <> enumIndicadorSimNao.Nao Then
                'Indicador corretora interna inválido
                fgAdicionaErro xmlErrosNegocio, 4103
            End If
        End If

        'Indicador MC interno
        Set objDomNode = .selectSingleNode("IN_MEMB_CPEN_INTE")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumIndicadorSimNao.Sim And _
               Val("0" & objDomNode.Text) <> enumIndicadorSimNao.Nao Then
                'Indicador MC interno inválido
                fgAdicionaErro xmlErrosNegocio, 4104
            End If
        End If
        
        'Tipo Remessa
        Set objDomNode = .selectSingleNode("TP_REME")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumTipoRemessaCETIP.Previa And _
               Val("0" & objDomNode.Text) <> enumTipoRemessaCETIP.Definitiva Then
                'Tipo Remessa inválido
                fgAdicionaErro xmlErrosNegocio, 4105
            End If
        End If
        
        'Valor Financeiro a liquidar
        Set objDomNode = .selectSingleNode("VA_OPER_LIQU")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal("0" & objDomNode.Text) <= 0 Then
                'Valor Financeiro a liquidar deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4106
            End If
        End If

        'Código Mensagem STR
        Set objDomNode = .selectSingleNode("CO_MESG_STR")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text <> strSTR0004 And _
               objDomNode.Text <> strSTR0007 Then
                'Código Mensagem STR inválido
                fgAdicionaErro xmlErrosNegocio, 4107
            End If
        End If
        
        'Sinal Taxa CETIP
        Set objDomNode = .selectSingleNode("TP_SINA_TAXA_CETIP")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) <> vbNullString And _
               objDomNode.Text <> "-" Then
                'Sinal Taxa CETIP inválido
                fgAdicionaErro xmlErrosNegocio, 4110
            End If
        End If
        
        'Data emissão CETIP
        Set objDomNode = .selectSingleNode("DT_EMIS_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text <> "0" And _
                objDomNode.Text <> vbNullString Then
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data emissão CETIP inválida
                    fgAdicionaErro xmlErrosNegocio, 4111
                End If
            End If
        End If
        
        'Data vencimento CETIP
        Set objDomNode = .selectSingleNode("DT_VENC_ATIV")
        If Not objDomNode Is Nothing Then
            If Val(objDomNode.Text) <> 0 Then
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data vencimento CETIP inválida
                    fgAdicionaErro xmlErrosNegocio, 4112
                End If
            End If
        End If
        
        'Data da aquisição
        Set objDomNode = .selectSingleNode("DT_AQUI_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If Val(objDomNode.Text) <> 0 Then
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data da aquisição inválida
                    fgAdicionaErro xmlErrosNegocio, 4113
                End If
            End If
        End If
        
        'Modalidade Liquidação
        Set objDomNode = .selectSingleNode("CO_MODA_LIQU_FINC")
        If Not objDomNode Is Nothing Then
            If Not fgExisteModalidadeLiquidacao(Val("0" & objDomNode.Text)) Then
                'Modalidade Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4114
            End If

            If lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesContratoDerivativo Then
                If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9020.CessaoContrato Or _
                   lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9082.PagamentoIntermediacaoContrato Or _
                   lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9082.RegistroIntermediacaoContratoSwap And _
                   Val(objDomNode.Text) = 0 Then
                    'Modalidade de liquidação obrigatória
                    fgAdicionaErro xmlErrosNegocio, 4167
                End If
            End If
            
        End If
        
        'Data de Agendamento
        Set objDomNode = .selectSingleNode("DT_AGND")
        If Not objDomNode Is Nothing Then
            If Trim$(Replace$(objDomNode.Text, "0", vbNullString)) <> vbNullString Then
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data de Agendamento inválida
                    fgAdicionaErro xmlErrosNegocio, 4115
                End If
            End If
        End If
        
        'Data de Retorno Compromisso
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV_RETN")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data de Retorno do Compromisso inválida
                fgAdicionaErro xmlErrosNegocio, 4168
            End If
        End If
        
        'Tipo de contrato SWAP
        If lngTipoMensagem = enumTipoMensagemLQS.RegistroContratoSWAP Or _
           lngTipoMensagem = enumTipoMensagemLQS.RegistroContratoSWAPCetip21 Then
            
            If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9001.RegistroContratoComModalidade Or _
               lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9001.RegistroRetroativoContratoComModalidade Then
            
                Set objDomNode = .selectSingleNode("TP_CNTR_SWAP")
                If Not objDomNode Is Nothing Then
                    If Not fgExisteTipoContratoSwap(Val("0" & objDomNode.Text)) Then
                        'Tipo de contrato SWAP
                        fgAdicionaErro xmlErrosNegocio, 4116
                    End If
                End If
            
            Else
            
                blnValidarCC = False
                
            End If
            
        End If
        
        'Código de moeda
        Set objDomNode = .selectSingleNode("CO_MOED")
        If Not objDomNode Is Nothing Then
            If Not fgExisteMoeda(Val("0" & objDomNode.Text)) Then
                'Código da moeda inválido
                fgAdicionaErro xmlErrosNegocio, 4117
            End If
        End If
        
        'Data da Liquidação
        If lngTipoMensagem = enumTipoMensagemLQS.OperacoesComCorretorasCETIP Then
            Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV")
            If Not objDomNode Is Nothing Then
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data da Liquidação inválida
                    fgAdicionaErro xmlErrosNegocio, 4068
                ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                    'Data da Liquidação não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4069
                ElseIf fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                    'Data da Liquidação inválida
                    fgAdicionaErro xmlErrosNegocio, 4068
                End If
            End If
        Else
            Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV")
            If Not objDomNode Is Nothing Then
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data da Liquidação inválida
                    fgAdicionaErro xmlErrosNegocio, 4068
                ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                    'Data da Liquidação não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4069
                ElseIf fgDtXML_To_Date(objDomNode.Text) <= flDataHoraServidor(enumFormatoDataHora.Data) Then
                    'Data da liquidação deve ser maior que hoje
                    fgAdicionaErro xmlErrosNegocio, 4070
                End If
            End If
        End If

        'Tipo de movimento financeiro
        Set objDomNode = .selectSingleNode("TP_MOVI_FINC")
        If Not objDomNode Is Nothing Then
            If Not fgExisteTipoMovimento(Val("0" & objDomNode.Text)) Then
                'Tipo de movimento financeiro inválido
                fgAdicionaErro xmlErrosNegocio, 4118
            End If
        End If
        
        'Finalidade IF
        Set objDomNode = .selectSingleNode("CO_FIND_IF")
        If Not objDomNode Is Nothing Then
            If Not fgExisteFinalidadeIF(Val("0" & objDomNode.Text)) Then
                'Finalidade IF inválida
                fgAdicionaErro xmlErrosNegocio, 4119
            End If
        End If
        
        'Tipo de fonte
        Set objDomNode = .selectSingleNode("TP_FONT")
        If Not objDomNode Is Nothing Then
            If Not fgExisteTipoFonte(Val("0" & objDomNode.Text)) Then
                'Tipo de fonte inválido
                fgAdicionaErro xmlErrosNegocio, 4120
            End If
        End If
        
        'Tipo de Cedente ou Adquirente
        Set objDomNode = .selectSingleNode("TP_CEDE_ADQU")
        If Not objDomNode Is Nothing Then
            If lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesContratoDerivativo And _
                lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9020.CessaoContrato And _
                Trim$(objDomNode.Text) = vbNullString Then
                'Tipo de Cedente ou Adquirente inválido
                fgAdicionaErro xmlErrosNegocio, 4121
            ElseIf objDomNode.Text <> "A" And _
                   objDomNode.Text <> "C" And _
                   Trim$(objDomNode.Text) <> vbNullString Then
                'Tipo de Cedente ou Adquirente inválido
                fgAdicionaErro xmlErrosNegocio, 4121
            End If
        End If
        
        'Identificador Anuente Câmara
        Set objDomNode = .selectSingleNode("CO_ANUE_CAMR")
        If Not objDomNode Is Nothing Then
            If lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesContratoDerivativo And _
                lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9020.CessaoContrato And _
                Val(objDomNode.Text) = 0 Then
                'Identificador Anuente Câmara inválido
                fgAdicionaErro xmlErrosNegocio, 4169
            End If
        End If
        
        'Indicador de boletim
        Set objDomNode = .selectSingleNode("IN_BOLE")
        If Not objDomNode Is Nothing Then
            If Not fgExisteIndicadorBoletim(Val("0" & objDomNode.Text)) Then
                'Indicador de boletim inválido
                fgAdicionaErro xmlErrosNegocio, 4122
            End If
        End If
        
        'Indicador Exercício Opção
        Set objDomNode = .selectSingleNode("IN_EXEC_OPCA")
        If Not objDomNode Is Nothing Then
            If Not fgExisteIndicadorExecOpcao(objDomNode.Text, objDomNode.nodeName) Then
                'Indicador Exercício Opção Inválido
                fgAdicionaErro xmlErrosNegocio, 4185
            End If
        End If

        'Indicador de Cross Rate
        Set objDomNode = .selectSingleNode("IN_CROS_RATE")
        If Not objDomNode Is Nothing Then
            If UCase$(objDomNode.Text) <> "S" And _
               UCase$(objDomNode.Text) <> "N" Then
                'Indicador de Cross Rate inválido
                fgAdicionaErro xmlErrosNegocio, 4125
            End If
        End If
        
        'Taxa de intermediação
        If lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesContratoDerivativo Then
            Set objDomNode = .selectSingleNode("PE_TAXA_INTM")
            If Not objDomNode Is Nothing Then
                If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9082.PagamentoIntermediacaoContrato Or _
                   lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9082.PagamentoIntermediacaoContrato Then
                   If fgVlrXml_To_Decimal(objDomNode.Text) = 0 Then
                    'Taxa de intermediação obrigatória
                    fgAdicionaErro xmlErrosNegocio, 4170
                   End If
                End If
            End If
        End If
        
        'Indicador de deslocamento
        Set objDomNode = .selectSingleNode("IN_DESL")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumIndicadorDeslocamento.D0 And _
               Val("0" & objDomNode.Text) <> enumIndicadorDeslocamento.DM1 And _
               Val("0" & objDomNode.Text) <> enumIndicadorDeslocamento.DM2 Then
                'Indicador de deslocamento inválido
                fgAdicionaErro xmlErrosNegocio, 4126
            End If
        End If

        'Indicador Participante/Contraparte
        Set objDomNode = .selectSingleNode("IN_PARP_CNPT")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text <> "C" And _
               objDomNode.Text <> "P" Then
                'Indicador Participante/Contraparte inválido
                fgAdicionaErro xmlErrosNegocio, 4171
            End If
        End If

        'Indicador de débito/crédito
        Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                lngDebitoCredito = Val(objDomNode.Text)
                If lngDebitoCredito <> enumTipoDebitoCredito.Credito And _
                   lngDebitoCredito <> enumTipoDebitoCredito.Debito Then
                    'Indicador de débito/crédito inválido
                    fgAdicionaErro xmlErrosNegocio, 4021
                End If
            End If
        End If

        'Sigla Sistema CETIP
        Set objDomNode = .selectSingleNode("SG_SIST_CETIP")
        If Not objDomNode Is Nothing Then
            If lngTipoSolicitacao <> enumTipoSolicitacao.Inclusao And _
               lngTipoSolicitacao <> enumTipoSolicitacao.Cancelamento Then
                If Not fgExisteSiglaSistemaCETIP(lngTipoMensagem, objDomNode.Text) Then
                    'Sigla de Sistema CETIP inválida para este tipo de mensagem
                    fgAdicionaErro xmlErrosNegocio, 4196
                Else
                    strSiglaSistemaCETIP = Trim$(objDomNode.Text)
                End If
            End If
        End If

        'Tipo de tabela de resgate
        Set objDomNode = .selectSingleNode("TP_TABE_RESG")
        If Not objDomNode Is Nothing Then
            
            strTipoTabelaResgate = Trim$(objDomNode.Text)
            
            If strSiglaSistemaCETIP = "CETIP" Then
                If strTipoTabelaResgate <> vbNullString Then
                    If strTipoTabelaResgate <> "M" And _
                       strTipoTabelaResgate <> "S" Then
                        'Tipo de tabela de resgate inválido
                        fgAdicionaErro xmlErrosNegocio, 4239
                    End If
                Else
                    'Tipo de tabela de resgate inválido
                    fgAdicionaErro xmlErrosNegocio, 4239
                End If
            Else
                If strTipoTabelaResgate <> vbNullString Then
                    If strTipoTabelaResgate <> "C" And _
                       strTipoTabelaResgate <> "M" And _
                       strTipoTabelaResgate <> "S" And _
                       strTipoTabelaResgate <> "N" Then
                        'Tipo de tabela de resgate inválido
                        fgAdicionaErro xmlErrosNegocio, 4239
                    End If
                End If
            End If
        End If

        'Validacao Layout 52 para novo Fluxo CTP4001
        'Ivan 05/05/2010
        If lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesInstFinancCETIP Then
            If strSubTipoAtivo = "DI" _
            Or strSubTipoAtivo = "DII" _
            Or strSubTipoAtivo = "DIM" _
            Or strSubTipoAtivo = "DIR" _
            Or strSubTipoAtivo = "DIRP" _
            Or strSubTipoAtivo = "DIRR" _
            Or strSubTipoAtivo = "DIRS" _
            Or strSubTipoAtivo = "DIRG" Then

                'Validacao fase complementacao
                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    
                    'Valida Sigla Sistema CETIP
                    If strSiglaSistemaCETIP = "CETIP21" Then
                
                        'Credito
                        If lngDebitoCredito = enumTipoDebitoCredito.Credito Then
                            'Valida Identificador Titulo CETIP
                            Set objDomNode = .selectSingleNode("NU_ATIV_MERC_CETIP")
                            If Not objDomNode Is Nothing Then
                                If Trim$(objDomNode.Text) = vbNullString Then
                                    'Identificador de Título inválido
                                    fgAdicionaErro xmlErrosNegocio, 4074
                                End If
                            End If
                        End If
                        
                        'Debito
                        If lngDebitoCredito = enumTipoDebitoCredito.Debito Then
                            
                            'Codigo Forma Pagamento CETIP
                            Set objDomNode = .selectSingleNode("CD_FORM_PAGTO")
                            If Not objDomNode Is Nothing Then
                                strCodFormaPagtoCTP = Trim$(objDomNode.Text)
                                If Trim$(objDomNode.Text) <> vbNullString Then
                                    If strCodFormaPagtoCTP <> "01" And _
                                       strCodFormaPagtoCTP <> "02" And _
                                       strCodFormaPagtoCTP <> "03" And _
                                       strCodFormaPagtoCTP <> "04" And _
                                       strCodFormaPagtoCTP <> "06" And _
                                       strCodFormaPagtoCTP <> "07" And _
                                       strCodFormaPagtoCTP <> "12" Then
                                        'Codigo Forma Pagamento CETIP invalido
                                        fgAdicionaErro xmlErrosNegocio, 4442
                                    End If
                                Else
                                    'Codigo Forma Pagamento CETIP obrigatorio para CETIP21
                                    fgAdicionaErro xmlErrosNegocio, 4441
                                End If
                            Else
                                'Codigo Forma Pagamento CETIP obrigatorio para CETIP21
                                fgAdicionaErro xmlErrosNegocio, 4441
                            End If
                            
                            'Codigo Forma Pagamento CETIP = Pre
                            If strCodFormaPagtoCTP = "12" Then
                            
                                'Valor resgate
                                Set objDomNode = .selectSingleNode("VA_FINC_BASE_RESG")
                                If Not objDomNode Is Nothing Then
                                    If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                                        'Valor Resgate é obrigatório
                                        fgAdicionaErro xmlErrosNegocio, 4199
                                    End If
                                Else
                                    'Valor Resgate é obrigatório
                                    fgAdicionaErro xmlErrosNegocio, 4199
                                End If
                                
                            'Codigo Forma Pagamento CETIP = Pos
                            Else
                            
                                'Codigo Indexador CETIP
                                Set objDomNode = .selectSingleNode("CO_INDX_CETIP")
                                If Not objDomNode Is Nothing Then
                                    strCodIndxCTP = Trim$(objDomNode.Text)
                                    If strCodIndxCTP <> vbNullString Then
                                        If strCodIndxCTP <> "1" And _
                                           strCodIndxCTP <> "3" And _
                                           strCodIndxCTP <> "9" And _
                                           strCodIndxCTP <> "10" And _
                                           strCodIndxCTP <> "11" And _
                                           strCodIndxCTP <> "16" And _
                                           strCodIndxCTP <> "18" And _
                                           strCodIndxCTP <> "20" And _
                                           strCodIndxCTP <> "23" And _
                                           strCodIndxCTP <> "99" Then
                                            'Codigo Indexador CETIP invalido
                                            fgAdicionaErro xmlErrosNegocio, 4436
                                        End If
                                    Else
                                        'Codigo Indexador CETIP obrigatorio para CETIP21
                                        fgAdicionaErro xmlErrosNegocio, 4435
                                    End If
                                Else
                                    'Codigo Indexador CETIP obrigatorio para CETIP21
                                    fgAdicionaErro xmlErrosNegocio, 4435
                                End If
                                
                                'Valor Nominal
                                Set objDomNode = .selectSingleNode("VA_NOML")
                                If Not objDomNode Is Nothing Then
                                    If Trim$(objDomNode.Text) = vbNullString Then
                                        'Valor Nominal obrigatorio para CETIP21
                                        fgAdicionaErro xmlErrosNegocio, 4447
                                    End If
                                End If
                                
                                'Taxa Flutuante Juros CETIP
                                Set objDomNode = .selectSingleNode("PE_FLUT_JURO_CETIP")
                                If Not objDomNode Is Nothing Then
                                    If Trim$(objDomNode.Text) = vbNullString Then
                                        'Taxa Flutuante Juros CETIP obrigatoria para CETIP21 Pos
                                        fgAdicionaErro xmlErrosNegocio, 4440
                                    End If
                                End If
                                
                                'Valor Preco Unitario Negociacao
                                Set objDomNode = .selectSingleNode("VA_PU_NEGO")
                                If Not objDomNode Is Nothing Then
                                    If Trim$(objDomNode.Text) = vbNullString Then
                                        'Valor Deposito obrigatorio para CETIP21
                                        fgAdicionaErro xmlErrosNegocio, 4443
                                    End If
                                End If
                                
                                'Taxa Juros CETIP
                                Set objDomNode = .selectSingleNode("PE_TAXA_CETIP")
                                If Not objDomNode Is Nothing Then
                                    If Trim$(objDomNode.Text) <> vbNullString _
                                    And Trim$(objDomNode.Text) <> "0" Then
                                        'Criterio Calculo Juros
                                        Set objDomNode = .selectSingleNode("CO_CRIT_CALC_JUROS")
                                        If Not objDomNode Is Nothing Then
                                            strCritCalcJuros = Trim$(objDomNode.Text)
                                            If strCritCalcJuros <> vbNullString Then
                                                If strCritCalcJuros <> "1" And _
                                                   strCritCalcJuros <> "2" And _
                                                   strCritCalcJuros <> "3" And _
                                                   strCritCalcJuros <> "4" And _
                                                   strCritCalcJuros <> "5" And _
                                                   strCritCalcJuros <> "6" Then
                                                    'Criterio Calculo Juros invalido
                                                    fgAdicionaErro xmlErrosNegocio, 4439
                                                End If
                                            Else
                                                'Criterio Calculo Juros obrigatorio para CETIP21 Pre
                                                fgAdicionaErro xmlErrosNegocio, 4438
                                            End If
                                        Else
                                            'Criterio Calculo Juros obrigatorio para CETIP21 Pre
                                            fgAdicionaErro xmlErrosNegocio, 4438
                                        End If
                                    End If
                                End If
                            
                            End If
                            
                            'Tipo de tabela de resgate
                            If strTipoTabelaResgate = vbNullString Then
                                'Tipo de tabela de resgate obrigatorio para CETIP21 Pre
                                fgAdicionaErro xmlErrosNegocio, 4437
                            End If
                            
                            'Indicador de Alteracao
                            Set objDomNode = .selectSingleNode("IN_ALTE")
                            If Not objDomNode Is Nothing Then
                                If Trim$(objDomNode.Text) <> vbNullString Then
                                    If Trim$(objDomNode.Text) <> "A" And _
                                       Trim$(objDomNode.Text) <> "E" And _
                                       Trim$(objDomNode.Text) <> "I" Then
                                        'Indicador de Alteracao invalido
                                        fgAdicionaErro xmlErrosNegocio, 4445
                                    End If
                                Else
                                    'Indicador de Alteracao obrigatorio para CETIP21
                                    fgAdicionaErro xmlErrosNegocio, 4444
                                End If
                            Else
                                'Indicador de Alteracao obrigatorio para CETIP21
                                fgAdicionaErro xmlErrosNegocio, 4444
                            End If
    
                            'Prazo Instrumento Financeiro
                            Set objDomNode = .selectSingleNode("PZ_INST_FINC")
                            If Not objDomNode Is Nothing Then
                                If Trim$(objDomNode.Text) = vbNullString Then
                                    'Prazo Instrumento Financeiro obrigatorio para CETIP21
                                    fgAdicionaErro xmlErrosNegocio, 4446
                                End If
                            End If
                            
                        End If
                        
                        'Valida Repeticao Distribuicao
                        If Not .selectSingleNode("REPE_DTBC") Is Nothing Then
                            
                            For Each objDomNodeRepet1 In .selectNodes("REPE_DTBC/*")
                                
                                'Data Inicio Distribuicao Publica
                                Set objDomNode = objDomNodeRepet1.selectSingleNode("DT_INIC_DTBC_PUBL")
                                If Not objDomNode Is Nothing Then
                                    If Not flValidaDataNumerica(objDomNode.Text) Then
                                        'Data Inicio Distribuicao Publica invalida
                                        fgAdicionaErro xmlErrosNegocio, 4448
                                    End If
                                Else
                                    'Data Inicio Distribuicao Publica obrigatoria para CETIP21
                                    fgAdicionaErro xmlErrosNegocio, 4449
                                End If
                                
                                'Data Fim Distribuicao Publica
                                Set objDomNode = objDomNodeRepet1.selectSingleNode("DT_FIM_DTBC_PUBL")
                                If Not objDomNode Is Nothing Then
                                    If Not flValidaDataNumerica(objDomNode.Text) Then
                                        'Data Fim Distribuicao Publica invalida
                                        fgAdicionaErro xmlErrosNegocio, 4450
                                    End If
                                Else
                                    'Data Fim Distribuicao Publica obrigatoria para CETIP21
                                    fgAdicionaErro xmlErrosNegocio, 4451
                                End If
                                
                                'Identificador Participante Coordenador Lider
                                Set objDomNode = objDomNodeRepet1.selectSingleNode("ID_PARP_COOR_LIDE")
                                If Not objDomNode Is Nothing Then
                                    If Trim$(objDomNode.Text) = vbNullString Then
                                        'Identificador Participante Coordenador Lider obrigatorio para CETIP21
                                        fgAdicionaErro xmlErrosNegocio, 4452
                                    End If
                                Else
                                    'Identificador Participante Coordenador Lider obrigatorio para CETIP21
                                    fgAdicionaErro xmlErrosNegocio, 4452
                                End If
                                
                                'Indicador Esforco Restrito
                                Set objDomNode = objDomNodeRepet1.selectSingleNode("IN_ESFO_RSTT")
                                If Not objDomNode Is Nothing Then
                                    If Trim$(objDomNode.Text) = vbNullString Then
                                        'Indicador Esforco Restrito obrigatorio para CETIP21
                                        fgAdicionaErro xmlErrosNegocio, 4453
                                    Else
                                        If UCase(Trim$(objDomNode.Text)) <> "S" _
                                        And UCase(Trim$(objDomNode.Text)) <> "N" Then
                                            'Indicador Esforco Restrito invalido para CETIP21
                                            fgAdicionaErro xmlErrosNegocio, 4457
                                        End If
                                    End If
                                Else
                                    'Indicador Esforco Restrito obrigatorio para CETIP21
                                    fgAdicionaErro xmlErrosNegocio, 4453
                                End If
                                                               
                                'Repeticao Classificacao
                                If Not objDomNodeRepet1.selectSingleNode("REPE_CSFCC") Is Nothing Then
                                
                                    For Each objDomNodeRepet2 In objDomNodeRepet1.selectSingleNode("REPE_CSFCC").childNodes
                                        
                                        'Nome Classificadora Risco
                                        Set objDomNode = objDomNodeRepet2.selectSingleNode("NO_CSFCD_RISC")
                                        If Not objDomNode Is Nothing Then
                                            If Trim$(objDomNode.Text) = vbNullString Then
                                                'Nome Classificadora Risco obrigatorio para CETIP21
                                                fgAdicionaErro xmlErrosNegocio, 4454
                                            End If
                                        End If

                                        'Codigo Classificacao Risco
                                        Set objDomNode = objDomNodeRepet2.selectSingleNode("CD_CSFCC_RISC")
                                        If Not objDomNode Is Nothing Then
                                            If Trim$(objDomNode.Text) = vbNullString Then
                                                'Codigo Classificacao Risco obrigatorio para CETIP21
                                                fgAdicionaErro xmlErrosNegocio, 4454
                                            End If
                                        End If
                                    
                                    Next objDomNodeRepet2
                                End If
                                    
                            Next objDomNodeRepet1
                        End If
                    Else
                        'Sigla de Sistema CETIP inválida para este tipo de mensagem
                        fgAdicionaErro xmlErrosNegocio, 4196
                    End If
                End If
            End If
        End If
        'Fim
        
        'Indicador Distribuicao Publica
        Set objDomNode = .selectSingleNode("IN_DTBC_PUBL")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) <> vbNullString Then
                If UCase(Trim$(objDomNode.Text)) <> "S" _
                And UCase(Trim$(objDomNode.Text)) <> "N" Then
                    'Indicador Distribuicao Publica invalido
                    fgAdicionaErro xmlErrosNegocio, 4458
                End If
            End If
        End If

        #If ValidaCC = 1 Then
        
            If blnValidarCC Then
            
                'Forma de Liquidação
                Set objDomNode = .selectSingleNode("CO_FORM_LIQU")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) <> enumFormaLiquidacao.ContaCorrente And _
                       Val("0" & objDomNode.Text) <> enumFormaLiquidacao.Contabil And _
                       Val("0" & objDomNode.Text) <> enumFormaLiquidacao.ContaCorrenteComTributacao Then
                        'Forma de Liquidação inválida
                        fgAdicionaErro xmlErrosNegocio, 4026
                    End If
            
                    If .selectSingleNode("CO_FORM_LIQU").Text <> enumFormaLiquidacao.Contabil And _
                      (lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                       lngTipoSolicitacao = enumTipoSolicitacao.Alteracao) Then
                    
                        'Código do banco
                        Set objDomNode = .selectSingleNode("CO_BANC")
                        If Not objDomNode Is Nothing Then
                            If Not fgExisteCodigoBanco(Val(objDomNode.Text)) Then
                                'Código do Banco inválido
                                fgAdicionaErro xmlErrosNegocio, 4027
                            End If
                        End If
                        
                        'Código da Agência
                        Set objDomNode = .selectSingleNode("CO_AGEN")
                        If Not objDomNode Is Nothing Then
                            If Val("0" & objDomNode.Text) = 0 Then
                                'Código da Agência inválido
                                fgAdicionaErro xmlErrosNegocio, 4028
                            End If
                        End If
            
                        'Número da Conta Corrente
                        Set objDomNode = .selectSingleNode("NU_CC")
                        If Not objDomNode Is Nothing Then
                            If Val("0" & objDomNode.Text) = 0 Then
                                'Código da conta corrente inválido
                                fgAdicionaErro xmlErrosNegocio, 4029
                            End If
                        End If
            
                        'Valor do Lançamento Conta Corrente
                        Set objDomNode = .selectSingleNode("VA_LANC_CC")
                        If Not objDomNode Is Nothing Then
                            If fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                                'Valor do lançamento na conta corrente inválido
                                fgAdicionaErro xmlErrosNegocio, 4030
                            Else
                                If lngTipoMensagem = enumTipoMensagemLQS.OperacoesComCorretorasCETIP Then
                                    If fgVlrXml_To_Decimal(objDomNode.Text) <> fgVlrXml_To_Decimal(.selectSingleNode("//VA_OPER_ATIV").Text) Then
                                        'Valor do lançamento na conta corrente inválido
                                        fgAdicionaErro xmlErrosNegocio, 4030
                                    End If
                                End If
                            End If
                        End If
                    
                    End If
                
                End If
                
            End If
            
        #End If
            
'Repetições-----------------------------------------

        'Repetição detalhes operações  - 50 - Operações Com Corretoras
        If Not .selectSingleNode("REPE_DTLH_OPER") Is Nothing Then
            For Each objDomNodeRepet1 In .selectSingleNode("REPE_DTLH_OPER").childNodes
            
                'Indicador Compra/Venda
                Set objDomNode = objDomNodeRepet1.selectSingleNode("IN_CPRA_VEND")
                If Not objDomNode Is Nothing Then
                    If Trim$(objDomNode.Text) <> vbNullString And _
                       UCase$(objDomNode.Text) <> "C" And _
                       UCase$(objDomNode.Text) <> "V" Then
                       'Indicador Compra/Venda inválido
                       fgAdicionaErro xmlErrosNegocio, 4127
                    End If
                End If
                
            Next objDomNodeRepet1
        End If
            
        'Repetição Informação de contrato - 72 - Registro de contrato Swap
        If Not .selectSingleNode("REPE_INFO_CNTR") Is Nothing Then
            
            For Each objDomNodeRepet1 In .selectNodes("REPE_INFO_CNTR/*")
                
                'Indicador Participante/Contraparte
                Set objDomNode = objDomNodeRepet1.selectSingleNode("IN_PARP_CNPT")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text <> "C" And _
                       objDomNode.Text <> "P" Then
                        'Indicador Participante/Contraparte inválido
                        fgAdicionaErro xmlErrosNegocio, 4171
                    End If
                End If
                
                'Código indexador Cetip
                Set objDomNode = objDomNodeRepet1.selectSingleNode("TP_CLIE")
                If Not objDomNode Is Nothing Then
                    If Trim(objDomNode.Text) <> vbNullString Then
                        If Not fgExisteTipoCliente(objDomNode.Text, objDomNode.nodeName) Then
                            'Tipo de Cliente Inválido
                            fgAdicionaErro xmlErrosNegocio, 4183
                        End If
                    End If
                End If
                
                'Código indexador Cetip
                Set objDomNode = objDomNodeRepet1.selectSingleNode("CO_INDX_CETIP")
                If Not objDomNode Is Nothing Then
                    If Trim$(objDomNode.Text) <> vbNullString Then
                        If Not fgExisteIndexadorCetip(objDomNode.Text, objDomNode.nodeName) Then
                            'Código Indexador Cetip Inválido
                            fgAdicionaErro xmlErrosNegocio, 4184
                        End If
                    End If
                End If
                
                'Tipo Indexador Especial CETIP
                Set objDomNode = objDomNodeRepet1.selectSingleNode("TP_INDX_ESPC")
                If Not objDomNode Is Nothing Then
                    If Not Val("0" & objDomNode.Text) = 0 Then
                        If Not fgExisteIndexadorEspecialCetip(objDomNode.Text, objDomNode.nodeName) Then
                            'Tipo Indexador Especial CETIP
                            fgAdicionaErro xmlErrosNegocio, 4215
                        End If
                    End If
                End If
                
                'Repetição Informação limite
                If Not objDomNodeRepet1.selectSingleNode("REPE_INFO_LIMI") Is Nothing Then
                     For Each objDomNodeRepet2 In objDomNodeRepet1.selectSingleNode("REPE_INFO_LIMI").childNodes
                        
                        'Indicador Limite
                        Set objDomNode = objDomNodeRepet2.selectSingleNode("IN_LIMI")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) <> vbNullString Then
                                If UCase$(objDomNode.Text) <> "I" And _
                                   UCase$(objDomNode.Text) <> "S" Then
                                    'Indicador de Limite inválido
                                    fgAdicionaErro xmlErrosNegocio, 4172
                                End If
                            End If
                        End If
                        
                    Next objDomNodeRepet2
                End If
                    
            Next objDomNodeRepet1
        End If
        
        'Repetição Informação Cliente 4- 74 - Informações Cliente
        If Not .selectSingleNode("REPE_INFO_CLIE") Is Nothing Then
            For Each objDomNodeRepet1 In .selectSingleNode("REPE_INFO_CLIE").childNodes
                'Indicador Participante/Contraparte
                Set objDomNode = objDomNodeRepet1.selectSingleNode("IN_PARP_CNPT")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text <> "C" And _
                       objDomNode.Text <> "P" Then
                        'Indicador Participante/Contraparte inválido
                        fgAdicionaErro xmlErrosNegocio, 4173
                    End If
                End If
            Next objDomNodeRepet1
        End If
        
        'Repetição Informação Trigger - 74 - Registro de dados complementares de contrato de SWAP
        If Not .selectSingleNode("REPE_INFO_TGAR") Is Nothing Then
            For Each objDomNodeRepet1 In .selectSingleNode("REPE_INFO_TGAR").childNodes
                
                If objDomNodeRepet1.baseName = "GRUP_INFO_TGAR" Then
                    'Tipo Trigger
                    Set objDomNode = objDomNodeRepet1.selectSingleNode("TP_TGAR_CNTR_SWAP")
                    If Trim$(objDomNode.Text) <> vbNullString Then
                        If Not fgExisteTipoTrigger(objDomNode.Text) Then
                            'Tipo de Trigger inválido
                            fgAdicionaErro xmlErrosNegocio, 4142
                        End If
                    End If
                    
                    'Indicador de Periodicidade
                    Set objDomNode = objDomNodeRepet1.selectSingleNode("IN_PERI")
                    If Not fgExisteIndicadorPeriodicidade(objDomNode.Text) Then
                        'Indicador de Periodicidade inválido
                        fgAdicionaErro xmlErrosNegocio, 4143
                    End If
                End If
                
            Next objDomNodeRepet1
        End If
        
        'Repetição informação prêmio - 74 - Registro de dados complementaares de contrato de SWAP
        'Indicador Forma Liquidacao Rebate
        
        Set objDomNode = .selectSingleNode("//CO_FORM_LIQU_REBA")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) <> vbNullString Then
                If Not fgExisteIndicadorFormaLiquRebate(objDomNode.Text) Then
                    'Indicador Forma Liquidacao Rebate inválido
                    fgAdicionaErro xmlErrosNegocio, 4174
                End If
            End If
        End If
        
        If blnValidarCC Then
            
            'Data do exercício
            Set objDomNode = .selectSingleNode("//DT_EXER")
            If Not objDomNode Is Nothing Then
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data do exercício inválida
                    fgAdicionaErro xmlErrosNegocio, 4146
                End If
            End If
            
        End If
        
        'Repetição Parâmetro referência - 80 - Antecipação Resgate Contrato derivativo
        If Not .selectSingleNode("REPE_PARM_REFE") Is Nothing Then
            For Each objDomNodeRepet1 In .selectSingleNode("REPE_PARM_REFE").childNodes
                
                'Indicador Participante/Contraparte
                Set objDomNode = objDomNodeRepet1.selectSingleNode("IN_PARP_CNPT")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text <> "C" And _
                       objDomNode.Text <> "P" Then
                        'Indicador Participante/Contraparte inválido
                        fgAdicionaErro xmlErrosNegocio, 4173
                    End If
                End If
                
            Next objDomNodeRepet1
        End If
        
        'Repetição Preço unitário ou fator SWAP - 82 - Lançamento de PU ou Fator de conteúdo derivativo
        If Not .selectSingleNode("REPE_INFO_PU_FATR_SWAP") Is Nothing Then
            For Each objDomNodeRepet1 In .selectSingleNode("REPE_INFO_PU_FATR_SWAP").childNodes
                
                'Indicador Participante/Contraparte
                Set objDomNode = objDomNodeRepet1.selectSingleNode("IN_PARP_CNPT")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text <> "C" And _
                       objDomNode.Text <> "P" Then
                        'Indicador Participante/Contraparte inválido
                        fgAdicionaErro xmlErrosNegocio, 4173
                    End If
                End If
                
                'Indicador de Tipo de referência
                Set objDomNode = objDomNodeRepet1.selectSingleNode("TP_REFE")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text <> "F" And _
                       objDomNode.Text <> "P" Then
                        'Indicador de Tipo de referência inválido
                        fgAdicionaErro xmlErrosNegocio, 4175
                    End If
                End If
                
                
            Next objDomNodeRepet1
        End If
        
        If lngTipoMensagem = enumTipoMensagemLQS.RegistroOperacaoesCETIP Then
            Call fgValidaMensagemDadosComplContratoSWAP(xmlDOMMensagem, xmlErrosNegocio, strDataOperacao)
        End If

        Select Case strSiglaSistemaCETIP
            
            Case "CETIP"
                Select Case lngTipoMensagem
                    Case enumTipoMensagemLQS.MovimentacoesInstFinancCETIP
                        
                        'Tipo Rentabilidade
                        Set objDomNode = .selectSingleNode("TP_RENT")
                        If Not objDomNode Is Nothing Then
                            If Not IsNumeric(objDomNode.Text) Then
                                'Tipo de rentabilida inválida
                                fgAdicionaErro xmlErrosNegocio, 4198
                            ElseIf objDomNode.Text <> enumTipoRentabilidade.Pre And _
                                   objDomNode.Text <> enumTipoRentabilidade.pos Then
                                'Tipo de rentabilida inválida
                                fgAdicionaErro xmlErrosNegocio, 4198
                            Else
                                lngTipoRentabilidade = objDomNode.Text
                            End If
                        End If
                        
                        If lngTipoRentabilidade = enumTipoRentabilidade.Pre Then
                            'Valor resgate
                            Set objDomNode = .selectSingleNode("VA_FINC_BASE_RESG")
                            If Not objDomNode Is Nothing Then
                                If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                                    'Valor Resgate é obrigatório
                                    fgAdicionaErro xmlErrosNegocio, 4199
                                End If
                            Else
                                'Valor Resgate é obrigatório
                                fgAdicionaErro xmlErrosNegocio, 4199
                            End If
                        End If
                End Select
        
        End Select

        'Data da operação original
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV_ORIG")
        If Not objDomNode Is Nothing Then
            If Val(objDomNode.Text) <> 0 Then
                If Not flValidaDataNumerica(Val(objDomNode.Text)) Then
                    'Data da operação original inválida
                    fgAdicionaErro xmlErrosNegocio, 4164
                ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                    'Data de origem da operação não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4057
                End If
            Else
                Select Case strSiglaSistemaCETIP
                    Case "CETIP21"
                        Select Case lngTipoMensagem
                            Case enumTipoMensagemLQS.AntecipacaoResgateContratoDerivativoCETIP
                                'Data da operação original inválida
                                fgAdicionaErro xmlErrosNegocio, 4164
                        End Select
                End Select
            End If
        Else
            Select Case strSiglaSistemaCETIP
                Case "CETIP21"
                    Select Case lngTipoMensagem
                        Case enumTipoMensagemLQS.AntecipacaoResgateContratoDerivativoCETIP
                            'Data da operação original inválida
                            fgAdicionaErro xmlErrosNegocio, 4164
                    End Select
            End Select
        End If
        
        'Identificador ponta cetip // // Nick (3/3/2016) * Retirar a validação do CO_IDEF_PNTA_CETIP / IN_MANU_PREM
        Set objDomNode = .selectSingleNode("CO_IDEF_PNTA_CETIP")
        
        Select Case strSiglaSistemaCETIP
            Case "CETIP21"
                Select Case lngTipoMensagem
                    Case enumTipoMensagemLQS.AntecipacaoResgateContratoDerivativoCETIP, _
                         enumTipoMensagemLQS.LancamentoPU_CETIP
        
                        If Not .selectSingleNode("CO_SUB_TIPO_ATIV") Is Nothing Then
                            If .selectSingleNode("CO_SUB_TIPO_ATIV").Text = "SWAP" Then
                            
                                If Not objDomNode Is Nothing Then
                                    If Not fgExisteDominioTagMensagemBACEN("IdentdPontaCTP", objDomNode.Text) Then
                                        'Identificador Ponta CETIP ausente ou conteúdo inválido (CO_IDEF_PNTA_CETIP).
                                        fgAdicionaErro xmlErrosNegocio, 4346
                                    End If
                                Else
                                    'Identificador Ponta CETIP ausente ou conteúdo inválido (CO_IDEF_PNTA_CETIP).
                                    fgAdicionaErro xmlErrosNegocio, 4346
                                End If
                            
                            End If
                        End If
                        
                End Select
        End Select
        
        'Indicador Manutenção Prêmio
        Set objDomNode = .selectSingleNode("IN_MANU_PREM")
        
        Select Case strSiglaSistemaCETIP
            Case "CETIP21"
                
                Select Case lngTipoMensagem
                    Case enumTipoMensagemLQS.AntecipacaoResgateContratoDerivativoCETIP
                    
                        If Not .selectSingleNode("CO_SUB_TIPO_ATIV") Is Nothing Then
                            If .selectSingleNode("CO_SUB_TIPO_ATIV").Text = "SWAP" Then
                            
                                If Not objDomNode Is Nothing Then
                                    If Trim(objDomNode.Text) <> vbNullString Then
                                        If UCase(Trim(objDomNode.Text)) <> "S" And _
                                           UCase(Trim(objDomNode.Text)) <> "N" Then
                                            'Indicador Manutenção Prêmio ausente ou conteúdo inválido (IN_MANU_PREM).
                                            fgAdicionaErro xmlErrosNegocio, 4347
                                        End If
                                    End If
                                End If
                                
                            End If
                        End If
                        
                End Select
        End Select
        
        'Tipo Atualização CETIP
        Set objDomNode = .selectSingleNode("TP_ATLZ_CETIP")
        
        Select Case strSiglaSistemaCETIP
            Case "CETIP21"
                Select Case lngTipoMensagem
                    Case enumTipoMensagemLQS.LancamentoPU_CETIP
        
                        If Not objDomNode Is Nothing Then
                            If Not fgExisteDominioTagMensagemBACEN("TpAtlzCTP", objDomNode.Text) Then
                                'Tipo Atualização CETIP ausente ou conteúdo inválido (TP_ATLZ_CETIP).
                                fgAdicionaErro xmlErrosNegocio, 4348
                            End If
                        Else
                            'Tipo Atualização CETIP ausente ou conteúdo inválido (TP_ATLZ_CETIP).
                            fgAdicionaErro xmlErrosNegocio, 4348
                        End If
                        
                End Select
        End Select
        
        'Identificador Curva CETIP
        Set objDomNode = .selectSingleNode("CO_IDEF_CRVA_CETIP")
        
        Select Case strSiglaSistemaCETIP
            Case "CETIP21"
                Select Case lngTipoMensagem
                    Case enumTipoMensagemLQS.LancamentoPU_CETIP
        
                        If Not objDomNode Is Nothing Then
                            If Not fgExisteDominioTagMensagemBACEN("IdentdPontaCTP", objDomNode.Text) Then
                                'Identificador Curva CETIP ausente ou conteúdo inválido (CO_IDEF_CRVA_CETIP).
                                fgAdicionaErro xmlErrosNegocio, 4349
                            End If
                        End If
                        
                End Select
        End Select
        
        'Data de início de resgate
        Set objDomNode = .selectSingleNode("DT_INIC_RESG")
        If Not objDomNode Is Nothing Then
            If strTipoTabelaResgate = "C" Or _
               strTipoTabelaResgate = "M" Or _
               Val(objDomNode.Text) <> 0 Then
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data de início de resgate inválida
                    fgAdicionaErro xmlErrosNegocio, 4238
                End If
            End If
        End If
        
        'Tipo fluxo caixa
        Set objDomNode = .selectSingleNode("TP_FLUX_CAIX")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text <> "0" And _
               objDomNode.Text <> "1" And _
               objDomNode.Text <> "2" Then
                'Tipo de fluxo de caixa inválido
                fgAdicionaErro xmlErrosNegocio, 4240
            End If
        End If
        
        'Adrian - 10/01/06 - Referente ao novo Book do Bacen, V2.1
        'Indicador de Vinculação de Reserva Técnica (spmente layout 52 e 56)
        If lngTipoMensagem = enumTipoMensagemLQS.MovimentacoesInstFinancCETIP Or _
           lngTipoMensagem = enumTipoMensagemLQS.ResgateFundoInvestimentoCETIP Then
            Set objDomNode = .selectSingleNode("IN_VINC_RES_TEC")
            If Not objDomNode Is Nothing Then
                If Trim$(objDomNode.Text) <> vbNullString And _
                   UCase$(Trim$(objDomNode.Text)) <> "S" And _
                   UCase$(Trim$(objDomNode.Text)) <> "N" Then
                    'Indicador de Vinculação de Reserva Técnica inválido
                    fgAdicionaErro xmlErrosNegocio, 4255
                End If
            Else
                If Not .selectSingleNode("SG_SIST_CETIP") Is Nothing Then
                    If .selectSingleNode("SG_SIST_CETIP").Text = "SCF" Then
                        'Indicador de Vinculação de Reserva Técnica obrigatório caso sigla de sistema CETIP igual a SCF
                        fgAdicionaErro xmlErrosNegocio, 4259
                    End If
                End If
            End If
        End If
        
        'Adrian - 10/01/06 - Referente ao novo Book do Bacen, V2.1
        'Indicador Título CETIP Inadimplente
        If lngTipoMensagem = enumTipoMensagemLQS.OperacaoDefinitivaCETIP Then
            Set objDomNode = .selectSingleNode("IN_TIT_CTP_INDP")
            If Not objDomNode Is Nothing Then
                If Trim$(objDomNode.Text) <> vbNullString And _
                   UCase$(Trim$(objDomNode.Text)) <> "S" And _
                   UCase$(Trim$(objDomNode.Text)) <> "N" Then
                    'Indicador de Titulo CETIP Inadimplente inválido
                    fgAdicionaErro xmlErrosNegocio, 4256
                End If
            Else
                If Not .selectSingleNode("CO_SUB_TIPO_ATIV") Is Nothing Then
                    If .selectSingleNode("CO_SUB_TIPO_ATIV").Text = "CCB" Then
                        'Indicador Título CETIP Inadimplente deverá ser preenchido caso Sub Tipo Ativo igual a CCB
                        fgAdicionaErro xmlErrosNegocio, 4265
                    End If
                End If
            End If
        End If
        
        If lngTipoMensagem = enumTipoMensagemLQS.OperacaoCompromissadaCETIP Then
            Set objDomNode = .selectSingleNode("IN_TIT_CTP_INDP")
            If Not objDomNode Is Nothing Then
                If Trim$(objDomNode.Text) <> vbNullString And _
                   UCase$(Trim$(objDomNode.Text)) <> "S" And _
                   UCase$(Trim$(objDomNode.Text)) <> "N" Then
                    'Indicador de Titulo CETIP Inadimplente inválido
                    fgAdicionaErro xmlErrosNegocio, 4256
                End If
            Else
                If strSiglaSistemaCETIP = "SNA" Then
                    If Not .selectSingleNode("CO_SUB_TIPO_ATIV") Is Nothing Then
                        If .selectSingleNode("CO_SUB_TIPO_ATIV").Text = "CCB" Then
                            'Indicador Título CETIP Inadimplente deverá ser preenchido caso Sub Tipo Ativo igual a CCB
                            fgAdicionaErro xmlErrosNegocio, 4265
                        End If
                    End If
                End If
            End If
        
'            'Código identificador de lastro
'            If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
'               lngTipoSolicitacao = enumTipoSolicitacao.CancelamentoPorLastro Then
'                Set objDomNode = .selectSingleNode("CO_IDEF_LAST")
'                If Not objDomNode Is Nothing Then
'                    If Trim$(objDomNode.Text) = vbNullString Then
'                        'Código Identificador de Lastro obrigatório para os tipos de solicitação 2 (Complementação) e 8 (Cancelamento por lastro).
'                        fgAdicionaErro xmlErrosNegocio, 4279
'                    End If
'                Else
'                    'Código Identificador de Lastro obrigatório para os tipos de solicitação 2 (Complementação) e 8 (Cancelamento por lastro).
'                    fgAdicionaErro xmlErrosNegocio, 4279
'                End If
'            End If
'
            If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               lngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then
                'Identificador Título
                Set objDomNode = .selectSingleNode("NU_ATIV_MERC_CETIP")
                If Not objDomNode Is Nothing Then
                    If Trim$(objDomNode.Text) = vbNullString Then
                        'Identificador de Título inválido
                        fgAdicionaErro xmlErrosNegocio, 4074
                    End If
                End If
            End If
        End If
        
        'Adrian - 10/01/06 - Referente ao novo Book do Bacen, V2.1
        'Data de Inicio
        Set objDomNode = .selectSingleNode("DT_INIC")
        If Not objDomNode Is Nothing Then
            strDataInicio = objDomNode.Text
        
            If fgDtXML_To_Date(objDomNode.Text) < fgDtXML_To_Date(.selectSingleNode("DT_OPER_ATIV").Text) Then
                If lngTipoContraparte = enumTipoContraparte.Cliente1 Then
                    If fgDtXML_To_Date(objDomNode.Text) < fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 1, enumPaginacao.Anterior) Then
                        'Data Inicio não pode ser menor que a D-1 da data da Operação para Cliente1
                        fgAdicionaErro xmlErrosNegocio, 4262
                    End If
                Else
                    'Data Inicio não pode ser menor que a data da Operação
                    fgAdicionaErro xmlErrosNegocio, 4260
                End If
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data inicio não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4261
            End If
        End If
        
        'KIDA - SGC
        'Tipo Canal Venda
        Set objDomNode = .selectSingleNode("//TP_CNAL_VEND")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumCanalDeVenda.SGC And _
               Val("0" & objDomNode.Text) <> enumCanalDeVenda.SGM And _
               Val("0" & objDomNode.Text) <> enumCanalDeVenda.Nenhum Then
                'Canal de Venda Inválido.
                fgAdicionaErro xmlErrosNegocio, 4291
            End If
        End If
        
        'Código Tipo Amortização
        Set objDomNode = .selectSingleNode("CO_TIPO_AMTZ")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) <> vbNullString Then
                If Not fgExisteDominioTagMensagemBACEN("TpAmtzc", Trim$(objDomNode.Text)) Then
                    'Código do Tipo de Amortização inválido.
                    fgAdicionaErro xmlErrosNegocio, 4345
                End If
            End If
        End If
        
        'Indicador Ajuste Taxa
        Set objDomNode = .selectSingleNode("IN_AJUS_TAXA")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) <> vbNullString Then
                If UCase(Trim$(objDomNode.Text)) <> "S" _
                And UCase(Trim$(objDomNode.Text)) <> "N" Then
                    'Indicador Ajuste Taxa invalido
                    fgAdicionaErro xmlErrosNegocio, 4459
                End If
            End If
        End If
        
        'Tipo Pagador Responsavel Ajuste Taxa
        Set objDomNode = .selectSingleNode("TP_PAGA_RESP_AJUS_TAXA")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) <> vbNullString Then
                If Trim$(objDomNode.Text) <> CStr(enumTipoPagadorResponsavel.VendedorComprador) _
                And Trim$(objDomNode.Text) <> CStr(enumTipoPagadorResponsavel.Vendedor) _
                And Trim$(objDomNode.Text) <> CStr(enumTipoPagadorResponsavel.Comprador) Then
                    'Tipo Pagador Responsavel Ajuste Taxa invalido
                    fgAdicionaErro xmlErrosNegocio, 4460
                End If
            End If
        End If
        
        'Indicador Limite
        Set objDomNode = .selectSingleNode("IN_LIMI")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) <> vbNullString Then
                If UCase(Trim$(objDomNode.Text)) <> "I" _
                And UCase(Trim$(objDomNode.Text)) <> "S" _
                And UCase(Trim$(objDomNode.Text)) <> "SI" Then
                    'Indicador de Limite invalido
                    fgAdicionaErro xmlErrosNegocio, 4172
                End If
            End If
        End If
        
        'Tipo Pagador Responsavel Pagamento Premio
        Set objDomNode = .selectSingleNode("TP_PAGA_RESP_PAGTO_PREM")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) <> vbNullString Then
                If Trim$(objDomNode.Text) <> CStr(enumTipoPagadorResponsavel.VendedorComprador) _
                And Trim$(objDomNode.Text) <> CStr(enumTipoPagadorResponsavel.Vendedor) _
                And Trim$(objDomNode.Text) <> CStr(enumTipoPagadorResponsavel.Comprador) Then
                    'Tipo Pagador Responsavel Pagamento Premio invalido
                    fgAdicionaErro xmlErrosNegocio, 4461
                End If
            End If
        End If
    
        ' Nick - Ativos Imobiliários
        ' Validação por tipo de mensagem (Mudança - Ativos Imobiliários)
        
        blnAtivosImobiliarios = False
'        Select Case strSubTipoAtivo
'            Case "CCI", "CRI", "LCI", "LH"
'                blnAtivosImobiliarios = True
'        End Select
        
        If blnAtivosImobiliarios And lngTipoContraparte = enumTipoContraparte.Cliente1 Then
        
            Select Case lngTipoMensagem
                Case enumTipoMensagemLQS.MovimentacoesInstFinancCETIP       '52
                
                    ' Valida Código da Operação CETIP = 11 / 1 ou 2
                    If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0011.RetiradaCustodia Or _
                       lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0011.DepositoEmissaoSemFinaceiro Or _
                       lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0011.DepositoEmissaoAplicacaoFinanceiro Then
                                                
                        'Verifica se os campos obrigatórios estão preenchidos
                        'Tipo de Pessoa
                        Set objDomNode = .selectSingleNode("TP_PESS")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4329
                            Else
                                If UCase$(Trim$(objDomNode.Text)) <> "F" And _
                                   UCase$(Trim$(objDomNode.Text)) <> "J" Then
                                    'Valor diferente do esperado
                                    fgAdicionaErro xmlErrosNegocio, 4329
                                End If
                            End If
                        Else
                            ' Erro: TAG TP_PESS Ausente
                            fgAdicionaErro xmlErrosNegocio, 4329
                        End If
                        
                        'CNPJ da Contraparte
                        Set objDomNode = .selectSingleNode("CO_CNPJ_CNPT")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4330
                            End If
                        Else
                            ' Erro: TAG CO_CNPJ_CNPT Ausente
                            fgAdicionaErro xmlErrosNegocio, 4330
                        End If
                    
                    End If
                    
                    ' Valida Código da Operação CETIP = 11
                    If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0011.RetiradaCustodia Then
'                        'Tipo de Operação Reserva Técnica
'                        Set objDomNode = .selectSingleNode("TP_OPER_RESE_TECN")
'                        If Not objDomNode Is Nothing Then
'                            If Trim$(objDomNode.Text) = vbNullString Then
'                                'Campo obrigatório Nullo
'                                fgAdicionaErro xmlErrosNegocio, 4331
'                            End If
'                        Else
'                            ' Erro: TAG TP_OPER_RESE_TECN Ausente
'                            fgAdicionaErro xmlErrosNegocio, 4331
'                        End If
'
                        'Data Operação Original
                        Set objDomNode = .selectSingleNode("DT_OPER_ATIV_ORIG")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4337
                            End If
                        Else
                            ' Erro: DT_OPER_ATIV_ORIG Ausente
                            fgAdicionaErro xmlErrosNegocio, 4337
                        End If
                                            
                        'Número Operação CTP Original
                        Set objDomNode = .selectSingleNode("NU_COMD_OPER_ORIG")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4338
                            End If
                        Else
                            ' Erro: NU_COMD_OPER_ORIG Ausente
                            fgAdicionaErro xmlErrosNegocio, 4338
                        End If
                    End If

                Case enumTipoMensagemLQS.MovimentacoesCustodiaCETIP         '54
                
                    ' Valida Código da Operação CETIP = 11 / 1 ou 2
                    If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0011.RetiradaCustodia Or _
                       lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0011.DepositoEmissaoSemFinaceiro Or _
                       lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0011.DepositoEmissaoAplicacaoFinanceiro Then
                                                
                        'Verifica se os campos obrigatórios estão preenchidos
                        'Tipo de Pessoa
                        Set objDomNode = .selectSingleNode("TP_PESS")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4329
                            Else
                                If UCase$(Trim$(objDomNode.Text)) <> "F" And _
                                   UCase$(Trim$(objDomNode.Text)) <> "J" Then
                                    'Valor diferente do esperado
                                    fgAdicionaErro xmlErrosNegocio, 4329
                                End If
                            End If
                        Else
                            ' Erro: TAG TP_PESS Ausente
                            fgAdicionaErro xmlErrosNegocio, 4329
                        End If
                        
                        'CNPJ da Contraparte
                        Set objDomNode = .selectSingleNode("CO_CNPJ_CNPT")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4330
                            End If
                        Else
                            ' Erro: TAG CO_CNPJ_CNPT Ausente
                            fgAdicionaErro xmlErrosNegocio, 4330
                        End If
                    
                    End If
                    
                    ' Valida Código da Operação CETIP = 11
                    If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0011.RetiradaCustodia Then
                        'Tipo de Operação Reserva Técnica
                        Set objDomNode = .selectSingleNode("TP_OPER_RESE_TECN")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) <> "D" And _
                               Trim$(objDomNode.Text) <> "V" Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4331
                            End If
                        Else
                            ' Erro: TAG TP_OPER_RESE_TECN Ausente
                            fgAdicionaErro xmlErrosNegocio, 4331
                        End If
                                            
                        'Data Operação Original
                        Set objDomNode = .selectSingleNode("DT_OPER_ATIV_ORIG")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4337
                            End If
                        Else
                            ' Erro: DT_OPER_ATIV_ORIG Ausente
                            fgAdicionaErro xmlErrosNegocio, 4337
                        End If
                                            
                        'Número Operação CTP Original
                        Set objDomNode = .selectSingleNode("NU_COMD_OPER_ORIG")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4338
                            End If
                        Else
                            ' Erro: NU_COMD_OPER_ORIG Ausente
                            fgAdicionaErro xmlErrosNegocio, 4338
                        End If
                    End If

                    ' Valida Código da Operação CETIP = 13
                    If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0013.Caucao Then
                        'Tipo de Garantia
                        Set objDomNode = .selectSingleNode("TP_GRTA")
                        If Not objDomNode Is Nothing Then
                            If Not fgExisteDominioTagMensagemBACEN("TpGar", Trim$(objDomNode.Text)) Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4339
                            End If
                        Else
                            ' Erro: TP_GRTA Ausente
                            fgAdicionaErro xmlErrosNegocio, 4339
                        End If
                    
                        'Indicador Título CETIP Inadimplente
                        Set objDomNode = .selectSingleNode("IN_TIT_CTP_INDP")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) <> "S" And _
                               Trim$(objDomNode.Text) <> "N" Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4344
                            End If
                        Else
                            ' Erro: TAG IN_TIT_CTP_INDP Ausente
                            fgAdicionaErro xmlErrosNegocio, 4344
                        End If
                    End If
                                            
                    ' Valida Código da Operação CETIP = 71
                    If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0013.VinculacaoDesvinculacaoContaMargem Then
                        'Indicador Título CETIP Inadimplente
                        Set objDomNode = .selectSingleNode("IN_TIT_CTP_INDP")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) <> "S" And _
                               Trim$(objDomNode.Text) <> "N" Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4344
                            End If
                        Else
                            ' Erro: TAG IN_TIT_CTP_INDP Ausente
                            fgAdicionaErro xmlErrosNegocio, 4344
                        End If
                    End If
                                            
                    ' Valida Código da Operação CETIP = 75
                    If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0021.TransferenciaGarantiasContaMargemCamara Then
                        'Identificador Investidor Câmara
                        Set objDomNode = .selectSingleNode("CO_INVE_CAMR")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4340
                            End If
                        Else
                            ' Erro: CO_INVE_CAMR Ausente
                            fgAdicionaErro xmlErrosNegocio, 4340
                        End If
                    End If
               
                Case enumTipoMensagemLQS.OperacaoDefinitivaCETIP '64
                
                    ' Valida Código da Operação CETIP = 52 / 14
                    If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0052.CompraVendaDefinitiva Or _
                       lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0052.ResgateAntecipado Then
                                                    
                        'Verifica se os campos obrigatórios estão preenchidos
                        'Tipo de Pessoa
                        Set objDomNode = .selectSingleNode("TP_PESS")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4329
                            Else
                                If UCase$(Trim$(objDomNode.Text)) <> "F" And _
                                   UCase$(Trim$(objDomNode.Text)) <> "J" Then
                                    'Valor diferente do esperado
                                    fgAdicionaErro xmlErrosNegocio, 4329
                                End If
                            End If
                        Else
                            ' Erro: TAG TP_PESS Ausente
                            fgAdicionaErro xmlErrosNegocio, 4329
                        End If
                                                    
                        'CNPJ da Contraparte
                        Set objDomNode = .selectSingleNode("CO_CNPJ_CNPT")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4330
                            End If
                        Else
                            ' Erro: TAG CO_CNPJ_CNPT Ausente
                            fgAdicionaErro xmlErrosNegocio, 4330
                        End If
                                                    
                        'Tipo de Operação Reserva Técnica
                        Set objDomNode = .selectSingleNode("TP_OPER_RESE_TECN")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) <> "D" And _
                               Trim$(objDomNode.Text) <> "V" Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4331
                            End If
                        Else
                            ' Erro: TAG TP_OPER_RESE_TECN Ausente
                            fgAdicionaErro xmlErrosNegocio, 4331
                        End If
                    
                    End If
            
                Case enumTipoMensagemLQS.OperacaoCompromissadaCETIP '66
            
                    ' Valida Código da Operação CETIP = 57 / 357
                    If lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0054.CompraVendaCompromissadaRentaPosFix Or _
                       lngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0054.RegistroRetroativoCompraVendaComprRentaPosFix Then
                    
                        'Verifica se os campos obrigatórios estão preenchidos
                        'Descrição Índice Valor Contratado Partes
                        Set objDomNode = .selectSingleNode("DS_INDC_VALR_CNTR_PART")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4332
                            End If
                        Else
                            ' Erro: TAG DS_INDC_VALR_CNTR_PART Ausente
                            fgAdicionaErro xmlErrosNegocio, 4332
                        End If
                                                        
                        'Tipo Indexador Índice Valor Contratado Partes
                        Set objDomNode = .selectSingleNode("TP_INDX_INDC_VALR_CNTR_PART")
                        If Not objDomNode Is Nothing Then
                            If Not fgExisteDominioTagMensagemBACEN("TpIndx", Trim$(objDomNode.Text)) Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4333
                            End If
                        Else
                            ' Erro: TAG TP_INDX_INDC_VALR_CNTR_PART Ausente
                            fgAdicionaErro xmlErrosNegocio, 4333
                        End If
                                                        
                        'Percentual Parâmetro Juros CETIP
                        Set objDomNode = .selectSingleNode("PE_PARM_JURO")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4334
                            End If
                        Else
                            ' Erro: TAG PE_PARM_JURO Ausente
                            fgAdicionaErro xmlErrosNegocio, 4334
                        End If
                                                        
                        'Taxa Juros CETIP
                        Set objDomNode = .selectSingleNode("PE_TAXA_JURO_CETIP")
                        If Not objDomNode Is Nothing Then
                            If Trim$(objDomNode.Text) = vbNullString Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4335
                            End If
                        Else
                            ' Erro: TAG PE_TAXA_JURO_CETIP Ausente
                            fgAdicionaErro xmlErrosNegocio, 4335
                        End If
                                                        
                        'Critério Cálculo Juros
                        Set objDomNode = .selectSingleNode("CR_CALC_JURO")
                        If Not objDomNode Is Nothing Then
                            If Not fgExisteDominioTagMensagemBACEN("CritCalcJuros", Trim$(objDomNode.Text)) Then
                                'Campo obrigatório Nullo
                                fgAdicionaErro xmlErrosNegocio, 4336
                            End If
                        Else
                            ' Erro: TAG CR_CALC_JURO Ausente
                            fgAdicionaErro xmlErrosNegocio, 4336
                        End If
                    
                    End If
            
            End Select
        End If
        'Validações referentes aos dados do Lote para Operações do Layout 50 do sistem LQC
        If lngTipoMensagem = enumTipoMensagemLQS.OperacoesComCorretorasCETIP _
            And strSiglaSistemaOrigem = "LQC" _
            And lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                If strSiglaSistemaOrigem = "LQC" Then
        
                    'Sequencial do Lote
                    Set objDomNode = .selectSingleNode("ID_LOTE")
                    If Not objDomNode Is Nothing Then
                        If objDomNode.Text = "" Then
                                'Identificador do Lote é Obrigatório
                                fgAdicionaErro xmlErrosNegocio, 4508
                        End If
                    Else
                        'Identificador do Lote é Obrigatório
                        fgAdicionaErro xmlErrosNegocio, 4508
                    End If
                    
                    'Tipo Débito e Crédito do Lote
                    Set objDomNode = .selectSingleNode("TP_DEB_CRED_LOTE")
                    If Not objDomNode Is Nothing Then
                        If objDomNode.Text = "" Then
                                'Tipo Débito e Crédito do Lote é Obrigatório.
                                fgAdicionaErro xmlErrosNegocio, 4509
                        End If
                        
                        lngTemp = Val(objDomNode.Text)
                        If lngTemp <> enumTipoDebitoCredito.Credito And _
                           lngTemp <> enumTipoDebitoCredito.Debito Then
                            'Indicador de débito/crédito inválido
                            fgAdicionaErro xmlErrosNegocio, 4021
                        End If
                    Else
                        'Tipo Débito e Crédito do Lote é Obrigatório.
                        fgAdicionaErro xmlErrosNegocio, 4509
                    End If
                    
                    'Valor total do Lote
                    Set objDomNode = .selectSingleNode("VA_TOT_LOTE")
                    If Not objDomNode Is Nothing Then
                        If fgVlrXml_To_Decimal(objDomNode.Text) < 0 Then
                            'Valor Total do Lote deve ser Maior ou Igual a Zero.
                            fgAdicionaErro xmlErrosNegocio, 4510
                        End If
                    Else
                        'Valor Total do Lote é Obrigatório.
                        fgAdicionaErro xmlErrosNegocio, 4511
                    End If
                        
                    
                    'Quantidade Total do Lote
                    Set objDomNode = .selectSingleNode("QT_OPER_LOTE")
                    If Not objDomNode Is Nothing Then
                        If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                            'Quantidade Total do Lote deve ser Maior que Zero.
                            fgAdicionaErro xmlErrosNegocio, 4512
                        End If
                    Else
                        'Quantidade Total do Lote é Obrigatório.
                        fgAdicionaErro xmlErrosNegocio, 4513
                    End If
                    
                End If
        End If
        
    End With
    
    


    
    fgConsisteMensagemCETIP = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    Exit Function

ErrorHandler:
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteMensagemCETIP", 0

End Function

Public Function fgExisteIndicadorPeriodicidade(ByVal pvntIndicadorPeriodicidade As Variant) As Boolean

Dim rsIndicadorPeriodicidade                As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsIndicadorPeriodicidade = fgQueryDominio("IN_PERI")
    
    With rsIndicadorPeriodicidade
    
        .Filter = " CO_DOMI = '" & pvntIndicadorPeriodicidade & "'"
    
        If .RecordCount > 0 Then
            gvntIndicadorPeriodicidade = pvntIndicadorPeriodicidade
            fgExisteIndicadorPeriodicidade = True
        End If
    End With
    Set rsIndicadorPeriodicidade = Nothing
Exit Function
ErrorHandler:
    Set rsIndicadorPeriodicidade = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteIndicadorPeriodicidade", 0
End Function

Public Function fgExisteIndicadorBoletim(ByVal plngIndicadorBoletim As Long) As Boolean

Dim rsIndicadorBoletim                      As ADODB.Recordset

On Error GoTo ErrorHandler
   
    Set rsIndicadorBoletim = fgQueryDominio("IN_BOLE")
    
    With rsIndicadorBoletim
    
        .Filter = " CO_DOMI = " & plngIndicadorBoletim
    
        If .RecordCount > 0 Then
            glngIndicadorBoletim = plngIndicadorBoletim
            fgExisteIndicadorBoletim = True
        End If
    End With
    
    Set rsIndicadorBoletim = Nothing
    
    Exit Function
ErrorHandler:
    
    Set rsIndicadorBoletim = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteIndicadorBoletim", 0

End Function

Public Function fgExisteIndexadorCetip(ByVal pvntCodigoIndexadorCetip As Variant, _
                                       ByVal pstrNomeTag As String) As Boolean

Dim rsIndexadorCetip                        As ADODB.Recordset

On Error GoTo ErrorHandler
    
    Set rsIndexadorCetip = fgQueryDominio(pstrNomeTag)
    
    With rsIndexadorCetip
    
        .Filter = " CO_DOMI = '" & CStr(pvntCodigoIndexadorCetip) & "'"
    
        If .RecordCount > 0 Then
            gvntCodigoIndexadorCetip = pvntCodigoIndexadorCetip
            fgExisteIndexadorCetip = True
        End If
    End With
    Set rsIndexadorCetip = Nothing
Exit Function
ErrorHandler:
    Set rsIndexadorCetip = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteIndexadorCetip", 0

End Function

Public Function fgExisteCodigoIndexadorTermoCetip(ByVal pvntCodigoIndexadorTermoCetip As Variant, _
                                                  ByVal pstrNomeTag As String) As Boolean

Dim rsCodigoIndexadorTermoCetip             As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsCodigoIndexadorTermoCetip = fgQueryDominio(pstrNomeTag)
    
    With rsCodigoIndexadorTermoCetip
    
        .Filter = " CO_DOMI = '" & CStr(pvntCodigoIndexadorTermoCetip) & "'"
    
        If .RecordCount > 0 Then
            gvntCodigoIndexadorTermoCetip = pvntCodigoIndexadorTermoCetip
            fgExisteCodigoIndexadorTermoCetip = True
        End If
    End With
    
    Set rsCodigoIndexadorTermoCetip = Nothing
    
Exit Function
ErrorHandler:
    
    Set rsCodigoIndexadorTermoCetip = Nothing
    
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteCodigoIndexadorTermoCetip", 0

End Function

Public Function fgExisteTipoIndexadorTermoCetip(ByVal pvntTipoIndexadorTermoCetip As Variant, _
                                                ByVal pstrNomeTag As String) As Boolean

Dim rsTipoIndexadorTermoCetip               As ADODB.Recordset

On Error GoTo ErrorHandler
    
    Set rsTipoIndexadorTermoCetip = fgQueryDominio(pstrNomeTag)
    
    With rsTipoIndexadorTermoCetip
    
        .Filter = " CO_DOMI = '" & CStr(pvntTipoIndexadorTermoCetip) & "'"
    
        If .RecordCount > 0 Then
            gvntTipoIndexadorTermoCetip = pvntTipoIndexadorTermoCetip
            fgExisteTipoIndexadorTermoCetip = True
        End If
    End With
    
    Set rsTipoIndexadorTermoCetip = Nothing
Exit Function
ErrorHandler:
    Set rsTipoIndexadorTermoCetip = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoIndexadorTermoCetip", 0

End Function


Public Function fgExisteIndexadorEspecialCetip(ByVal pvntCodigoIndexadorEspecialCetip As Variant, _
                                               ByVal pstrNomeTag As String) As Boolean

Dim rsIndexadorEspecialCetip                As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsIndexadorEspecialCetip = fgQueryDominio(pstrNomeTag)
    
    With rsIndexadorEspecialCetip
    
        .Filter = " CO_DOMI = '" & CStr(pvntCodigoIndexadorEspecialCetip) & "'"
    
        If .RecordCount > 0 Then
            gvntCodigoIndexadorEspecialCetip = pvntCodigoIndexadorEspecialCetip
            fgExisteIndexadorEspecialCetip = True
        End If
    End With
    
    Set rsIndexadorEspecialCetip = Nothing
    
    Exit Function
ErrorHandler:
    
    Set rsIndexadorEspecialCetip = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteIndexadorEspecialCetip", 0

End Function

Public Function fgExisteTipoCliente(ByVal pvntTipoCliente As Variant, _
                                    ByVal pstrNomeTag As String) As Boolean

Dim rsTipoCliente                           As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsTipoCliente = fgQueryDominio(pstrNomeTag)
    
    With rsTipoCliente
    
        .Filter = " CO_DOMI = " & pvntTipoCliente
    
        If .RecordCount > 0 Then
            gvntTipoCliente = pvntTipoCliente
            fgExisteTipoCliente = True
        End If
    End With
    
    Set rsTipoCliente = Nothing
    
Exit Function
ErrorHandler:
    Set rsTipoCliente = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoCliente", 0

End Function

Public Function fgExisteIndicadorExecOpcao(ByVal pstrIndicadorExecOpacao As String, _
                                           ByVal pstrNomeTag As String) As Boolean

Dim rsIndicadorExecrcOpcao                  As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsIndicadorExecrcOpcao = fgQueryDominio(pstrNomeTag)
    
    With rsIndicadorExecrcOpcao
    
        .Filter = " CO_DOMI = '" & pstrIndicadorExecOpacao & "'"
    
        If .RecordCount > 0 Then
            gstrIndicadorExecrOpacao = pstrIndicadorExecOpacao
            fgExisteIndicadorExecOpcao = True
        End If
    End With
    
    Set rsIndicadorExecrcOpcao = Nothing
Exit Function
ErrorHandler:
    Set rsIndicadorExecrcOpcao = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteIndicadorExecOpcao", 0

End Function


'Concluir
Public Function fgExisteTipoTrigger(ByVal pvntTipoTrigger As Variant) As Boolean

Dim rsTipoTrigger                           As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsTipoTrigger = fgQueryDominio("TP_TGAR_CNTR_SWAP")
    
    With rsTipoTrigger
    
        .Filter = " CO_DOMI = '" & pvntTipoTrigger & "'"
    
        If .RecordCount > 0 Then
            gvntTipoTrigger = pvntTipoTrigger
            fgExisteTipoTrigger = True
        End If
    End With
    Set rsTipoTrigger = Nothing
Exit Function
ErrorHandler:
    Set rsTipoTrigger = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoTrigger", 0

End Function


'Concluir
Public Function fgExisteTipoFonte(ByVal plngTipoFonte As Long) As Boolean

Dim rsTipoFonte                             As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsTipoFonte = fgQueryDominio("TP_FONT")
    
    With rsTipoFonte
    
        .Filter = " CO_DOMI = " & plngTipoFonte
    
        If .RecordCount > 0 Then
            glngTipoFonte = plngTipoFonte
            fgExisteTipoFonte = True
        End If
    End With
    Set rsTipoFonte = Nothing
Exit Function
ErrorHandler:
    Set rsTipoFonte = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoFonte", 0

End Function


Public Function fgExisteFinalidadeIF(ByVal plngFinalidadeIF As Long) As Boolean

Dim rsFinalidadeIF                          As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsFinalidadeIF = fgQueryDominio("CO_FIND_IF")
    
    With rsFinalidadeIF
    
        .Filter = " CO_DOMI = " & plngFinalidadeIF
    
        If .RecordCount > 0 Then
            glngFinalidadeIF = plngFinalidadeIF
            fgExisteFinalidadeIF = True
        End If
    End With
    Set rsFinalidadeIF = Nothing
Exit Function
ErrorHandler:
    Set rsFinalidadeIF = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteFinalidadeIF", 0

End Function


Public Function fgExisteTipoMovimento(ByVal plngTipoMovimento As Long) As Boolean

Dim rsTipoMovimentoFinanceiro               As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsTipoMovimentoFinanceiro = fgQueryDominioInterno("TP_MOVI_FINC")
    
    With rsTipoMovimentoFinanceiro
    
        .Filter = " CO_DOMI = " & plngTipoMovimento
    
        If .RecordCount > 0 Then
            glngTipoMovimento = plngTipoMovimento
            fgExisteTipoMovimento = True
        End If
    End With
    Set rsTipoMovimentoFinanceiro = Nothing
Exit Function
ErrorHandler:
    Set rsTipoMovimentoFinanceiro = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoMovimento", 0


End Function

Public Function fgExisteMoeda(ByVal plngCodigoMoeda As Long) As Boolean

Dim rsCodigoMoeda                           As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsCodigoMoeda = fgQueryDominio("CO_MOED")
    
    With rsCodigoMoeda
    
        .Filter = " CO_DOMI = " & plngCodigoMoeda
    
        If .RecordCount > 0 Then
            glngCodigoMoeda = plngCodigoMoeda
            fgExisteMoeda = True
        End If
    End With
    Set rsCodigoMoeda = Nothing
Exit Function
ErrorHandler:
    Set rsCodigoMoeda = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteMoeda", 0

End Function


Public Function fgExisteModalidadeLiquidacao(ByVal plngModalidadeLiquidacao As Long) As Boolean

Dim rsModalidadeLiquidacao                  As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsModalidadeLiquidacao = fgQueryDominio("CO_MODA_LIQU_FINC")
   
    With rsModalidadeLiquidacao
    
        .Filter = " CO_DOMI = " & plngModalidadeLiquidacao
    
        If .RecordCount > 0 Then
            glngModalidadeLiquidacao = plngModalidadeLiquidacao
            fgExisteModalidadeLiquidacao = True
        End If
    End With
    Set rsModalidadeLiquidacao = Nothing
Exit Function
ErrorHandler:
    Set rsModalidadeLiquidacao = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteModalidadeLiquidacao", 0

End Function

Public Function fgExisteTipoContratoSwap(ByVal plngTipoContratoSwap As Long) As Boolean

Dim rsTipoContratoSwap                      As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsTipoContratoSwap = fgQueryDominio("TP_CNTR_SWAP")
    
    With rsTipoContratoSwap
    
        .Filter = " CO_DOMI = " & plngTipoContratoSwap
    
        If .RecordCount > 0 Then
            glngTipoContratoSwap = plngTipoContratoSwap
            fgExisteTipoContratoSwap = True
        End If
    End With
    Set rsTipoContratoSwap = Nothing
Exit Function
ErrorHandler:
    Set rsTipoContratoSwap = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteTipoContratoSwap", 0

End Function


Public Function fgExisteSubTipoAtivo(ByVal pstrSubTipoAtivo As String) As Boolean

Dim rsSubTipoAtivo                          As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsSubTipoAtivo = fgQueryDominio("CO_SUB_TIPO_ATIV")
    
    With rsSubTipoAtivo
    
        .Filter = " CO_DOMI = '" & pstrSubTipoAtivo & "' "
    
        If .RecordCount > 0 Then
            gstrSubTipoAtivo = pstrSubTipoAtivo
            fgExisteSubTipoAtivo = True
        End If
    End With
    Set rsSubTipoAtivo = Nothing
Exit Function
ErrorHandler:
    Set rsSubTipoAtivo = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteSubTipoAtivo", 0
End Function

' Verificar a validade da sigla de sistema CETIP informada para o Layout.

Public Function fgExisteSiglaSistemaCETIP(ByVal plngNumeroLayout As Long, _
                                          ByVal pstrSiglaSistema As String) As Boolean

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlSiglaSistema                         As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
   
    Set xmlSiglaSistema = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlSiglaSistema.Load(App.Path & "\SiglaSistemaCETIP_Dominios.xml")
    
    If xmlSiglaSistema.xml = vbNullString Then Exit Function
    
    For Each objDomNode In xmlSiglaSistema.selectNodes("//Grupo_Layout/Layout[Numero=" & plngNumeroLayout & "]/*")
        If objDomNode.Text = pstrSiglaSistema Then
            fgExisteSiglaSistemaCETIP = True
            Exit Function
        End If
    Next
    
    Set xmlSiglaSistema = Nothing
    Exit Function

ErrorHandler:
    Set xmlSiglaSistema = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteSiglaSistemaCETIP", 0

End Function

'Utilzado por operação CETIP
Public Function fgExisteCodigoOperacaoCETIP(ByVal plngCodigoOperacaoCETIP As Long, _
                                            ByVal plngTipoMensagem As Long, _
                                            ByVal plngTipoSolicitacao As Long) As Boolean

On Error GoTo ErrorHandler

    Select Case plngTipoMensagem
        Case enumTipoMensagemLQS.MovimentacoesInstFinancCETIP
            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0011.DepositoEmissaoAplicacaoFinanceiro Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0011.DepositoEmissaoSemFinaceiro
                                              
            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then
            
                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoDepositoEmissaoComFinanceiro Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoDepositoEmissaoSemFinanceiro
                                              
            End If
            
        Case enumTipoMensagemLQS.ResgateFundoInvestimentoCETIP
            
            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then
               
               fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0012.LancamentoFinanceiroResgate
               
            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then
            
                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoLancamentoFinanceiroResgate
            
            End If
            
        Case enumTipoMensagemLQS.MovimentacoesCustodiaCETIP
        
            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then
               
                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0011.RetiradaCustodia Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0013.Caucao Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0013.DesvinculacaoCaucao Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0013.VinculacaoReservaTecnica Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0013.AntecipacaoDesvinculacaoReservaTecnica Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0013.VinculacaoDesvinculacaoContaMargem Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0021.TransferenciaCustodiaCamaraNegociacao Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0021.TransferenciaCDPParaINSSSemFinanceiro Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0021.MovimentacaoGarantiaCamaras Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0021.TransferenciaGarantiasContaMargemCamara Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0021.TransferenciaCustodiaEntreContasCliente

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then
                
                
            fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoDepositoEmissaoSemFinanceiro Or _
                                          plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoDepositoEmissaoComFinanceiro Or _
                                          plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoRetiradaCustodia Or _
                                          plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoLancamentoFinanceiroResgate Or _
                                          plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoEspecificacaoLancamentoFinancAplic Or _
                                          plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoEspecificaoLancamentoFinancResgate Or _
                                          plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoCompraVendaDefinitiva Or _
                                          plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoRetencaoIR Or _
                                          plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoVinculacaoReservaTecnica Or _
                                          plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoAntecipacaoDesvinculacaoReservaTecnica
    

            
            
            End If
            
        Case enumTipoMensagemLQS.ExercicioDesistenciaCETIP

            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then
               
                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0012.ExercicioDireitoVendaEmissor Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0012.ExercicioDireitaNaoRepactuacao Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0012.DesistenciaExercicioDireitoNaoRepactuacao Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0012.DesistenciaExercicioDireitoVendaEmissor

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then
                
                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoLancamentoPUFatorContrato

            End If
                                          
        Case enumTipoMensagemLQS.ConversaoPermutaValorImobCETIP

            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0017.RetiradaCustodiaConversaoAcoes

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoRetiradaCustodiaConversaoAcoes

            End If

        Case enumTipoMensagemLQS.EspecificacaoQuantidadesCotasCETIP

            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0022.EspecificacaoLancamentoFinancAplic Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0022.EspecificacaoLancamentoFinancResgate

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoEspecificaoLancamentoFinancResgate Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoEspecificacaoLancamentoFinancAplic

            End If

        Case enumTipoMensagemLQS.OperacaoDefinitivaCETIP

            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then

               fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0052.CompraVendaDefinitiva Or _
                                             plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0052.RegistroRetroativaCompraVendaDefinitiva Or _
                                             plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0052.DistribuicaoPrimariaValorImobiliario Or _
                                             plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0052.ResgateAntecipado Or _
                                             plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0052.IntermediacaoEmissaoSistemaCETIP

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoDistribuicaoValoresMobiliarios Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoResgateAntecipado Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoIntermediacaoEmissaoSistemaCETIP Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoCompraVendaDefinitiva Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoPostecipadoCompraVendaDefinitiva

            End If

        Case enumTipoMensagemLQS.OperacaoCompromissadaCETIP

            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
               plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoPorLastro Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Alteracao Then

               fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0054.CompraVendaCompromissadaRentaPosFix Or _
                                             plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0054.CompraVendaCompromissadaRentaPreFix Or _
                                             plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0054.CessaoRetrocessao Or _
                                             plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0054.RegistroRetroativoCompraVendaComprRentaPreFix Or _
                                             plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0054.RegistroRetroativoCompraVendaComprRentaPosFix Or _
                                             plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0054.RegistroDocumCompraVendaComprRentaPreFix

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoCompraVendaComprRentPosFix Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoCompraVendaComprRentPreFix Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoPagamentoIntermedContrato Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoPostecipadoCompraVendaComprRentPreFix Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoPostecipadoCompraVendaComprRentPosFix

            End If

        Case enumTipoMensagemLQS.OperacaoRetornoAntecipacaoCETIP

            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Alteracao Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0055.AntecipacaoCompraVendaCompromissada Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0055.AntecipacaoCessaoRetrocessao Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0055.RegistroRetroativoAntecipacaoCompraVendaCompr Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0056.RetornoCompraVendaComprRentPreFix Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0056.RetornoCompraVendaComprRentPosFix

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoAntecipacaoCompraVendaCompr Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoPuVoltaOperRetnCompVendComprPosFix Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoRegistroIntermedContratoSwap

            End If

        Case enumTipoMensagemLQS.OperacaoRetencaoIRF_CETIP

            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0064.RetencaoIR

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoRetencaoIR

            End If

        Case enumTipoMensagemLQS.RegistroContratoSWAP, enumTipoMensagemLQS.RegistroContratoSWAPCetip21

            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9001.RegistroContratoSemModalidade Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9001.RegistroContratoComModalidade Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9001.RegistroRetroativoContratoComModalidade Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9001.RegistroRetroativoContratoSemModalidade

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Or _
                   plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoDepositoEmissaoSemFinanceiro Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoDepositoEmissaoComFinanceiro

            End If

        Case enumTipoMensagemLQS.RegistroOperacaoesCETIP

            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9002.RegistroContratoComModalidade Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9002.RegistroRetroativoContratoComModalidade

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoDepositoEmissaoComFinanceiro

            End If

        Case enumTipoMensagemLQS.RegistroContratoTermoCETIP

            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9003.RegistroContratoSemModalidade Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9003.RegistroRetroativoContratoSemModalidade

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoDepositoEmissaoSemFinanceiro

            End If

        Case enumTipoMensagemLQS.ExercicioOpcaoContratoSwapCETIP

            fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9010.ExercicioOpcaoSWAP

        Case enumTipoMensagemLQS.AntecipacaoResgateContratoDerivativoCETIP

            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Alteracao Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9014.ResgateAntecipado Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9014.RegistroRetroativoResgateAntecipado

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoResgateAntecipado

            End If

        Case enumTipoMensagemLQS.LancamentoPU_CETIP

            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9015.ExercicioDireitoVendaEmissor

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoLancamentoPUFatorContrato

            End If

        Case enumTipoMensagemLQS.MovimentacoesContratoDerivativo

            If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9020.CessaoContrato Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9021.ConfirmacaoCessaoContrato Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9082.PagamentoIntermediacaoContrato Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP9082.RegistroIntermediacaoContratoSwap

            ElseIf plngTipoSolicitacao = enumTipoSolicitacao.CancelamentoComMensagem Then

                fgExisteCodigoOperacaoCETIP = plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoCessaoContrato Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoPagamentoIntermedContrato Or _
                                              plngCodigoOperacaoCETIP = enumOperacaoCETIP_CTP0100.EstornoRegistroIntermedContratoSwap

            End If

        Case Else
            fgExisteCodigoOperacaoCETIP = False

    End Select

Exit Function
ErrorHandler:

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgExisteCodigoOperacaoCETIP", 0


End Function

Public Function fgConsisteMensagemA8(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim datDataVigencia                         As Date
Dim datLimiteFech                           As Date
Dim lngCodigoEmpresaFusi                    As Long
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim lngTipoMensagem                         As Long
Dim lngTipoSolicitacao                      As Long
Dim lngFormaLiquidacao                      As Long
Dim lngTipoRedesconto                       As Long
Dim lngTemp                                 As Long
Dim datTemp                                 As Date
Dim lngTipoVinculo                          As Long
Dim lngTipoPagamentoRDSC                    As Long
Dim blnEntradaManual                        As Boolean
Dim lngTipoLeilao                           As Long
Dim lngTipoPagamentoRedesconto              As Long

Dim strTipoLiquidCompromis                  As String
Dim strTipoLiquidDefin                      As String
Dim strPrecoUnitarioRetorno                 As String
Dim strValorFinanceiroRetorno               As String

'Consiste se o Campo TP_TITL_BMA esta preenchido quando o CO_LOCA_LIQU = BMA (17)
Dim objVL                                   As A6A7A8.clsVeiculoLegal
Dim xmlVL                                   As MSXML2.DOMDocument40
Dim strAux                                  As String

Dim dblPrecoUnitario                        As Double
Dim dblPercentualValor                      As Double

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")

        If Not fgExisteTipoMensagem(.selectSingleNode("TP_MESG").Text, datDataVigencia) Then
            'Tipo de mensagem inválida
            fgAdicionaErro xmlErrosNegocio, 4003
        Else
            lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)
        End If

        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        End If

        If Not fgExisteSistema(.selectSingleNode("SG_SIST_ORIG").Text, glngCodigoEmpresa, datDataVigencia) Then
            'Sistema inexistente
            fgAdicionaErro xmlErrosNegocio, 4004
        End If

        If Not .selectSingleNode("SG_SIST_DEST").Text = "A8" Then
            'Sistema destino incorreto
            fgAdicionaErro xmlErrosNegocio, 4005
        End If
        
        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
        
            If Not blnEntradaManual Then
                If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                               objDomNode.Text) Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If
        
        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) = vbNullString Then
                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                   'Código do local de liquidação inválido
                   fgAdicionaErro xmlErrosNegocio, 4008
                End If
            Else
                If Not fgExisteLocalLiquidacao(glngCodigoEmpresaFusionada, _
                                               objDomNode.Text, _
                                               datDataVigencia) Then
                    'Código do local de liquidação inválido
                    fgAdicionaErro xmlErrosNegocio, 4008
                End If
            End If
        End If
        
        'Tipo Conta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If
        
        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(glngCodigoEmpresaFusionada, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If

        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(glngCodigoEmpresaFusionada, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If
        
        'Código do indexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If
        
        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If lngTipoMensagem = enumEvento.InformacoesCadastraisVeiculoLegal Then
                If Trim$(objDomNode.Text) = vbNullString Then
                        'Veiculo legal inválido
                        fgAdicionaErro xmlErrosNegocio, 4007
                End If
            Else
                If Not Trim$(objDomNode.Text) = vbNullString Then
                    If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                                objDomNode.Text, _
                                                datDataVigencia, _
                                                .selectSingleNode("CO_EMPR").Text) Then
                        'Veiculo legal inválido
                        fgAdicionaErro xmlErrosNegocio, 4007
                    End If
                End If
            End If
        End If
        
        'Indicador de débito/crédito
        Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                lngTemp = Val(objDomNode.Text)
                If lngTemp <> enumTipoDebitoCredito.Credito And _
                   lngTemp <> enumTipoDebitoCredito.Debito Then
                    'Indicador de débito/crédito inválido
                    fgAdicionaErro xmlErrosNegocio, 4021
                End If
            End If
        End If
        
        'Tipo de Back Office
        Set objDomNode = .selectSingleNode("TP_BKOF")
        If Not objDomNode Is Nothing Then
            If Not fgExisteBackOffice(objDomNode.Text, datDataVigencia) Then
                'Tipo Back Office inexistente
                fgAdicionaErro xmlErrosNegocio, 4058
            End If
        End If
        
        'Código do Produto
        'KIDA - 30/06/2009
        'RATS - 916
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            Else
                'KIDA - 24/11/2009
                'RATS - 953
                If lngTipoMensagem <> enumTipoMensagemBUS.VinculoDesvinculoTransferencia And glngTipoBackOffice <> enumTipoBackOffice.Tesouraria Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            End If
        Else
            'A maioria dos layouts CAM nao tem CD_PROD
            Set objDomNode = .selectSingleNode("TP_MESG")
            If Not objDomNode Is Nothing Then
                If objDomNode.Text <> enumTipoMensagemBUS.ContratacaoMercadoPrimario And objDomNode.Text <> enumTipoMensagemBUS.EdicaoContratacaoMercadoPrimario _
                And objDomNode.Text <> enumTipoMensagemBUS.ConfirmacaoEdicaoContratacaoMercadoPrimario And objDomNode.Text <> enumTipoMensagemBUS.AlteracaoContrato _
                And objDomNode.Text <> enumTipoMensagemBUS.EdicaoAlteracaoContrato And objDomNode.Text <> enumTipoMensagemBUS.ConfirmacaoEdicaoAlteracaoContrato _
                And objDomNode.Text <> enumTipoMensagemBUS.LiquidacaoMercadoPrimario And objDomNode.Text <> enumTipoMensagemBUS.BaixaValorLiquidar _
                And objDomNode.Text <> enumTipoMensagemBUS.RestabelecimentoBaixa And objDomNode.Text <> enumTipoMensagemBUS.CancelamentoValorLiquidar _
                And objDomNode.Text <> enumTipoMensagemBUS.EdicaoCancelamentoValorLiquidar And objDomNode.Text <> enumTipoMensagemBUS.ConfirmacaoEdicaoCancelamentoValorLiquidar _
                And objDomNode.Text <> enumTipoMensagemBUS.VinculacaoContratos And objDomNode.Text <> enumTipoMensagemBUS.AnulacaoEvento _
                And objDomNode.Text <> enumTipoMensagemBUS.CorretoraRequisitaClausulasEspecificas And objDomNode.Text <> enumTipoMensagemBUS.IFInformaClausulasEspecificas _
                And objDomNode.Text <> enumTipoMensagemBUS.ManutencaoCadastroAgenciaCentralizadoraCambio And objDomNode.Text <> enumTipoMensagemBUS.CredenciamentoDescredenciamentoDispostoRMCCI _
                And objDomNode.Text <> enumTipoMensagemBUS.IncorporacaoContratos And objDomNode.Text <> enumTipoMensagemBUS.AceiteRejeicaoIncorporacaoContratos _
                And objDomNode.Text <> enumTipoMensagemBUS.ConsultaContratosEmSer And objDomNode.Text <> enumTipoMensagemBUS.ConsultaEventosUmDia _
                And objDomNode.Text <> enumTipoMensagemBUS.ConsultaDetalhamentoContratoInterbancario And objDomNode.Text <> enumTipoMensagemBUS.ConsultaEventosContratoMercadoPrimario _
                And objDomNode.Text <> enumTipoMensagemBUS.ConsultaEventosContratoIntermediadoMercadoPrimario And objDomNode.Text <> enumTipoMensagemBUS.ConsultaHistoricoIncorporacoes _
                And objDomNode.Text <> enumTipoMensagemBUS.ConsultaContratosIncorporacao And objDomNode.Text <> enumTipoMensagemBUS.ConsultaCadeiaIncorporacoesContrato _
                And objDomNode.Text <> enumTipoMensagemBUS.ConsultaPosicaoCambioMoeda And objDomNode.Text <> enumTipoMensagemBUS.AtualizaçãoInclusãoInstrucoesPagamento _
                And objDomNode.Text <> enumTipoMensagemBUS.ConsultaInstrucoesPagamento And objDomNode.Text <> enumTipoMensagemBUS.RegistroOperacaoArbitragem _
                And objDomNode.Text <> enumTipoMensagemBUS.IFCamaraConsultaContratosCambioMercadoInterbancario _
                And objDomNode.Text <> enumTipoMensagemBUS.IFInformaTIRemContrapartidaaRagadorouRecebedorPaís And objDomNode.Text <> enumTipoMensagemBUS.IFInformaTIRemContrapartidaOutraCDE _
                And objDomNode.Text <> enumTipoMensagemBUS.IFInformaTIRemContrapartidaOperacaoCambialPropria And objDomNode.Text <> enumTipoMensagemBUS.IFRequisitaInclusaoemCadastroCDE _
                And objDomNode.Text <> enumTipoMensagemBUS.IFRequisitaAlteracaoCadastroCDE And objDomNode.Text <> enumTipoMensagemBUS.IFRequisitaExclusaoCadastroCDE _
                And objDomNode.Text <> enumTipoMensagemBUS.IFInformaAnulacaoRegistroTIR And objDomNode.Text <> enumTipoMensagemBUS.IFConsultaCDE _
                And objDomNode.Text <> enumTipoMensagemBUS.IFConsultaTIRUmDia And objDomNode.Text <> enumTipoMensagemBUS.IFConsultaDetalhamentoTIR Then
                               
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            End If
        End If
        
        'Código do Produto PJ Câmara
        Set objDomNode = .selectSingleNode("CO_PROD_CAMR")
        If Not objDomNode Is Nothing Then
            If glngTipoBackOffice = enumTipoBackOffice.Tesouraria And _
               lngTipoMensagem = enumTipoMensagemLQS.VinculoDesvinculoTransf And _
               lngTipoVinculo = enumTipoVinculo.TransferenciaCamaraLDL And _
               Val("0" & objDomNode.Text) = 0 Then
                    'Produto PJ Câmara inválido
                    fgAdicionaErro xmlErrosNegocio, 4147
                'End If
            Else
                If Val("0" & objDomNode.Text) <> 0 Then
                    If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                           objDomNode.Text, _
                                           datDataVigencia) Then
                        'Produto PJ Câmara inválido
                        fgAdicionaErro xmlErrosNegocio, 4147
                    End If
                End If
            End If
        Else
            'KIDA - Book 30.1
            
            If glngTipoBackOffice = enumTipoBackOffice.Tesouraria And _
               lngTipoMensagem = enumTipoMensagemLQS.VinculoDesvinculoTransf Then
                'Produto PJ Câmara inválido
                fgAdicionaErro xmlErrosNegocio, 4147
            End If
        End If

        'Código da conta Selic
        Set objDomNode = .selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) = 0 And _
               lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                'Código da conta Selic obrigatório
                fgAdicionaErro xmlErrosNegocio, 4140
            End If
        End If

        'Código da conta Selic da contraparte
        Set objDomNode = .selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) = 0 And _
               lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                'Código da conta Selic da contraparte obrigatório
                fgAdicionaErro xmlErrosNegocio, 4141
            End If
        End If

        'Forma de Liquidação
        lngFormaLiquidacao = 0
        Set objDomNode = .selectSingleNode("CO_FORM_LIQU")
        If Not objDomNode Is Nothing Then
            lngTemp = Val("0" & objDomNode.Text)
            If lngTemp <> enumFormaLiquidacao.ContaCorrente And _
               lngTemp <> enumFormaLiquidacao.Contabil Then
                'Forma de Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4026
            Else
                lngFormaLiquidacao = lngTemp
            End If
        End If

        'Data Operação
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            datTemp = fgDtXML_To_Date(objDomNode.Text)
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da operação inválida
                fgAdicionaErro xmlErrosNegocio, 4024
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                    'Data da operação não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4059
                Else
                    If datTemp < flDataHoraServidor(enumFormatoDataHora.Data) Then
                        If lngTipoMensagem = enumTipoMensagemLQS.Definitiva Then
                            If datTemp < fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 1, enumPaginacao.Anterior) Then
                                If datTemp = fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 2, enumPaginacao.Anterior) Then
                                    If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente Then
                                        '4211 - Forma de Liquidação inválida para data retroativa
                                        fgAdicionaErro xmlErrosNegocio, 4211
                                    End If
                                End If
                                If Not .selectSingleNode("TP_CNPT") Is Nothing Then
                                    If Trim(.selectSingleNode("TP_CNPT").Text) <> "" Then
                                        If CLng(.selectSingleNode("TP_CNPT").Text) = enumTipoContraparte.Cliente1 Or _
                                            CLng(.selectSingleNode("TP_CNPT").Text) = enumTipoContraparte.Interno Then
                                            If datTemp < fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 2, enumPaginacao.Anterior) Then
                                                'Data da operação inválida
                                                fgAdicionaErro xmlErrosNegocio, 4024
                                            End If
                                        Else
                                            If datTemp <= fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 2, enumPaginacao.Anterior) Then
                                                'Data da operação inválida
                                                fgAdicionaErro xmlErrosNegocio, 4024
                                            End If
                                        End If
                                    Else
                                        'Data da operação inválida
                                        fgAdicionaErro xmlErrosNegocio, 4024
                                    End If
                                Else
                                    'Data da operação inválida
                                    fgAdicionaErro xmlErrosNegocio, 4024
                                End If
                            End If
                        ElseIf lngTipoMensagem = enumTipoMensagemLQS.Compromissada Then
                            If datTemp < fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 1, enumPaginacao.Anterior) Then
                                'Data da operação inválida
                                fgAdicionaErro xmlErrosNegocio, 4024
                            End If
                        Else
                            'Data da operação inválida
                            fgAdicionaErro xmlErrosNegocio, 4024
                        End If
                    End If
                End If
            End If
        End If

        If datTemp <> flDataHoraServidor(enumFormatoDataHora.Data) And _
          (lngTipoMensagem = enumTipoMensagemLQS.Redesconto Or _
           lngTipoMensagem = enumTipoMensagemLQS.PgtoRedesconto Or _
           lngTipoMensagem = enumTipoMensagemLQS.ConversaoRedesconto Or _
           lngTipoMensagem = enumTipoMensagemLQS.DespesasSelic Or _
           lngTipoMensagem = enumTipoMensagemLQS.EventosSelic) Then
            
            'Data da operação inválida
            fgAdicionaErro xmlErrosNegocio, 4024
        
        ElseIf datTemp > flDataHoraServidor(enumFormatoDataHora.Data) And _
              (lngTipoMensagem = enumTipoMensagemLQS.Definitiva Or _
               lngTipoMensagem = enumTipoMensagemLQS.Compromissada) Then
            
            'Data da operação inválida
            fgAdicionaErro xmlErrosNegocio, 4024
            
        'Projeto Sevilha - Adrian 21/06/2006
        ElseIf datTemp > flDataHoraServidor(enumFormatoDataHora.Data) And _
              lngTipoMensagem = enumTipoMensagemLQS.VinculoDesvinculoTransf Then
            
            If datTemp >= fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 2, enumPaginacao.proximo) Then
                'Data da operação inválida
                fgAdicionaErro xmlErrosNegocio, 4024
            End If
            
        End If

        'Pikachu
        'Esta condicao de compilação foi colocada para atender a versão de Homologação/Produção
        'Sem Conta Corrente
        #If ValidaCC = 1 Then
            
            'Código do Banco
            Set objDomNode = .selectSingleNode("CO_BANC")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                  (lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                   lngTipoSolicitacao = enumTipoSolicitacao.Alteracao) Then
                    If Not fgExisteCodigoBanco(Val(objDomNode.Text)) Then
                        'Código do Banco inválido
                        fgAdicionaErro xmlErrosNegocio, 4027
                    End If
                ElseIf lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                      (lngTipoMensagem = enumTipoMensagemLQS.EventosSelic Or _
                       lngTipoMensagem = enumTipoMensagemLQS.DespesasSelic) Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código do Banco inválido
                        fgAdicionaErro xmlErrosNegocio, 4027
                    'Projeto Sevilha - Adrian 21/06/2006
                    'ElseIf lngCodBanco <> 0 Then
                    '    If Val("0" & objDomNode.Text) <> lngCodBanco Then
                    '        'Código do Banco inválido
                    '        fgAdicionaErro xmlErrosNegocio, 4027
                    '    End If
                    End If
                End If
            End If
    
            'Código da Agência
            Set objDomNode = .selectSingleNode("CO_AGEN")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                    lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da Agência inválido
                        fgAdicionaErro xmlErrosNegocio, 4028
                    End If
                ElseIf lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                      (lngTipoMensagem = enumTipoMensagemLQS.EventosSelic Or _
                       lngTipoMensagem = enumTipoMensagemLQS.DespesasSelic) Then
                        If Val("0" & objDomNode.Text) = 0 Then
                            'Código da Agência inválido
                            fgAdicionaErro xmlErrosNegocio, 4028
                        End If
                End If
            End If
    
            'Número da conta corrente
            Set objDomNode = .selectSingleNode("NU_CC")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                    lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4029
                    End If
                ElseIf lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                      (lngTipoMensagem = enumTipoMensagemLQS.EventosSelic Or _
                       lngTipoMensagem = enumTipoMensagemLQS.DespesasSelic) Then
                        If Val("0" & objDomNode.Text) = 0 Then
                            'Código da conta corrente inválido
                            fgAdicionaErro xmlErrosNegocio, 4029
                        End If
                End If
            End If
    
            'Valor Lançamento Conta Corrente
            Set objDomNode = .selectSingleNode("VA_LANC_CC")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                    lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    If fgVlrXml_To_Decimal(objDomNode.Text) = 0 Then
                        'Valor do lançamento na conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4030
                    Else
                        If fgVlrXml_To_Decimal(objDomNode.Text) > fgVlrXml_To_Decimal(.selectSingleNode("//VA_OPER_ATIV").Text) Then
                            'Valor do lançamento na conta corrente inválido
                            fgAdicionaErro xmlErrosNegocio, 4030
                        End If
                    End If
                ElseIf lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                      (lngTipoMensagem = enumTipoMensagemLQS.EventosSelic Or _
                       lngTipoMensagem = enumTipoMensagemLQS.DespesasSelic Or _
                       lngTipoMensagem = enumTipoMensagemLQS.TermoDataLiquidacao) Then
                        If fgVlrXml_To_Decimal(objDomNode.Text) = 0 Then
                            'Valor do lançamento na conta corrente inválido
                            fgAdicionaErro xmlErrosNegocio, 4030
                        Else
                            If fgVlrXml_To_Decimal(objDomNode.Text) > fgVlrXml_To_Decimal(.selectSingleNode("//VA_OPER_ATIV").Text) Then
                                'Valor do lançamento na conta corrente inválido
                                fgAdicionaErro xmlErrosNegocio, 4030
                            End If
                       End If
                End If
            End If


        #End If

        'Tipo de compromisso
        Set objDomNode = .selectSingleNode("TP_CPRO_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If lngTipoMensagem = enumTipoMensagemLQS.Compromissada And _
               Trim$(objDomNode.Text) = vbNullString Then
                'Tipo de compromisso inválido
                fgAdicionaErro xmlErrosNegocio, 4031
            End If
            If Trim$(objDomNode.Text) <> vbNullString Then
                If Not fgExisteTipoCompromisso(objDomNode.Text) Then
                    'Tipo de compromisso inválido
                    fgAdicionaErro xmlErrosNegocio, 4031
                End If
            End If
        End If

        'Tipo de compromisso retorno
        Set objDomNode = .selectSingleNode("TP_CPRO_RETN_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text <> vbNullString Then
                If Not fgExisteTipoCompromissoRetn(objDomNode.Text) Then
                    'Tipo de compromisso de retorno inválido
                    fgAdicionaErro xmlErrosNegocio, 4032
                End If
            Else
                If lngTipoMensagem = enumTipoMensagemLQS.RetornoCompromissada Then
                    'Tipo de compromisso de retorno é obrigatório para Tipo de mensagem Retorno Compromissada
                    fgAdicionaErro xmlErrosNegocio, 4033
                End If
            End If
        End If
        
        'Tipo Liquidação
        Set objDomNode = .selectSingleNode("TP_LIQU")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text <> vbNullString Then
                If Not fgExisteTipoLiquidacaoMensageria(objDomNode.Text) Then
                    'Tipo Liquidação do Mensageria inválido
                    fgAdicionaErro xmlErrosNegocio, 4034
                End If
            End If
        End If
        
        'Tipo de leilão
        Set objDomNode = .selectSingleNode("TP_LEIL")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text <> vbNullString Then
                If Not fgExisteTipoLeilao(objDomNode.Text) Then
                    'Tipo de leilão inválido
                    fgAdicionaErro xmlErrosNegocio, 4035
                Else
                    lngTipoLeilao = Val("0" & objDomNode.Text)
                End If
            End If
        End If

        'Número Operação Selic
        Set objDomNode = .selectSingleNode("NU_COMD_OPER")
        If Not objDomNode Is Nothing Then
            If Len(objDomNode.Text) > 6 Or _
               (Not IsNumeric(objDomNode.Text) And Trim$(objDomNode.Text) <> vbNullString) Then
                'Número do comando inválido
                fgAdicionaErro xmlErrosNegocio, 4100
            ElseIf (lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                    lngTipoSolicitacao = enumTipoSolicitacao.Alteracao) And _
                   lngTipoMensagem <> enumTipoMensagemLQS.TermoDataLiquidacao And _
                   lngTipoMensagem <> enumTipoMensagemLQS.EventosSelic And _
                   fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                'Número do comando obrigatório
                fgAdicionaErro xmlErrosNegocio, 4023
            End If
        End If


        'Tipo Vínculo
        If lngTipoMensagem = enumTipoMensagemLQS.VinculoDesvinculoTransf Then
            Set objDomNode = .selectSingleNode("TP_VINC_DSVN_TRAF")
            If Not objDomNode Is Nothing Then
                If Not objDomNode.Text = vbNullString Then
                    lngTipoVinculo = Val("0" & objDomNode.Text)
                    If Not fgExisteTipoVinculo(lngTipoVinculo) Then
                        'Tipo de Vínculo/Desvínculo/Transferência inexistente
                        fgAdicionaErro xmlErrosNegocio, 4036
                    End If
                End If
            End If
    
            'Tipo Transferência
            Set objDomNode = .selectSingleNode("TP_TRAF")
            If Not objDomNode Is Nothing Then
                If lngTipoVinculo = enumTipoVinculo.TransferenciaSemMovFinanceira And _
                    objDomNode.Text = vbNullString Then
                    'Tipo de Transferência obrigatório para Transferência sem movimentação financeira
                    fgAdicionaErro xmlErrosNegocio, 4037
                ElseIf objDomNode.Text <> vbNullString Then
                    If Not fgExisteTipoTransferencia(Val("0" & objDomNode.Text)) Then
                        'Tipo de Transferência inválido
                        fgAdicionaErro xmlErrosNegocio, 4038
                    End If
                End If
            End If
            
            'Tipo Transferência LDL
            Set objDomNode = .selectSingleNode("TP_TRAF_LDL")
            If Not objDomNode Is Nothing Then
                If lngTipoVinculo = enumTipoVinculo.TransferenciaCamaraLDL And Trim(objDomNode.Text) = vbNullString Then
                    'Tipo de Transferência LDL obrigatório para Transferência para Câmara LDL
                    fgAdicionaErro xmlErrosNegocio, 4039
                ElseIf objDomNode.Text <> vbNullString Then
                    If Not fgExisteTipoTransferenciaLDL(Val("0" & objDomNode.Text)) Then
                        'Tipo de Transferência LDL inválido
                        fgAdicionaErro xmlErrosNegocio, 4040
                    End If
                End If
            Else
                If lngTipoVinculo = enumTipoVinculo.TransferenciaCamaraLDL Then
                    'Tipo de Transferência LDL obrigatório para Transferência para Câmara LDL
                    fgAdicionaErro xmlErrosNegocio, 4039
                End If
            End If

            Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
            If Not objDomNode Is Nothing Then
                If fgVlrXml_To_Decimal("0" & objDomNode.Text) <= 0 Then
                    If lngTipoVinculo <> enumTipoVinculo.TransferenciaCamaraLDL And _
                        lngTipoVinculo <> enumTipoVinculo.TransferenciaSemMovFinanceira Then
                        'Valor Financeiro deve ser maior que zero
                        fgAdicionaErro xmlErrosNegocio, 4049
                    End If
                End If
            End If

            If glngTipoBackOffice = enumTipoBackOffice.Tesouraria Then
                Set objDomNode = .selectSingleNode("CO_PROD")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text <> vbNullString And objDomNode.Text <> "0" Then
                        If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                               objDomNode.Text, _
                                               datDataVigencia) Then
                            '4011 - Produto inválido
                            fgAdicionaErro xmlErrosNegocio, 4011
                            fgGravaArquivo "ERRO_PRODUTO", "Produto    :" & .selectSingleNode("CO_EMPR").Text & vbCrLf & _
                                                           "Empresa    :" & objDomNode.Text & vbCrLf & _
                                                           "Data Atual :" & datDataVigencia & vbCrLf
                        End If
                    End If
                End If
            Else
                Set objDomNode = .selectSingleNode("CO_PROD")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text <> vbNullString And _
                        objDomNode.Text <> "0" Then
                        If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                               objDomNode.Text, _
                                               datDataVigencia) Then
                            '4011 - Produto inválido
                            fgAdicionaErro xmlErrosNegocio, 4011
                            fgGravaArquivo "ERRO_PRODUTO", "Produto    :" & .selectSingleNode("CO_EMPR").Text & vbCrLf & _
                                                           "Empresa    :" & objDomNode.Text & vbCrLf & _
                                                           "Data Atual :" & datDataVigencia & vbCrLf
                        End If
                    End If
                End If
            End If
        End If

        'Tipo Operação Rotina Abertura
        Set objDomNode = .selectSingleNode("TP_OPER_ROTI_ABER")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text <> vbNullString Then
                If Not fgExisteOperacaoRotinaAbertura(objDomNode.Text) Then
                    'Tipo de Operação Rotina de Abertura inválido
                    fgAdicionaErro xmlErrosNegocio, 4041
                End If
            End If
        End If

        If lngTipoMensagem = enumTipoMensagemLQS.Redesconto Or _
           lngTipoMensagem = enumTipoMensagemLQS.Compromissada Then

            'Tipo de Redesconto
            Set objDomNode = .selectSingleNode("TP_RDSC")
            If Not objDomNode Is Nothing Then
                If Not fgExisteTipoRedesconto(objDomNode.Text) Then
                    'Tipo de redesconto inválido
                    fgAdicionaErro xmlErrosNegocio, 4042
                Else
                    lngTipoRedesconto = objDomNode.Text
                End If
            End If

            'Local de liquidação associado
            Set objDomNode = .selectSingleNode("CO_LOCA_LIQU_ASSO")
            If Not objDomNode Is Nothing Then
                If objDomNode.Text <> vbNullString Then
                    If Not fgExisteLocalLiquidacao(glngCodigoEmpresaFusionada, _
                                                   objDomNode.Text, _
                                                   datDataVigencia) Then
                        'Tipo de Liquidação Associado inválido
                        fgAdicionaErro xmlErrosNegocio, 4043
                    End If
                End If
            End If

            'Identificador Título
            Set objDomNode = .selectSingleNode("NU_ATIV_MERC")
            If Not objDomNode Is Nothing Then
                If Trim$(objDomNode.Text) = vbNullString Then
                    If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                       'Identificador de Título inválido
                       fgAdicionaErro xmlErrosNegocio, 4074
                    End If
                End If
            End If

            'Descrição Título Selic
            Set objDomNode = .selectSingleNode("DE_ATIV_MERC")
            If Not objDomNode Is Nothing Then
                If Trim$(objDomNode.Text) = vbNullString Then
                    If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                       'Descrição Título Selic inválido
                       fgAdicionaErro xmlErrosNegocio, 4075
                    End If
                End If
            End If

            'Data de Vencimento
            Set objDomNode = .selectSingleNode("DT_VENC_ATIV")
            If Not objDomNode Is Nothing Then
                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    If Not flValidaDataNumerica(objDomNode.Text) Then
                        'Data de Vencimento inválida
                        fgAdicionaErro xmlErrosNegocio, 4063
                    End If
                End If
            End If

            'Quantidade de Títulos
            Set objDomNode = .selectSingleNode("QT_ATIV_MERC")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) = 0 Then
                    If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                       'Quantidade de Títulos deve ser maior que zero
                       fgAdicionaErro xmlErrosNegocio, 4050
                    End If
                End If
            End If

            Set objDomNode = .selectSingleNode("PU_ATIV_MERC")
            If Not objDomNode Is Nothing Then
                If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                    If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                        'Preço Unitário deve ser maior que zero
                        fgAdicionaErro xmlErrosNegocio, 4053
                    End If
                End If
            End If

            If lngTipoRedesconto = enumTipoRedesconto.RedescIntradiaAquisicao Or _
                lngTipoRedesconto = enumTipoRedesconto.RedescIntradiaGarantia Then
                Set objDomNode = .selectSingleNode("CO_CHAV_ASSO_SELIC")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                           'Chave de Associação Selic é obrigatória
                           fgAdicionaErro xmlErrosNegocio, 4025
                        End If
                    End If
                End If
            End If

        End If

        If lngTipoMensagem = enumTipoMensagemLQS.PgtoRedesconto Then

            'Tipo de pagamento redesconto
            Set objDomNode = .selectSingleNode("TP_PGTO_RDSC")
            If Not objDomNode Is Nothing Then
                If objDomNode.Text <> vbNullString Then
                    If Not fgExisteTipoPagamentoRedesc(Val("0" & objDomNode.Text)) Then
                        'Tipo de pagamento redesconto inválido
                        fgAdicionaErro xmlErrosNegocio, 4044
                    Else
                        lngTipoPagamentoRDSC = CLng(objDomNode.Text)
                    End If
                End If
            End If

            'Tipo de pagamento
            Set objDomNode = .selectSingleNode("TP_PGTO")
            If Not objDomNode Is Nothing Then
                If objDomNode.Text <> vbNullString Then
                    If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                        If Not fgExisteTipoPagamento(objDomNode.Text) Then
                            'Tipo de pagamento inválido
                            fgAdicionaErro xmlErrosNegocio, 4045
                        End If
                    End If
                End If
            End If

            'Preço Unitário Retorno
            Set objDomNode = .selectSingleNode("VA_PU_RETN")
            If Not objDomNode Is Nothing Then
                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    If objDomNode.Text <> vbNullString Then
                        If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                            'Preço Unitário de Retorno deve ser maior que zero
                            fgAdicionaErro xmlErrosNegocio, 4056
                        End If
                    End If
                End If
            End If

            'Quantidade de Títulos
            Set objDomNode = .selectSingleNode("QT_ATIV_MERC")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) = 0 Then
                    If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                       'Quantidade de Títulos deve ser maior que zero
                       fgAdicionaErro xmlErrosNegocio, 4050
                    End If
                End If
            End If

            'Chave associação Selic
            If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                Set objDomNode = .selectSingleNode("CO_CHAV_ASSO_SELIC")
                If Not objDomNode Is Nothing Then
                    If lngTipoPagamentoRDSC = enumTipoPagtoRedesconto.PagamentoRedescontoAssocAVenda Then
                        If Val("0" & objDomNode.Text) = 0 Then
                            'Chave de Associação Selic é obrigatória
                            fgAdicionaErro xmlErrosNegocio, 4025
                        End If
                    End If
                Else
                    'Chave de Associação Selic é obrigatória
                    fgAdicionaErro xmlErrosNegocio, 4025
                End If
            End If
        End If

        If lngTipoMensagem = enumTipoMensagemLQS.ConversaoRedesconto Then
            'Preço Unitário de Conversão ou Recontratação
            Set objDomNode = .selectSingleNode("VA_PU_CVER_RCNT")
            If Not objDomNode Is Nothing Then
                If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                    'Preço Unitário de Conversão ou Recontratação deve ser  maior que zero
                    fgAdicionaErro xmlErrosNegocio, 4046
                End If
            End If

            'Quantidade Título conversão ou recontratação
            Set objDomNode = .selectSingleNode("QT_TITU_CVER_RCNT")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) <= 0 Then
                    'Quantidade Título conversão ou recontratação deve ser maior que zero
                    fgAdicionaErro xmlErrosNegocio, 4047
                End If
            End If
            
            'Valor Financeiro Conversão ou recontratação
            Set objDomNode = .selectSingleNode("VA_OPER_ATIV_RETN")
            If Not objDomNode Is Nothing Then
                If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 Then
                    'Valor Financeiro Conversão ou recontratação deve ser maior que zero
                    fgAdicionaErro xmlErrosNegocio, 4048
                End If
            End If
            
        End If

        'Valor Financeiro
        If lngTipoMensagem <> enumTipoMensagemLQS.VinculoDesvinculoTransf Then
            Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
            If Not objDomNode Is Nothing Then
                If fgVlrXml_To_Decimal("0" & objDomNode.Text) <= 0 Then
                    If Not ((lngTipoMensagem = enumTipoMensagemLQS.TermoD0 Or _
                            lngTipoMensagem = enumTipoMensagemLQS.EventosSelic Or _
                            lngTipoMensagem = enumTipoMensagemLQS.DespesasSelic) And _
                            fgVlrXml_To_Decimal(objDomNode.Text) = 0) Then
                        'Valor Financeiro deve ser maior que zero
                        fgAdicionaErro xmlErrosNegocio, 4049
                    End If
                End If
            End If
        End If

        'Quantidade de Títulos
        If lngTipoMensagem <> enumTipoMensagemLQS.Redesconto And _
           lngTipoMensagem <> enumTipoMensagemLQS.Compromissada Then
            Set objDomNode = .selectSingleNode("QT_ATIV_MERC")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) <= 0 Then
                    'Quantidade de Títulos deve ser maior que zero
                    fgAdicionaErro xmlErrosNegocio, 4050
                End If
            End If
        End If

        'Valor Financeiro Retorno
        Set objDomNode = .selectSingleNode("VA_OPER_ATIV_RETN")
        If Not objDomNode Is Nothing Then
            If lngTipoMensagem = enumTipoMensagemLQS.Compromissada Then
                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao And _
                   Trim$(objDomNode.Text) = vbNullString Then
                    'Valor Financeiro Retorno deve ser maior que zero
                    fgAdicionaErro xmlErrosNegocio, 4051
                End If
            End If
            If lngTipoMensagem <> enumTipoMensagemLQS.TermoD0 Then
                If objDomNode.Text <> vbNullString Then
                    If fgVlrXml_To_Decimal(objDomNode.Text) < 0 Then
                        'Valor Financeiro Retorno deve ser maior que zero
                        fgAdicionaErro xmlErrosNegocio, 4051
                    End If
                End If
            End If
        End If

        'Número Controle RDC Original
        Set objDomNode = .selectSingleNode("NU_CTRL_RDSC_ORIG")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text = "0" Or _
                Trim(objDomNode.Text) = "" Then
                'Número Controle RDC Original inválido
                fgAdicionaErro xmlErrosNegocio, 4052
            End If
        End If

        'Preço Unitário
        dblPrecoUnitario = 0
        dblPercentualValor = 0
        
        If lngTipoMensagem = enumTipoMensagemLQS.TermoD0 Then
            Set objDomNode = .selectSingleNode("PU_ATIV_MERC")
            If Not objDomNode Is Nothing Then
                dblPrecoUnitario = fgVlrXml_To_Decimal(objDomNode.Text)
            End If
        
            Set objDomNode = .selectSingleNode("PE_VALO_PAR")
            If Not objDomNode Is Nothing Then
                dblPercentualValor = fgVlrXml_To_Decimal(objDomNode.Text)
            End If
            
            If (dblPrecoUnitario <> 0 And dblPercentualValor <> 0) Or _
               (dblPrecoUnitario = 0 And dblPercentualValor = 0) Then
                'Apenas Preço Unitáro ou Percentual Valor Par deve ser preenchido
                fgAdicionaErro xmlErrosNegocio, 4054
                
            ElseIf dblPrecoUnitario <= 0 And dblPercentualValor = 0 Then
                'Preço Unitário deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4053
            
            End If
        
            Set objDomNode = .selectSingleNode("TP_LIQU_COMP")
            If Not objDomNode Is Nothing Then
                strTipoLiquidCompromis = Trim$(objDomNode.Text)
            End If
        
            Set objDomNode = .selectSingleNode("TP_LIQU_DEFN")
            If Not objDomNode Is Nothing Then
                strTipoLiquidDefin = Trim$(objDomNode.Text)
            End If
        
            Set objDomNode = .selectSingleNode("PU_ATIV_MERC_RETN")
            If Not objDomNode Is Nothing Then
                strPrecoUnitarioRetorno = Trim$(objDomNode.Text)
            End If
        
            Set objDomNode = .selectSingleNode("VA_OPER_ATIV_RETN")
            If Not objDomNode Is Nothing Then
                strValorFinanceiroRetorno = Trim$(objDomNode.Text)
            End If
        
            If strTipoLiquidCompromis = vbNullString And strTipoLiquidDefin = vbNullString Then
                'Tipo de Liquidação Compromissada e Tipo de Liquidação Definitiva em branco. O preenchimento de pelo menos um dos campos é obrigatório.
                fgAdicionaErro xmlErrosNegocio, 4267
        
            ElseIf strTipoLiquidCompromis <> vbNullString And strTipoLiquidDefin <> vbNullString Then
                'Tipo de Liquidação Compromissada e Tipo de Liquidação Definitiva preenchidos simultaneamente. Apenas um dos dois campos pode ser preenchido.
                fgAdicionaErro xmlErrosNegocio, 4268
        
            ElseIf strTipoLiquidCompromis <> vbNullString Then
                Select Case strTipoLiquidCompromis
                    Case "01", "02", "03", "04", "05", "06", "07", "08"
                        If fgVlrXml_To_Decimal(strPrecoUnitarioRetorno) <= 0 Then
                            'Para Operação Compromissada a Termo D0, o Preço Unitário Retorno deve ser maior que 0 (zero).
                            fgAdicionaErro xmlErrosNegocio, 4271
                        End If
                        
                        If fgVlrXml_To_Decimal(strValorFinanceiroRetorno) <= 0 Then
                            'Para Operação Compromissada a Termo D0, o Valor Financeiro Retorno deve ser maior que 0 (zero).
                            fgAdicionaErro xmlErrosNegocio, 4272
                        End If
                        
                    Case Else
                        'Tipo de Liquidação Compromissada inválido.
                        fgAdicionaErro xmlErrosNegocio, 4270
                        
                End Select
        
            ElseIf strTipoLiquidDefin <> vbNullString Then
                Select Case strTipoLiquidDefin
                    Case "01", "02"
                    Case Else
                        'Tipo de Liquidação Definitiva inválido.
                        fgAdicionaErro xmlErrosNegocio, 4269
                        
                End Select
        
            End If
        
        ElseIf lngTipoMensagem <> enumTipoMensagemLQS.Redesconto And _
               lngTipoMensagem <> enumTipoMensagemLQS.Compromissada Then
               
            Set objDomNode = .selectSingleNode("PU_ATIV_MERC")
            If Not objDomNode Is Nothing Then
                If fgVlrXml_To_Decimal(objDomNode.Text) <= 0 And _
                   lngTipoMensagem <> enumTipoMensagemLQS.TermoD0 And _
                   (lngTipoMensagem <> enumTipoMensagemLQS.VinculoDesvinculoTransf Or _
                    lngTipoMensagem = enumTipoMensagemLQS.VinculoDesvinculoTransf And _
                    lngTipoVinculo <> enumTipoVinculo.Vinculo And enumTipoVinculo.Desvinculo <> 2) Then
                    'Preço Unitário deve ser maior que zero
                    fgAdicionaErro xmlErrosNegocio, 4053
                End If
            End If
        
        End If

        'Percentual Valor Par
        Set objDomNode = .selectSingleNode("PE_VALO_PAR")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal(objDomNode.Text) < 0 Then
                'Percentual valor par deve ser maior ou igual a zero
                fgAdicionaErro xmlErrosNegocio, 4055
            End If
        End If
        
        'Data de envio
        Set objDomNode = .selectSingleNode("DT_REME")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            End If
        End If
        
        'Hora de envio
        Set objDomNode = .selectSingleNode("HO_REME")
        If Not objDomNode Is Nothing Then
            If Not flValidaHoraNumerica(objDomNode.Text) Then
                'Hora da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4061
            End If
        End If
        
        'Hora de agendamento
        Set objDomNode = .selectSingleNode("HO_AGND")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 And _
            Not flValidaHoraNumerica(objDomNode.Text) Then
                'Horário de Agendamento inválido
                fgAdicionaErro xmlErrosNegocio, 4062
            End If
        End If

        'Data de vencimento
        Set objDomNode = .selectSingleNode("DT_VENC_ATIV")
        If Not objDomNode Is Nothing Then
            If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao And _
              (lngTipoMensagem = enumTipoMensagemLQS.Compromissada Or _
               lngTipoMensagem = enumTipoMensagemLQS.Redesconto) Then
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data de vencimento inválida
                    fgAdicionaErro xmlErrosNegocio, 4063
                End If
            ElseIf lngTipoMensagem <> enumTipoMensagemLQS.Compromissada And _
                   lngTipoMensagem <> enumTipoMensagemLQS.Redesconto Then
                If Not flValidaDataNumerica(objDomNode.Text) Then
                 'Data de vencimento inválida
                 fgAdicionaErro xmlErrosNegocio, 4063
                 End If
            End If
        End If

        'Data da Operação de Retorno
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV_RETN")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                If Val("0" & objDomNode.Text) <> 0 Then
                    'Data da Operação de Retorno inválida
                    fgAdicionaErro xmlErrosNegocio, 4064
                Else
                    .selectSingleNode("DT_OPER_ATIV_RETN").Text = vbNullString
                End If
            End If
        End If

        'Preço unitário de retorno
        Set objDomNode = .selectSingleNode("VA_PU_RETN")
        If Not objDomNode Is Nothing Then
            If lngTipoMensagem = enumTipoMensagemLQS.Compromissada Then
                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao And _
                   fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                    If fgVlrXml_To_Decimal("0" & .selectSingleNode("TP_CPRO_OPER_ATIV").Text) = 1 Then
                        'Preço unitário Retorno deve ser maior que zero
                        fgAdicionaErro xmlErrosNegocio, 4056
                    End If
                End If
            End If
            If objDomNode.Text <> vbNullString Then
                If fgVlrXml_To_Decimal(objDomNode.Text) < 0 Then
                    'Preço untitário Retorno deve ser maior que zero
                    fgAdicionaErro xmlErrosNegocio, 4056
                End If
            End If
        End If

        'Data da Operação Original
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV_ORIG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Operação Original inválida
                fgAdicionaErro xmlErrosNegocio, 4065
            Else
                If Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                    'Data da Operação Original não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4066
                ElseIf fgDtXML_To_Date(objDomNode.Text) > flDataHoraServidor(Data) Then
                    'Data da Operação Original deve ser anterior a hoje
                    fgAdicionaErro xmlErrosNegocio, 4067
                End If
            End If
        End If
        
        'Data da Liquidação
        Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4068
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da Liquidação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4069
            ElseIf fgDtXML_To_Date(objDomNode.Text) <= flDataHoraServidor(enumFormatoDataHora.Data) Then
                'Data da liquidação deve ser maior que hoje
                fgAdicionaErro xmlErrosNegocio, 4070
            End If
        End If

        'Pikachu - 01/03/2005
        'Não consisitir nesta versão
        'Identificador do Lastro
'        Set objDomNode = .selectSingleNode("CO_IDEF_LAST")
'        If Not objDomNode Is Nothing Then
'            If Trim$(objDomNode.Text) = vbNullString Then
'                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
'                    lngTipoSolicitacao = enumTipoSolicitacao.Alteracao Or _
'                    lngTipoSolicitacao = enumTipoSolicitacao.CancelamentoPorLastro Then
'                   'Identificador do Lastro inválido
'                   fgAdicionaErro xmlErrosNegocio, 4203
'                End If
'            End If
'        End If

        'Consiste campos obrigatórios para CO_LOCA_LIQU = BMA (17) para tipos de mensagem Leilao (11)
        If lngTipoMensagem = enumTipoMensagemLQS.Leilao Then
            If CLng(fgSelectSingleNodeText(xmlDOMMensagem, "//CO_LOCA_LIQU")) = enumLocalLiquidacao.BMA Then
                Set objVL = CreateObject("A6A7A8.clsVeiculoLegal")
                Set xmlVL = CreateObject("MSXML2.DOMDocument.4.0")
            
                strAux = objVL.Ler(fgSelectSingleNodeText(xmlDOMMensagem, "//CO_VEIC_LEGA"), _
                                   fgSelectSingleNodeText(xmlDOMMensagem, "//SG_SIST_ORIG"))
                                    
                If xmlVL.loadXML(strAux) Then
            
                    If fgSelectSingleNodeText(xmlVL, "//TP_TITL_BMA") = vbNullString Then
                        'Tipo titular BMA não configurado para o veículo legal.
                        fgAdicionaErro xmlErrosNegocio, 4177
                    End If
                
                    If fgSelectSingleNodeText(xmlVL, "//CO_TITL_BMA") = vbNullString Then
                        'Código titular BMA não configurado para o veículo legal.
                        fgAdicionaErro xmlErrosNegocio, 4178
                    End If
                
                End If
                
                Set objVL = Nothing
                Set xmlVL = Nothing
            End If
        End If
        
        'Tipo Unilateralidade
        Set objDomNode = .selectSingleNode("TP_UNIL")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                lngTemp = Val(objDomNode.Text)
                If lngTemp <> enumTipoUnilateralidade.NenhumaParte _
                And lngTemp <> enumTipoUnilateralidade.QualquerParte _
                And lngTemp <> enumTipoUnilateralidade.SomenteCedente _
                And lngTemp <> enumTipoUnilateralidade.SomenteCessionario Then
                    'Tipo Unilateralidade invalido
                    fgAdicionaErro xmlErrosNegocio, 4456
                End If
            End If
        End If
        
        'Data Inicio Liquidacao Compromisso
        Set objDomNode = .selectSingleNode("DT_INIC_LIQU_CPRO")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data Inicio Liquidacao Compromisso invalida
                fgAdicionaErro xmlErrosNegocio, 4462
            End If
        End If
        
    End With

    fgConsisteMensagemA8 = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    Set objVL = Nothing
    Set xmlVL = Nothing
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteMensagemA8", 0
End Function

Public Function flValidaDataNumerica(ByVal pDataNumerica As Variant) As Boolean

Dim liAno                                  As Integer
Dim liMes                                  As Integer
Dim liDia                                  As Integer

    flValidaDataNumerica = False
    
    If Not IsNumeric(pDataNumerica) Then Exit Function
    
    'Valida o comprimento da data
    If Len(Trim$(pDataNumerica)) <> 8 Then Exit Function
    
    liAno = Mid$(pDataNumerica, 1, 4)
    liMes = Mid$(pDataNumerica, 5, 2)
    liDia = Mid$(pDataNumerica, 7, 2)
    
    'Valida dia para o mes de fevereiro sendo ano bissexto ou não.
    If (CInt(liAno) Mod 4) = 0 And CInt(liMes) = 2 Then
       If CInt(liDia) > 29 Then Exit Function
    ElseIf (CInt(liAno) Mod 4) <> 0 And CInt(liMes) = 2 Then
        If CInt(liDia) > 28 Then Exit Function
    End If
    
    'Validação geral da data.
    If CInt(liAno) < 1 Or CInt(liAno) > 9999 Then Exit Function
    If CInt(liMes) < 1 Or CInt(liMes) > 12 Then Exit Function
    If CInt(liDia) < 1 Or CInt(liDia) > 31 Then Exit Function
    
    flValidaDataNumerica = True
    
End Function

Public Function flValidaHoraNumerica(ByVal pHora As Variant) As Boolean

Dim iHora                                   As Integer
Dim iMinuto                                 As Integer

    flValidaHoraNumerica = False
    If Not IsNumeric(pHora) Then Exit Function
    'Valida o comprimento da hora
    If Len(Trim$(pHora)) <> 4 Then Exit Function
    iHora = Mid$(pHora, 1, 2)
    iMinuto = Mid$(pHora, 3, 2)
    'Valida hora completa
    If CInt(iHora) < 0 Or CInt(iHora) > 23 Then Exit Function
    If CInt(iMinuto) < 0 Or CInt(iMinuto) > 59 Then Exit Function
    
    flValidaHoraNumerica = True

End Function

Public Function fgValidaMensagemEspecificacaoInterna(ByRef pxmlErro As MSXML2.DOMDocument40, _
                                                     ByVal pstrNumeroControle As String, _
                                                     ByVal plngTipoSolicitacao As Long, _
                                                     ByVal plngQuantidadeTitulo As Long) As Long

Dim strSQL                                  As String
Dim strWhereTipoOper                        As String
Dim rsQuery                                 As ADODB.Recordset
Dim blnValido                               As Boolean
Dim lngErro                                 As Long

On Error GoTo ErrorHandler
    
    '1. Forma de liquidação tem que ser sempre 1 (com conta corrente), portanto código banco, agência, etc. deverão estar preenchidos
    '2. Se Numero Especificação Original não for de uma operação de especificação LIQUIDADA, rejeita
    '3. Se for um cancelmento da remessa (tipo solicitação 3), rejeitar somente se C/C já foi integrado
    
    If plngTipoSolicitacao = enumTipoSolicitacao.Inclusao Or _
       plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
        
        strSQL = " SELECT NU_CTRL_MESG_SPB_ORIG, " & _
                 "        TP_OPER,               " & _
                 "        CO_ULTI_SITU_PROC,     " & _
                 "        QT_ATIV_MERC           " & vbCrLf & _
                 " FROM   A8.TB_OPER_ATIV        " & vbCrLf & _
                 " WHERE  NU_CTRL_MESG_SPB_ORIG = '" & pstrNumeroControle & "' AND " & vbCrLf
    
        strWhereTipoOper = "        TP_OPER IN ( " & enumTipoOperacaoLQS.EspecDefinitivaIntermediacao & " , " & _
                                                     enumTipoOperacaoLQS.EspecDefinitivaCobertura & " , " & _
                                                     enumTipoOperacaoLQS.EspecTermoCobertura & " , " & _
                                                     enumTipoOperacaoLQS.EspecTermoIntermediacao & " , " & _
                                                     enumTipoOperacaoLQS.EspecCompromissadaCobertura & " , " & _
                                                     enumTipoOperacaoLQS.EspecCompromissadaIntermediacao & ")"
    
        strSQL = strSQL & strWhereTipoOper
        
        Set rsQuery = QuerySQL(strSQL)
    
        If Not rsQuery.EOF Then
            If rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.Liquidada Then
                'Número de controle de especificação deve ser igual a uma especificação de intermediação/cobertura
                 fgAdicionaErro pxmlErro, 4132
            End If
            
            If rsQuery!QT_ATIV_MERC < plngQuantidadeTitulo Then
                'Quantidade de Títulos da operação incompatível com a Operação original.
                fgAdicionaErro pxmlErro, 4154
            End If
            
        Else
            'Número de controle de especificação deve ser igual a uma especificação de intermediação/cobertura
             fgAdicionaErro pxmlErro, 4132
        End If
 
    ElseIf plngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then
    
        strSQL = " SELECT  B.CO_ULTI_SITU_PROC " & _
                 "   FROM  A8.TB_OPER_ATIV A, " & _
                 "         A8.TB_LANC_CC_CNTB B " & _
                 "  WHERE  A.NU_SEQU_OPER_ATIV = B.NU_SEQU_OPER_ATIV " & _
                 "    AND  NU_CTRL_MESG_SPB_ORIG = '" & pstrNumeroControle & "'"
        
        Set rsQuery = QuerySQL(strSQL)
    
        If Not rsQuery.EOF Then
            If rsQuery!CO_ULTI_SITU_PROC = 106 Then
                 'Operação com status de Conta Corrente inválido para cancelamento.
                 fgAdicionaErro pxmlErro, 3054
            End If
        End If
    End If
    
    Set rsQuery = Nothing

    Exit Function
ErrorHandler:

    Set rsQuery = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgValidaMensagemEspecificacaoInterna", 0
End Function

Public Function fgConsisteTituloRegistroOperacao_BMA(ByRef pxmlRemessa As MSXML2.DOMDocument40, _
                                                     ByRef pxmlErro As MSXML2.DOMDocument40, _
                                                     ByVal plngTipoNegociacao As Long, _
                                                     ByVal plngSubTipoNegociacao, _
                                                     ByVal plngTipoSolicitacao As Long) As Boolean

Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    With pxmlRemessa.selectSingleNode("/MESG")
                
        'Descrição de título
        Set objDomNode = .selectSingleNode("NU_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) = vbNullString Then
                If plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                   'Identificador de Título inválido
                   fgAdicionaErro pxmlErro, 4074
                End If
            End If
        End If
        
        'Identificador de título SELIC
        Set objDomNode = .selectSingleNode("DE_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) = vbNullString Then
                If plngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                   'Descrição Título Selic inválido
                   fgAdicionaErro pxmlErro, 4075
                End If
            End If
        End If
        
        'Data de vencimento
        Set objDomNode = .selectSingleNode("DT_VENC_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data de vencimento inválida
                fgAdicionaErro pxmlErro, 4063
            End If
        End If

    End With
    
    Exit Function
ErrorHandler:

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteTituloRegistroOperacao_BMA", 0
    
End Function

Public Function fgValidaMensagemEspecificacao(ByVal pstrNumeroControle As String, _
                                              ByVal plngTipoEspecificacao As Long, _
                                              ByVal pvntQuantidade As Variant, _
                                              ByVal plngTipoNegoBMA As Long, _
                                              ByRef pxmlErro As MSXML2.DOMDocument40, _
                                              ByRef xmlDOMMensagem As MSXML2.DOMDocument40, _
                                              ByVal plngCodigoEmpresa As Long, _
                                              ByVal pstrCodigoVeiculoLegal As String) As Long

Dim strSQL                                  As String
Dim strWhereTipoOper                        As String
Dim rsQuery                                 As ADODB.Recordset
Dim blnValido                               As Boolean
Dim lngErro                                 As Long
Dim strDHInicio                             As String
Dim strDHFim                                As String

Dim blnOriginalRegistrada                   As Boolean

On Error GoTo ErrorHandler
    
    blnOriginalRegistrada = False
    
    If plngTipoEspecificacao = enumTipoEspecificacao.Cobertura Or _
       plngTipoEspecificacao = enumTipoEspecificacao.Intermediacao Then

        strDHInicio = fgDtHrXML_To_Oracle(fgDt_To_Xml(flDataHoraServidor(DataAux)) & "000000")
        strDHFim = fgDtHrXML_To_Oracle(fgDt_To_Xml(flDataHoraServidor(DataAux)) & "235959")

        strSQL = " SELECT NU_CTRL_MESG_SPB_ORIG, " & _
                 "        TP_OPER,               " & _
                 "        CO_ULTI_SITU_PROC,     " & _
                 "        QT_ATIV_MERC           " & vbCrLf & _
                 "  FROM  A8.TB_OPER_ATIV        " & vbCrLf & _
                 " WHERE  NU_COMD_OPER = '" & pstrNumeroControle & "'" & vbCrLf & _
                 "   AND  CO_EMPR      =  " & plngCodigoEmpresa & vbCrLf & _
                 "   AND  CO_VEIC_LEGA = '" & Trim$(pstrCodigoVeiculoLegal) & "'" & _
                 "   AND  DT_OPER_ATIV BETWEEN " & strDHInicio & _
                 "   AND " & strDHFim

        Select Case plngTipoNegoBMA

            Case enumTipoNegociacaoBMA.DefinitivaD0
                
                strSQL = strSQL & _
                    " AND TP_OPER IN (" & enumTipoOperacaoLQS.DefinitivaCobertaBMA & "," & _
                                          enumTipoOperacaoLQS.DefinitivaDescobertaBMA & ")"
                strSQL = strSQL & _
                    " ORDER BY DH_ULTI_ATLZ DESC "
                    
                Set rsQuery = QuerySQL(strSQL)

                If Not rsQuery.EOF Then
                    Do While Not rsQuery.EOF
                       If rsQuery!CO_ULTI_SITU_PROC = enumStatusOperacao.Registrada Or _
                          rsQuery!CO_ULTI_SITU_PROC = enumStatusOperacao.RegistradaAutomatica Then
                           blnOriginalRegistrada = True
                           Exit Do
                       End If
                       rsQuery.MoveNext
                    Loop
                                            
                    If Not blnOriginalRegistrada Then
                        'Operação não pode ser especificada, verificar status.
                        fgAdicionaErro pxmlErro, 4158
                    Else
                        If fgVlrXml_To_Decimal("0" & pvntQuantidade) <= rsQuery!QT_ATIV_MERC Then
                        
                        Else
                           'Quantidade de Títulos da Especificação incompatível com a Operação original.
                           fgAdicionaErro pxmlErro, 4148
                        End If
                    End If
                Else
                    'Operação Definitiva original não encontrado, verificar número do comando.
                    fgAdicionaErro pxmlErro, 4200
                End If

            Case enumTipoNegociacaoBMA.TermoPapelDecorridoComCorrecao, _
                 enumTipoNegociacaoBMA.TermoPapelDecorridoSemCorrecao, _
                 enumTipoNegociacaoBMA.TermoLeilao
        
                strSQL = strSQL & _
                    " AND TP_OPER IN (" & enumTipoOperacaoLQS.OperacaoTermoCobertaBMA & "," & _
                                          enumTipoOperacaoLQS.OperacaoTermodesCobertaBMA & ")"
                strSQL = strSQL & _
                    " ORDER BY DH_ULTI_ATLZ DESC "
        
                Set rsQuery = QuerySQL(strSQL)

                If Not rsQuery.EOF Then
                
                    Do While Not rsQuery.EOF
                       If rsQuery!CO_ULTI_SITU_PROC = enumStatusOperacao.Registrada Or _
                          rsQuery!CO_ULTI_SITU_PROC = enumStatusOperacao.RegistradaAutomatica Or _
                          rsQuery!CO_ULTI_SITU_PROC = enumStatusOperacao.LiquidadaFisicamente Or _
                          rsQuery!CO_ULTI_SITU_PROC = enumStatusOperacao.LiquidadaFisicamenteAutomatica Then

                           blnOriginalRegistrada = True
                           Exit Do
                       End If
                       rsQuery.MoveNext
                    Loop
                
                    If Not blnOriginalRegistrada Then
                        'Operação não pode ser especificada, verificar status.
                        fgAdicionaErro pxmlErro, 4158
                    Else
                        If fgVlrXml_To_Decimal("0" & pvntQuantidade) <= rsQuery!QT_ATIV_MERC Then
                        
                        Else
                           'Quantidade de Títulos da Especificação incompatível com a Operação original.
                           fgAdicionaErro pxmlErro, 4148
                        End If
                    End If
                    
                Else
                    'Operação a termo original não encontrado, verificar número do comando.
                    fgAdicionaErro pxmlErro, 4201
                End If

            Case enumTipoNegociacaoBMA.Compromissada
            
                strSQL = strSQL & _
                    " AND TP_OPER IN (" & enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA & "," & _
                                          enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA & "," & _
                                          enumTipoOperacaoLQS.CompromissadaGenericaAVista & ")"
                strSQL = strSQL & _
                    " ORDER BY DH_ULTI_ATLZ DESC "

                Set rsQuery = QuerySQL(strSQL)

                If Not rsQuery.EOF Then

                    If rsQuery!TP_OPER = enumTipoOperacaoLQS.CompromissadaGenericaAVista Then
                        Do While Not rsQuery.EOF
                            If rsQuery!CO_ULTI_SITU_PROC = enumStatusOperacao.AConciliarBMA0013 Or _
                               rsQuery!CO_ULTI_SITU_PROC = enumStatusOperacao.Registrada Or _
                               rsQuery!CO_ULTI_SITU_PROC = enumStatusOperacao.RegistradaAutomatica Then
                            
                                blnOriginalRegistrada = True
                                Exit Do
                            End If
                            rsQuery.MoveNext
                        Loop
                    Else
                        Do While Not rsQuery.EOF
                            If rsQuery!CO_ULTI_SITU_PROC = enumStatusOperacao.Registrada Or _
                               rsQuery!CO_ULTI_SITU_PROC = enumStatusOperacao.RegistradaAutomatica Then
                                blnOriginalRegistrada = True
                                Exit Do
                            End If
                            rsQuery.MoveNext
                        Loop
                    End If
                Else
                    'Operação Compromissada original não encontrado, verificar número do comando.
                    fgAdicionaErro pxmlErro, 4202
                End If

        End Select

        strSQL = " SELECT NU_CTRL_MESG_SPB_ORIG, " & _
                 "        TP_OPER,               " & _
                 "        CO_ULTI_SITU_PROC,     " & _
                 "        QT_ATIV_MERC           " & vbCrLf & _
                 " FROM   A8.TB_OPER_ATIV        " & vbCrLf & _
                 " WHERE  NU_COMD_OPER = '" & pstrNumeroControle & "' " & vbCrLf & _
                 "   AND  CO_EMPR      =  " & plngCodigoEmpresa & vbCrLf & _
                 "   AND  CO_VEIC_LEGA = '" & Trim$(pstrCodigoVeiculoLegal) & "'" & _
                 "   AND  DT_OPER_ATIV BETWEEN " & strDHInicio & _
                 "   AND " & strDHFim
            
            
        strWhereTipoOper = " and       TP_OPER IN ( " & enumTipoOperacaoLQS.EspecDefinitivaIntermediacao & " , " & _
                                                        enumTipoOperacaoLQS.EspecDefinitivaCobertura & " , " & _
                                                        enumTipoOperacaoLQS.EspecTermoIntermediacao & " , " & _
                                                        enumTipoOperacaoLQS.EspecTermoCobertura & " , " & _
                                                        enumTipoOperacaoLQS.EspecCompromissadaCobertura & " , " & _
                                                        enumTipoOperacaoLQS.EspecCompromissadaIntermediacao & " ) "

        strSQL = strSQL & strWhereTipoOper
        strSQL = strSQL & _
            " ORDER BY DH_ULTI_ATLZ DESC "

        Set rsQuery = QuerySQL(strSQL)

        If Not rsQuery.EOF Then
                   
            Select Case rsQuery!TP_OPER
                   
                Case enumTipoOperacaoLQS.EspecCompromissadaCobertura, _
                     enumTipoOperacaoLQS.EspecCompromissadaIntermediacao, _
                     enumTipoOperacaoLQS.EspecDefinitivaCobertura, _
                     enumTipoOperacaoLQS.EspecDefinitivaIntermediacao, _
                     enumTipoOperacaoLQS.EspecTermoCobertura, _
                     enumTipoOperacaoLQS.EspecTermoIntermediacao
                     
                     If rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.Liquidada And _
                        rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.Rejeitada And _
                        rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.CanceladaOrigem And _
                        rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.Cancelada And _
                        rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.EmSer And _
                        rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.Concordancia And _
                        rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.ConcordanciaAutomatica And _
                        rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.LiquidadaFisicamente And _
                        rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.LiquidadaFisicamenteAutomatica Then
                         'Operação não pode ser especificada, verificar status.
                         fgAdicionaErro pxmlErro, 4158
                     End If
            
            End Select
            
        End If
        
    ElseIf plngTipoEspecificacao = enumTipoEspecificacao.Cancelamento Then
            
            strSQL = " SELECT NU_CTRL_MESG_SPB_ORIG, " & _
                     "        TP_OPER,               " & _
                     "        CO_ULTI_SITU_PROC,     " & _
                     "        QT_ATIV_MERC           " & vbCrLf & _
                     " FROM   A8.TB_OPER_ATIV        " & vbCrLf & _
                     " WHERE  NU_CTRL_MESG_SPB_ORIG = '" & pstrNumeroControle & "' AND " & vbCrLf
        
            strWhereTipoOper = "        TP_OPER IN ( " & enumTipoOperacaoLQS.EspecDefinitivaIntermediacao & " , " & _
                                                         enumTipoOperacaoLQS.EspecDefinitivaCobertura & " , " & _
                                                         enumTipoOperacaoLQS.EspecTermoIntermediacao & " , " & _
                                                         enumTipoOperacaoLQS.EspecTermoCobertura & " , " & _
                                                         enumTipoOperacaoLQS.EspecCompromissadaCobertura & " , " & _
                                                         enumTipoOperacaoLQS.EspecCompromissadaIntermediacao & " ) "
                                                         
            strSQL = strSQL & strWhereTipoOper
            strSQL = strSQL & _
                " ORDER BY DH_ULTI_ATLZ DESC "

            Set rsQuery = QuerySQL(strSQL)
                
            If Not rsQuery.EOF Then
                    
                    If rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.Liquidada Then
                        'Operação não pode ser especificada, verificar status.
                        fgAdicionaErro pxmlErro, 4158
                    End If
                
                    strSQL = " SELECT CO_ULTI_SITU_PROC      " & _
                             " FROM   A8.TB_OPER_ATIV        " & _
                             " WHERE  NU_CTRL_MESG_SPB_ORIG = '" & pstrNumeroControle & "'" & _
                             "   AND  TP_OPER = " & enumTipoOperacaoLQS.OperacaoDefinitivaInternaBMA
                             
                    Set rsQuery = QuerySQL(strSQL)
                                    
                    If Not rsQuery.EOF Then
                        If rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.Cancelada And _
                           rsQuery!CO_ULTI_SITU_PROC <> enumStatusOperacao.CanceladaOrigem Then
                            'Operações de especificação vinculadas a uma operaçao interna não podem ser canceladas.
                            fgAdicionaErro pxmlErro, 4151
                        End If
                    End If
            Else
                'Número de controle de especificação deve ser igual a uma operação original
                fgAdicionaErro pxmlErro, 4133
            End If
    End If
    
    blnValido = lngErro = 0
    fgValidaMensagemEspecificacao = blnValido
    Set rsQuery = Nothing

    Exit Function
    
ErrorHandler:

    Set rsQuery = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgValidaMensagemEspecificacao", 0
    
End Function


Public Function fgValidaMensagemEspecificacaoCETIP(ByVal pstrNumeroControle As String, _
                                                   ByVal plngCodigoEmpresa As Long, _
                                                   ByVal pstrCodigoVeiculoLegal As String, _
                                                   ByRef pxmlErro As MSXML2.DOMDocument40) As Long

Dim strSQL                                  As String
Dim strWhereTipoOper                        As String
Dim rsQuery                                 As ADODB.Recordset
Dim blnValido                               As Boolean
Dim lngErro                                 As Long

On Error GoTo ErrorHandler

    pstrNumeroControle = IIf(IsNumeric(pstrNumeroControle), Val(pstrNumeroControle), pstrNumeroControle)
    
    strSQL = " SELECT NU_CTRL_MESG_SPB_ORIG, " & _
             "        TP_OPER,               " & _
             "        CO_ULTI_SITU_PROC      " & _
             "  FROM  A8.TB_OPER_ATIV        " & vbCrLf & _
             " WHERE  NU_COMD_OPER      = '" & pstrNumeroControle & "'" & vbCrLf & _
             "   AND  CO_EMPR           =  " & plngCodigoEmpresa & vbCrLf & _
             "   AND  CO_VEIC_LEGA      = '" & Trim$(pstrCodigoVeiculoLegal) & "'" & _
             "   AND  CO_ULTI_SITU_PROC =  " & enumStatusOperacao.Liquidada & vbCrLf & _
             "   AND TP_OPER IN (" & enumTipoOperacaoLQS.MovimentacaoInstrumentoFinanceiro & "," & _
                                     enumTipoOperacaoLQS.ResgateFundoInvestimento & "," & _
                                     enumTipoOperacaoLQS.MovInstrumentoFinanceiroConciliacao & "," & _
                                     enumTipoOperacaoLQS.AplicacaoFundoInvestimentoCETIP & "," & _
                                     enumTipoOperacaoLQS.DepositoFundoInvestimentoCETIP & "," & _
                                     enumTipoOperacaoLQS.DepositoFundoInvestimentoConciliacaoCETIP & ")" & _
             " ORDER BY DH_ULTI_ATLZ DESC "
 
    Set rsQuery = QuerySQL(strSQL)

    If rsQuery.EOF Then
        'Operação original com Status Liquidada não encontrada. Favor verificar o status da operação original, ou o número de comando da especificação.
        fgAdicionaErro pxmlErro, 4253
    End If

    blnValido = lngErro = 0
    fgValidaMensagemEspecificacaoCETIP = blnValido
    Set rsQuery = Nothing

    Exit Function

ErrorHandler:

    Set rsQuery = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgValidaMensagemEspecificacaoCETIP", 0

End Function

Public Function fgConsisteMensagemAlteracaoCC(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim datDataVigencia                         As Date
Dim lngTipoMensagem                         As Long
Dim lngEmpresa                              As Long
Dim lngCodigoEmpresaFusi                    As Long
Dim lngTipoSolicitacao                      As Long
Dim blnEntradaManual                        As Boolean
Dim lngDebitoCredito                        As Long

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")
    
        datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)
    
        lngTipoMensagem = fgVlrXml_To_Decimal(.selectSingleNode("TP_MESG").Text)
        
        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        Else
            lngEmpresa = glngCodigoEmpresa
            lngCodigoEmpresaFusi = glngCodigoEmpresaFusionada
        End If
        
        If Not fgExisteSistema(.selectSingleNode("SG_SIST_ORIG").Text, glngCodigoEmpresa, datDataVigencia) Then
            'Sistema inexistente
            fgAdicionaErro xmlErrosNegocio, 4004
        End If

        If Not .selectSingleNode("SG_SIST_DEST").Text = "A8" Then
            'Sistema destino incorreto
            fgAdicionaErro xmlErrosNegocio, 4005
        End If

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
        
            If Not blnEntradaManual Then
                If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                               objDomNode.Text) Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If

        'Data da Mensagem Invalida
        Set objDomNode = .selectSingleNode("DT_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            ElseIf fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            End If
        End If
        
        'Hora de envio
        Set objDomNode = .selectSingleNode("HO_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaHoraNumerica(fgCompletaString(objDomNode.Text, "0", 4, True)) Then
                'Hora da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4061
            End If
        End If

        'Indicador de débito/crédito
        Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                lngDebitoCredito = Val(objDomNode.Text)
                If lngDebitoCredito <> enumTipoDebitoCredito.Credito And _
                   lngDebitoCredito <> enumTipoDebitoCredito.Debito Then
                    'Indicador de débito/crédito inválido
                    fgAdicionaErro xmlErrosNegocio, 4021
                End If
            End If
        End If
        
        'Veículo Legal
        If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                    .selectSingleNode("CO_VEIC_LEGA").Text, _
                                    datDataVigencia, _
                                    .selectSingleNode("CO_EMPR").Text) Then
            'Veículo legal inexistente
            fgAdicionaErro xmlErrosNegocio, 4007
        End If
        
        'Data da operação
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            'Data da Operação não pode ser maior que hoje
            If fgDtXML_To_Date(objDomNode.Text) > flDataHoraServidor(enumFormatoDataHora.Data) Then
                'Data de movimento inválida
                fgAdicionaErro xmlErrosNegocio, 4012
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da operação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4059
            End If
        End If

        'Valor financeiro
        Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal("0" & objDomNode.Text) = 0 Then
                'Valor Financeiro deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4049
            End If
        End If
                    
        'Código da Agência
        Set objDomNode = .selectSingleNode("CO_AGEN")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) = 0 Then
                'Código da Agência inválido
                fgAdicionaErro xmlErrosNegocio, 4028
            End If
        End If
    
        'Número da Conta Corrente
        Set objDomNode = .selectSingleNode("NU_CC")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) = 0 Then
                'Código da conta corrente inválido
                fgAdicionaErro xmlErrosNegocio, 4029
            End If
        End If
        
    End With
    
    Call flValidaRemessaAlteraçãoCC(xmlDOMMensagem, xmlErrosNegocio)
    
    fgConsisteMensagemAlteracaoCC = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteMensagemAlteracaoCC", 0


End Function

Private Function flValidaRemessaAlteraçãoCC(ByRef pxmlRemessa As MSXML2.DOMDocument40, _
                                            ByRef pxmlErro As MSXML2.DOMDocument40) As Boolean

'Pikachu - 21/10/2004
'Release 2 - A8
'Demanda 967 - Alteração dos controles de operações recebidas pelo SLCC,
'Alteração de dados de conta corrente

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset
Dim pvntNuSequOperAtiv                      As Variant

On Error GoTo ErrorHandler

    'Verifica se existe operação original para alteração dados Conta Corrente
    strSQL = "  SELECT  NU_SEQU_OPER_ATIV  " & vbCrLf & _
             "  FROM    A8.TB_OPER_ATIV A, " & vbCrLf & _
             "          A8.TB_TIPO_OPER B  " & vbCrLf & _
             "  WHERE   A.TP_OPER           =   B.TP_OPER " & vbCrLf & _
             "    AND   DT_OPER_ATIV        =  " & fgDtXML_To_Oracle(pxmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text) & vbCrLf & _
             "    AND   CO_VEIC_LEGA        = '" & pxmlRemessa.selectSingleNode("//CO_VEIC_LEGA").Text & "'" & vbCrLf & _
             "    AND   IN_OPER_DEBT_CRED   =  " & pxmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text & vbCrLf & _
             "    AND   VA_OPER_ATIV        =  " & Replace(pxmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text, ",", ".") & vbCrLf & _
             "    AND   TP_MESG_RECB_INTE   = '" & Val(pxmlRemessa.selectSingleNode("//TP_MESG_ORIG").Text) & "'" & vbCrLf & _
             "    AND   CO_OPER_ATIV        = '" & pxmlRemessa.selectSingleNode("//CO_OPER_ATIV").Text & "'"
    
    Set objRS = QuerySQL(strSQL)
    
    If objRS.EOF Then
        'Operação não localizada no SLCC para alteração.
        fgAdicionaErro pxmlErro, 3114
        Set objRS = Nothing
        Exit Function
    End If
    
    pvntNuSequOperAtiv = objRS!NU_SEQU_OPER_ATIV
    
    'Verifica Status de lançamento conta corrente
    strSQL = " SELECT CO_ULTI_SITU_PROC,  " & _
             "        DH_ULTI_ATLZ        " & _
             "   FROM A8.TB_LANC_CC_CNTB  " & _
             "  WHERE NU_SEQU_OPER_ATIV = " & pvntNuSequOperAtiv
        
    Set objRS = QuerySQL(strSQL)
    
    If objRS.EOF Then
        'Operação não localizada no SLCC para alteração.
        fgAdicionaErro pxmlErro, 3114
        Set objRS = Nothing
        Exit Function
    Else
    
        If objRS!CO_ULTI_SITU_PROC <> enumStatusIntegracao.Disponível And _
           objRS!CO_ULTI_SITU_PROC <> enumStatusIntegracao.EnviadoCC And _
           objRS!CO_ULTI_SITU_PROC <> enumStatusIntegracao.Suspenso And _
           objRS!CO_ULTI_SITU_PROC <> enumStatusIntegracao.Antecipado Then
            'Status de lançamento conta corrente não permite alteração.
            fgAdicionaErro pxmlErro, 9999
            Set objRS = Nothing
            Exit Function
        Else
            Call fgAppendNode(pxmlRemessa, "MESG", "NU_SEQU_OPER_ATIV", pvntNuSequOperAtiv)
            Call fgAppendNode(pxmlRemessa, "MESG", "DH_ULTI_ATLZ", fgDtHr_To_Xml(objRS!DH_ULTI_ATLZ))
        End If
    End If
    
    objRS.Close
    Set objRS = Nothing
    Exit Function
ErrorHandler:
    Set objRS = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "flValidaRemessaAlteraçãoCC", 0

End Function

Public Function fgConsisteMensagemLivreMovimentacao(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim datDataVigencia                         As Date
Dim lngCodigoEmpresaFusi                    As Long
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40

Dim lngTipoMensagem                         As Long
Dim lngTipoSolicitacao                      As Long

Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")

        If Not fgExisteTipoMensagem(.selectSingleNode("TP_MESG").Text, datDataVigencia) Then
            'Tipo de mensagem inválida
            fgAdicionaErro xmlErrosNegocio, 4003
        Else
            lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)
        End If

        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        End If

        If Not fgExisteSistema(.selectSingleNode("SG_SIST_ORIG").Text, glngCodigoEmpresa, datDataVigencia) Then
            'Sistema inexistente
            fgAdicionaErro xmlErrosNegocio, 4004
        End If

        If Not .selectSingleNode("SG_SIST_DEST").Text = "A8" Then
            'Sistema destino incorreto
            fgAdicionaErro xmlErrosNegocio, 4005
        End If
        
        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then

            If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                           objDomNode.Text) Then
                'Tipo Solicitação inválido
                fgAdicionaErro xmlErrosNegocio, 4016
            End If

            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If

        'Tipo Conta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If

        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(glngCodigoEmpresaFusionada, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If

        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(glngCodigoEmpresaFusionada, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If

        'Código do indexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If

        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            Else
                'Produto inválido
                fgAdicionaErro xmlErrosNegocio, 4011
            End If
        End If

        'Tipo Conta
        Set objDomNode = .selectSingleNode("TP_CONT_ESTO")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If

        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM_ESTO")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(glngCodigoEmpresaFusionada, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If

        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC_ESTO")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(glngCodigoEmpresaFusionada, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If

        'Código do indexador
        Set objDomNode = .selectSingleNode("CO_INDX_ESTO")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If

        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD_ESTO")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            Else
                'Produto inválido
                fgAdicionaErro xmlErrosNegocio, 4011
            End If
        End If

        'Valor Financeiro
        Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal("0" & objDomNode.Text) <= 0 Then
                'Valor Financeiro deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4049
            End If
        End If

    End With

    fgConsisteMensagemLivreMovimentacao = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteMensagemLivreMovimentacao", 0
End Function

Public Function fgConsisteMensagemCBLC(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim datDataVigencia                         As Date
Dim datLimiteFech                           As Date
Dim lngCodigoEmpresaFusi                    As Long
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40

Dim lngTipoMensagem                         As Long
Dim lngTipoSolicitacao                      As Long
Dim lngTipoContraparte                      As Long
Dim lngFormaLiquidacao                      As Long
Dim lngTemp                                 As Long
Dim datTemp                                 As Date
Dim blnEntradaManual                        As Boolean
Dim intClienteQualificado                   As Integer

Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")

        If Not fgExisteTipoMensagem(.selectSingleNode("TP_MESG").Text, datDataVigencia) Then
            'Tipo de mensagem inválida
            fgAdicionaErro xmlErrosNegocio, 4003
        Else
            lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)
        End If

        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        End If
        
        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
        
            If Not blnEntradaManual Then
                If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                               objDomNode.Text) Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If
        
        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) = vbNullString Then
                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                   'Código do local de liquidação inválido
                   fgAdicionaErro xmlErrosNegocio, 4008
                End If
            Else
                If Not fgExisteLocalLiquidacao(glngCodigoEmpresaFusionada, _
                                               objDomNode.Text, _
                                               datDataVigencia) Then
                    'Código do local de liquidação inválido
                    fgAdicionaErro xmlErrosNegocio, 4008
                Else
                    If lngTipoMensagem = enumTipoMensagemLQS.RegistroLiquidacaoMultilateralCBLC Then
                        If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.CLBCAcoes Then
                            'Código do local de liquidação inválido
                            fgAdicionaErro xmlErrosNegocio, 4008
                        End If
                    End If
                End If
            End If
        'KIDA - CBLC - 19/09/2008
        Else
            'Código do local de liquidação inválido
            fgAdicionaErro xmlErrosNegocio, 4008
        End If

        'Tipo Conta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If
        
        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(glngCodigoEmpresaFusionada, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If

        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(glngCodigoEmpresaFusionada, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If
        
        'Código do indexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If

        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If lngTipoMensagem = enumEvento.InformacoesCadastraisVeiculoLegal Then
                If Trim$(objDomNode.Text) = vbNullString Then
                        'Veiculo legal inválido
                        fgAdicionaErro xmlErrosNegocio, 4007
                End If
            Else
                If Not Trim$(objDomNode.Text) = vbNullString Then
                    If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                                objDomNode.Text, _
                                                datDataVigencia, _
                                                .selectSingleNode("CO_EMPR").Text) Then
                        'Veiculo legal inválido
                        fgAdicionaErro xmlErrosNegocio, 4007
                    End If
                End If
            End If
        End If

        'Código Participante Negociação
        Set objDomNode = .selectSingleNode("CO_PARP_NEGO")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                If Not fgExisteParticipanteNegociacao(.selectSingleNode("CO_PARP_NEGO").Text) Then
                    'Código Participante Negociação inválido.
                    fgAdicionaErro xmlErrosNegocio, 4210
                End If
            End If
        End If

        'Indicador de débito/crédito
        Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                lngTemp = Val(objDomNode.Text)
                If lngTemp <> enumTipoDebitoCredito.Credito And _
                   lngTemp <> enumTipoDebitoCredito.Debito Then
                    'Indicador de débito/crédito inválido
                    fgAdicionaErro xmlErrosNegocio, 4021
                End If
            End If
        End If

        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            End If
        End If

        'Forma de Liquidação
        lngFormaLiquidacao = 0
        Set objDomNode = .selectSingleNode("CO_FORM_LIQU")
        If Not objDomNode Is Nothing Then
            lngTemp = Val("0" & objDomNode.Text)
            If lngTemp <> enumFormaLiquidacao.ContaCorrente And _
               lngTemp <> enumFormaLiquidacao.Contabil Then
                'Forma de Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4026
            Else
                lngFormaLiquidacao = lngTemp
            End If
        End If

        'Número Operação
        Set objDomNode = .selectSingleNode("NU_COMD_OPER")
        If Not objDomNode Is Nothing Then
            If Len(objDomNode.Text) > 8 Or _
               (Not IsNumeric(objDomNode.Text) And Trim$(objDomNode.Text) <> vbNullString) Then
                'Número do comando inválido
                fgAdicionaErro xmlErrosNegocio, 4100
            End If
        End If

        Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal("0" & objDomNode.Text) <= 0 Then
                'Valor Financeiro deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4049
            End If
        End If

        'Data da Liquidação
        Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4068
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da Liquidação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4069
            ElseIf fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                'KIDA - CBLC - 29/09/2008
                fgAdicionaErro xmlErrosNegocio, 4068
            End If
        End If
        
        'Data da Negociação
        Set objDomNode = .selectSingleNode("DT_NEGO")
        
        If lngTipoMensagem = enumTipoMensagemLQS.RegistroLiquidacaoMultilateralCBLC Then
            If Not objDomNode Is Nothing Then
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data da Negociação inválida
                    fgAdicionaErro xmlErrosNegocio, 4206
                ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                    'Data da Negociação não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4207
                ElseIf fgDtXML_To_Date(objDomNode.Text) < flDataHoraServidor(enumFormatoDataHora.Data) Then
                    'Data da Negociação deve ser maior que hoje
                    fgAdicionaErro xmlErrosNegocio, 4208
                End If
            End If
        Else
            If Not objDomNode Is Nothing Then
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data da Negociação inválida
                    fgAdicionaErro xmlErrosNegocio, 4206
                ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                    'Data da Negociação não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4207
                ElseIf fgDtXML_To_Date(objDomNode.Text) < fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 3, Anterior) Then
                    'Data da Negociação não é um dia útil
                    fgAdicionaErro xmlErrosNegocio, 4206
                End If
            End If
        End If
        
        'Tipo Contraparte
        Set objDomNode = .selectSingleNode("TP_CNPT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumTipoContraparte.Externo And _
               Val("0" & objDomNode.Text) <> enumTipoContraparte.Interno Then
                'Tipo Contraparte Inválido.
                fgAdicionaErro xmlErrosNegocio, 4182
            End If
        Else
            lngTipoContraparte = 0
        End If

        'Identificador Contraparte Câmara
        Set objDomNode = .selectSingleNode("CO_CNPT_CAMR")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) <> vbNullString Then
                If lngTipoContraparte = enumTipoContraparte.Interno Then
                    'Identificador Contraparte Câmara inválido.
                    fgAdicionaErro xmlErrosNegocio, 4209
                End If
            End If
        End If

        'Identificador de Cliente Qualificado CBLC
        Set objDomNode = .selectSingleNode("IN_CLIE_QULF")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumIndicadorSimNao.Sim And _
               Val("0" & objDomNode.Text) <> enumIndicadorSimNao.Nao Then
                'Identificador de Cliente Qualificado CBLC inválido.
                fgAdicionaErro xmlErrosNegocio, 4204
            Else
                intClienteQualificado = Val("0" & objDomNode.Text)
            End If
        End If

        'Código do Cliente Qualificado CBLC
        Set objDomNode = .selectSingleNode("CO_CLIE_QULF")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) = 0 Then
                If intClienteQualificado = enumIndicadorSimNao.Sim Then
                    'Código de Cliente Qualificado CBLC inválido.
                    fgAdicionaErro xmlErrosNegocio, 4205
                End If
            End If
        End If

        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_GRUP_LANC_FINC")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) = vbNullString Then
                'Código do Grupo de Lançamento Financeiro inválido.
                fgAdicionaErro xmlErrosNegocio, 4212
            Else
                If Not fgExisteGrupoLancFinanc(objDomNode.Text) Then
                    'Código do Grupo de Lançamento Financeiro inválido.
                    fgAdicionaErro xmlErrosNegocio, 4212
                End If
            End If
        End If

        'Pikachu
        'Esta condicao de compilação foi colocada para atender a versão de Homologação/Produção
        'Sem Conta Corrente
        #If ValidaCC = 1 Then
            
            'Código do Banco
            Set objDomNode = .selectSingleNode("CO_BANC")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                  (lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                   lngTipoSolicitacao = enumTipoSolicitacao.Alteracao) Then
                    If Not fgExisteCodigoBanco(Val(objDomNode.Text)) Then
                        'Código do Banco inválido
                        fgAdicionaErro xmlErrosNegocio, 4027
                    End If
                End If
            End If
    
            'Código da Agência
            Set objDomNode = .selectSingleNode("CO_AGEN")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                    lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da Agência inválido
                        fgAdicionaErro xmlErrosNegocio, 4028
                    End If
                End If
            End If
    
            'Número da conta corrente
            Set objDomNode = .selectSingleNode("NU_CC")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                    lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4029
                    End If
                End If
            End If
    
            'Valor Lançamento Conta Corrente
            Set objDomNode = .selectSingleNode("VA_LANC_CC")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                    lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    If fgVlrXml_To_Decimal(objDomNode.Text) = 0 Then
                        'Valor do lançamento na conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4030
                    Else
                        If fgVlrXml_To_Decimal(objDomNode.Text) <= fgVlrXml_To_Decimal(.selectSingleNode("//VA_OPER_ATIV").Text) Then

                        Else
                            'Valor do lançamento na conta corrente inválido
                            fgAdicionaErro xmlErrosNegocio, 4030
                        End If
                    End If
                End If
            End If

        #End If

    End With

    fgConsisteMensagemCBLC = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteMensagemCBLC", 0

End Function

Public Function fgConsisteMensagemPagDespesas(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String
' Função pra Validação do Layout 154
Dim datDataVigencia                         As Date
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40

Dim lngTemp                                 As Long
Dim lngFormaLiquidacao                      As Long
Dim lngTEDStrPag                            As Long
Dim lngOrigemRecursos                       As Long

Dim lngTipoMensagem                         As Long
Dim lngTipoSolicitacao                      As Long
Dim blnEntradaManual                        As Boolean

Dim objDomNode                              As MSXML2.IXMLDOMNode


    On Error GoTo ErrorHandler
    
        
    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)
    
    With xmlDOMMensagem.selectSingleNode("/MESG")
        'Tipo de mensagem inválida
        If Not fgExisteTipoMensagem(.selectSingleNode("TP_MESG").Text, datDataVigencia) Then
            fgAdicionaErro xmlErrosNegocio, 4003
        Else
            lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)
        End If
    
        'Empresa inválida
        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            fgAdicionaErro xmlErrosNegocio, 4006
        End If

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not fgTipoSolicitacaoValido(lngTipoMensagem, objDomNode.Text) Then
                 'Tipo Solicitação inválido
                 fgAdicionaErro xmlErrosNegocio, 4016
            Else
                 lngTipoSolicitacao = Val("0" & objDomNode.Text)
            End If
        End If

        'Data da Liquidação
        Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4068
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da Liquidação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4069
            ElseIf fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                
                If lngTipoMensagem <> enumTipoMensagemLQS.EnvioPagDespesas Then
                        If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                            'Data da Liquidação inválida
                            fgAdicionaErro xmlErrosNegocio, 4068
                        End If
                Else
                        'Data da Liquidação inválida
                        fgAdicionaErro xmlErrosNegocio, 4068
                End If
            End If
        End If

        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                        objDomNode.Text, _
                                        datDataVigencia, _
                                        .selectSingleNode("CO_EMPR").Text) Then
                'Veiculo legal inválido
                fgAdicionaErro xmlErrosNegocio, 4007
            End If
        End If
          
        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            End If
        End If
        'Origem Recursos
       
        Set objDomNode = .selectSingleNode("CO_ORIG_RECU")
        If Not objDomNode Is Nothing Then
            lngTemp = Val("0" & objDomNode.Text)
            If lngTemp <> enumOrigemRecurso.Isenta And _
               lngTemp <> enumOrigemRecurso.TransfIsentaTributada And _
               lngTemp <> enumOrigemRecurso.Tributada Then
               'Origem Recurso Inválida
               fgAdicionaErro xmlErrosNegocio, 4283
            Else
                lngOrigemRecursos = lngTemp
                
            End If
        End If
        
        'Forma de Liquidação
        lngFormaLiquidacao = 0
        Set objDomNode = .selectSingleNode("CO_FORM_LIQU")
        If Not objDomNode Is Nothing Then
            lngTemp = Val("0" & objDomNode.Text)
            If lngTemp <> enumFormaLiquidacao.ContaCorrente And _
               lngTemp <> enumFormaLiquidacao.Str And _
               lngTemp <> enumFormaLiquidacao.Boleto And _
               lngTemp <> enumFormaLiquidacao.Tributos Then
                'Forma de Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4026
            Else
                lngFormaLiquidacao = lngTemp
            End If
        End If
        
        'Código STR
        If lngFormaLiquidacao = enumFormaLiquidacao.Str Then
            Set objDomNode = .selectSingleNode("CO_MESG_STR_PAG")
            lngTemp = Val("0" & objDomNode.Text)
            
            If lngTemp <> enumTEDStrPag.strTransfContaClienteIF And _
               lngTemp <> enumTEDStrPag.strTransfIFContaCliente And _
               lngTemp <> enumTEDStrPag.strTransfContasDiferentesTitula Then
               'Código STR Inválido
                fgAdicionaErro xmlErrosNegocio, 4284
            End If
            
        End If

        'Dados da COnta Tributada do Debitado
        If lngOrigemRecursos = enumOrigemRecurso.TransfIsentaTributada Or _
           lngOrigemRecursos = enumOrigemRecurso.Tributada Then
            
            'Código do Banco Debt Tributada
            Set objDomNode = .selectSingleNode("CO_BANC_DEBT_TRIB")
            If Not objDomNode Is Nothing Then
                If Not fgExisteCodigoBanco(Val(objDomNode.Text)) Then
                    If Not fgExisteCodigoBancoCustodia(Val(objDomNode.Text)) Then
                        'Código do Banco inválido
                        fgAdicionaErro xmlErrosNegocio, 4309
                    End If
                End If
            End If
            
            'Código da Agência Debt Tributada
            Set objDomNode = .selectSingleNode("CO_AGEN_DEBT_TRIB")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) = 0 Then
                    'Código da Agência inválido
                    fgAdicionaErro xmlErrosNegocio, 4310
                End If
            End If
            'Número da Conta Débt Tributada
            
            Set objDomNode = .selectSingleNode("NU_CONT_DEBT_TRIB")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) = 0 Then
                    'Código da conta corrente inválido
                     fgAdicionaErro xmlErrosNegocio, 4311
                End If
            End If
            'Tipo da Conta Debt Trib
            Set objDomNode = .selectSingleNode("TP_CONT_DEBT_TRIB")
            If Not objDomNode Is Nothing Then
                If Not fgExisteTipoContaCorrente(objDomNode.Text) Then
                    'Tipo da Conta Inválido
                    '4285'
                    fgAdicionaErro xmlErrosNegocio, 4312
                    
                End If
            End If
            
            'Tipo de Pessoa Conta Debt Trib
            Set objDomNode = .selectSingleNode("TP_PESS_DEBT_TRIB")
            If Not objDomNode Is Nothing Then
                If Not fgExisteTipoPessoa(objDomNode.Text) Then
                    'Tipo de Pessoa Conta Debt Inválido 4313
                    fgAdicionaErro xmlErrosNegocio, 4313
                End If
            End If
        
        End If
        
        'Dados da Conta da Contraparte
        If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente Or _
           lngFormaLiquidacao = enumFormaLiquidacao.Str Then
           
           'Código ISPB
            Set objDomNode = .selectSingleNode("CO_ISPB_IF_CRED")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) = 0 Then
                    'Código ISPB Inválido 4287
                    fgAdicionaErro xmlErrosNegocio, 4287
                End If
            End If
            
            'Código do Banco Creditado
            Set objDomNode = .selectSingleNode("CO_BANC_CRED")
            If Not objDomNode Is Nothing Then
                If Not fgExisteCodigoBanco(Val(objDomNode.Text)) Then
                    If Not fgExisteCodigoBancoCustodia(Val(objDomNode.Text)) Then
                        'Código do Banco inválido
                        fgAdicionaErro xmlErrosNegocio, 4314
                    End If
                End If
            End If
            
            
            'Código da Agência Creditada
            Set objDomNode = .selectSingleNode("CO_AGEN_CRED")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) = 0 Then
                    'Código da Agência inválido
                    fgAdicionaErro xmlErrosNegocio, 4315
                End If
            End If
            'Número da Conta Creditada
            
            Set objDomNode = .selectSingleNode("NU_CONT_CRED")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) = 0 Then
                    'Código da conta corrente inválido
                     fgAdicionaErro xmlErrosNegocio, 4316
                End If
            End If
            
            'Tipo da Conta Creditada
            Set objDomNode = .selectSingleNode("TP_CONT_CRED")
            If Not objDomNode Is Nothing Then
                If Not fgExisteTipoContaCorrente(objDomNode.Text) Then
                    'Tipo da Conta Inválido
                    '4317'
                    fgAdicionaErro xmlErrosNegocio, 4317
                    
                End If
            End If
            
            'Tipo de Pessoa Conta Creditada
            Set objDomNode = .selectSingleNode("TP_PESS_CRED")
            If Not objDomNode Is Nothing Then
                If Not fgExisteTipoPessoa(objDomNode.Text) Then
                    'Tipo de Pessoa Conta Debt Inválido 4318
                    fgAdicionaErro xmlErrosNegocio, 4318
                End If
            End If
            
            'CNPJ/CPF Creditada 1
            'Set objDomNode = .selectSingleNode("CO_CNPJ_CPF_CRED_1")
            'If Not objDomNode Is Nothing Then
            '    If Len(Trim(objDomNode.Text)) = 0 Then
            '        'CNPJ/CPF Inválido 4288
            '        fgAdicionaErro xmlErrosNegocio, 4288
            '    End If
            'End If
            'Nome do Titular da Conta Creditada 1
            'Set objDomNode = .selectSingleNode("NO_TITU_CRED_1")
            'If Not objDomNode Is Nothing Then
            '    If Len(Trim(objDomNode.Text)) = 0 Then
            '        'Nome do Titular Inválido 4289
            '        fgAdicionaErro xmlErrosNegocio, 4289
            '    End If
            'End If

           'CNPJ/CPF Creditada 2
            'Set objDomNode = .selectSingleNode("CO_CNPJ_CPF_CRED_2")
            'If Not objDomNode Is Nothing Then
            '    If Len(Trim(objDomNode.Text)) = 0 Then
            '        'CNPJ/CPF Inválido 4288
            '        fgAdicionaErro xmlErrosNegocio, 4288
            '    End If
            'End If

            'Nome do Titular da Conta Creditada 2
            'Set objDomNode = .selectSingleNode("NO_TITU_CRED_2")
            'If Not objDomNode Is Nothing Then
            '    If Len(Trim(objDomNode.Text)) = 0 Then
            '        'Nome do Titular Inválido 4289
            '        fgAdicionaErro xmlErrosNegocio, 4289
            '    End If
            'End If
        End If
        
        'Dados da Conta Isenta
        If lngOrigemRecursos = enumOrigemRecurso.TransfIsentaTributada Or _
           lngOrigemRecursos = enumOrigemRecurso.Isenta Then
            'Código do Banco Debitado Isento
            Set objDomNode = .selectSingleNode("CO_BANC_DEBT_ISEN")
            If Not objDomNode Is Nothing Then
                If Not fgExisteCodigoBanco(Val(objDomNode.Text)) Then
                    If Not fgExisteCodigoBancoCustodia(Val(objDomNode.Text)) Then
                        'Código do Banco inválido
                        fgAdicionaErro xmlErrosNegocio, 4321
                    End If
                End If
            End If
        
            'Código da Agência Debitada Isenta
            Set objDomNode = .selectSingleNode("CO_AGEN_DEBT_ISEN")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) = 0 Then
                    'Código da Agência inválido
                    fgAdicionaErro xmlErrosNegocio, 4322
                End If
            End If
            'Número da Conta Debitada Isenta
            
            Set objDomNode = .selectSingleNode("NU_CONT_DEBT_ISEN")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) = 0 Then
                    'Código da conta corrente inválido
                     fgAdicionaErro xmlErrosNegocio, 4323
                End If
            End If
            'Tipo da Conta Debitada isenta
            Set objDomNode = .selectSingleNode("TP_CONT_DEBT_ISEN")
            If Not objDomNode Is Nothing Then
                If Not fgExisteTipoContaCorrente(objDomNode.Text) Then
                    'Tipo da Conta Inválido
                    '4285'
                    fgAdicionaErro xmlErrosNegocio, 4324
                    
                End If
            End If
            
            'Tipo de Pessoa Conta Debitada Isenta
            Set objDomNode = .selectSingleNode("TP_PESS_DEBT_ISEN")
            If Not objDomNode Is Nothing Then
                If Not fgExisteTipoPessoa(objDomNode.Text) Then
                    'Tipo de Pessoa Conta Debt Inválido 4325
                    fgAdicionaErro xmlErrosNegocio, 4325
                End If
            End If
        
        End If
        
        'Valor Lançamento Conta Corrente
        Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal(objDomNode.Text) = 0 Then
                'Valor do lançamento na conta corrente inválido
                 fgAdicionaErro xmlErrosNegocio, 4030
            End If
         End If
        'Valor Lançamento Conta Corrente Acrescido CPMF
        Set objDomNode = .selectSingleNode("VA_OPER_ATIV_CPMF")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal(objDomNode.Text) = 0 Then
                'Valor do lançamento na conta corrente inválido
                 fgAdicionaErro xmlErrosNegocio, 4326
            End If
         End If

        'Tipo Conta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If

        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(glngCodigoEmpresaFusionada, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If

        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(glngCodigoEmpresaFusionada, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If

        'Código do indexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If
    
    End With

    fgConsisteMensagemPagDespesas = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing
    
    

Exit Function

ErrorHandler:
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteMensagemPagDespesas", 0


End Function

Public Function fgConsisteMensagemTED(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String
' FUNCÃO DE VALIDAÇÃO 150
''''''''''''''''''''''''''
Dim datDataVigencia                         As Date
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40

Dim lngTemp                                 As Long
Dim lngFormaLiquidacao                      As Long
Dim lngTEDStrPag                            As Long
Dim lngTemp_StrPag                          As Long

Dim lngTipoMensagem                         As Long
Dim lngTipoSolicitacao                      As Long
Dim blnEntradaManual                        As Boolean

Dim objDomNode                              As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")

        'Tipo de mensagem inválida
        If Not fgExisteTipoMensagem(.selectSingleNode("TP_MESG").Text, datDataVigencia) Then
            fgAdicionaErro xmlErrosNegocio, 4003
        Else
            lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)
        End If
    
        'Empresa inválida
        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            fgAdicionaErro xmlErrosNegocio, 4006
        End If

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not fgTipoSolicitacaoValido(lngTipoMensagem, objDomNode.Text) Then
                 'Tipo Solicitação inválido
                 fgAdicionaErro xmlErrosNegocio, 4016
            Else
                 lngTipoSolicitacao = Val("0" & objDomNode.Text)
            End If
        End If

        'Data da Liquidação
        Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4068
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da Liquidação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4069
            ElseIf fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                
                If lngTipoMensagem <> enumTipoMensagemLQS.EnvioTEDClientes Then
                        If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                            'Data da Liquidação inválida
                            fgAdicionaErro xmlErrosNegocio, 4068
                        End If
                    Else
                        'Data da Liquidação inválida
                        fgAdicionaErro xmlErrosNegocio, 4068
                End If
            End If
        End If

        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                        objDomNode.Text, _
                                        datDataVigencia, _
                                        .selectSingleNode("CO_EMPR").Text) Then
                'Veiculo legal inválido
                fgAdicionaErro xmlErrosNegocio, 4007
            End If
        End If
          
        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            End If
        End If
        
        'Forma de Liquidação
        lngFormaLiquidacao = 0
        Set objDomNode = .selectSingleNode("CO_FORM_LIQU")
        If Not objDomNode Is Nothing Then
            lngTemp = Val("0" & objDomNode.Text)
            If lngTemp <> enumFormaLiquidacao.ContaCorrente And _
               lngTemp <> enumFormaLiquidacao.Contabil Then
                'Forma de Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4026
            Else
                lngFormaLiquidacao = lngTemp
            End If
        End If

        #If ValidaCC = 1 Then

            'Código do Banco
            Set objDomNode = .selectSingleNode("CO_BANC_DEBT")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente Then
                    If Not fgExisteCodigoBanco(Val(objDomNode.Text)) Then
                        'Código do Banco inválido
                        fgAdicionaErro xmlErrosNegocio, 4027
                    End If
                End If
            End If

            'Código da Agência
            Set objDomNode = .selectSingleNode("CO_AGEN_DEBT")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da Agência inválido
                        fgAdicionaErro xmlErrosNegocio, 4028
                    End If
                End If
            End If

            'Número da conta corrente
            Set objDomNode = .selectSingleNode("NU_CONT_DEBT")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4029
                    End If
                End If
            End If

            'Valor Lançamento Conta Corrente
            Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente Then
                    If fgVlrXml_To_Decimal(objDomNode.Text) = 0 Then
                        'Valor do lançamento na conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4030
                    End If
                End If
            End If

        #End If

        'Código da Mensagem STR / PAG
        Set objDomNode = .selectSingleNode("CO_MESG_STR_PAG")
        If objDomNode Is Nothing Then
            fgAdicionaErro xmlErrosNegocio, 4280
        Else
            Select Case Val("0" & objDomNode.Text)
            
                Case enumTEDStrPag.pagTransfContaClienteIF, _
                     enumTEDStrPag.pagTransfContasDiferentesTitula, _
                     enumTEDStrPag.pagTransfContasMesmaTitula, _
                     enumTEDStrPag.pagTransfEnvolvContasInvestimento, _
                     enumTEDStrPag.pagTransfIFContaCliente, _
                     enumTEDStrPag.pagTransfReservasBancDepositoJudi
            
                Case enumTEDStrPag.strTransfContaClienteIF, _
                     enumTEDStrPag.strTransfContasDiferentesTitula, _
                     enumTEDStrPag.strTransfContasMesmaTitula, _
                     enumTEDStrPag.strTransfEnvolvContasInvestimento, _
                     enumTEDStrPag.strTransfIFContaCliente, _
                     enumTEDStrPag.strTransfReservasBancDepositoJudi
                     
                Case Else
                
                    fgAdicionaErro xmlErrosNegocio, 4280
                     
            End Select

        End If
    
        'Tipo Conta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If

        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(glngCodigoEmpresaFusionada, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If

        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(glngCodigoEmpresaFusionada, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If

        'Código do indexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If
    
    End With

    fgConsisteMensagemTED = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    Exit Function

ErrorHandler:
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteMensagemTED", 0

End Function

Public Function fgConsisteLancamentoContaCorrenteOperacoesManuais(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim datDataVigencia                         As Date
Dim lngTipoMensagem                         As Long
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40

Dim objDomNode                              As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")

        'Tipo de mensagem inválida
        If Not fgExisteTipoMensagem(.selectSingleNode("TP_MESG").Text, datDataVigencia) Then
            fgAdicionaErro xmlErrosNegocio, 4003
        Else
            lngTipoMensagem = Val(.selectSingleNode("TP_MESG").Text)
        End If
    
        'Empresa inválida
        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            fgAdicionaErro xmlErrosNegocio, 4006
        End If

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not fgTipoSolicitacaoValido(lngTipoMensagem, objDomNode.Text) Then
                 'Tipo Solicitação inválido
                 fgAdicionaErro xmlErrosNegocio, 4016
            End If
        End If

        'Data da operação
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                
                'Data de movimento inválida
                fgAdicionaErro xmlErrosNegocio, 4012
            
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                
                'Data da operação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4059
            
            ElseIf fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) And _
                   fgDtXML_To_Date(objDomNode.Text) <> fgAdicionarDiasUteis(flDataHoraServidor(enumFormatoDataHora.Data), 1, enumPaginacao.Anterior) Then
        
                'Data de movimento inválida
                fgAdicionaErro xmlErrosNegocio, 4012
            
            End If
        End If

        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                        objDomNode.Text, _
                                        datDataVigencia, _
                                        .selectSingleNode("CO_EMPR").Text) Then
                'Veiculo legal inválido
                fgAdicionaErro xmlErrosNegocio, 4007
            End If
        End If
          
        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.CONTA_CORRENTE Then
                'Código do local de liquidação inválido
                fgAdicionaErro xmlErrosNegocio, 4008
            End If
        End If

        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            End If
        End If
        
        'Código do Banco
        Set objDomNode = .selectSingleNode("CO_BANC_DEBT")
        If Not objDomNode Is Nothing Then
            If Not fgExisteCodigoBanco(Val(objDomNode.Text)) Then
                'Código do Banco inválido
                fgAdicionaErro xmlErrosNegocio, 4027
            End If
        End If

        'Código da Agência
        Set objDomNode = .selectSingleNode("CO_AGEN_DEBT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) = 0 Then
                'Código da Agência inválido
                fgAdicionaErro xmlErrosNegocio, 4028
            End If
        End If

        'Número da conta corrente
        Set objDomNode = .selectSingleNode("NU_CONT_DEBT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) = 0 Then
                'Código da conta corrente inválido
                fgAdicionaErro xmlErrosNegocio, 4029
            End If
        End If

        'Valor Lançamento Conta Corrente
        Set objDomNode = .selectSingleNode("VA_LANC_CC")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal(objDomNode.Text) = 0 Then
                'Valor do lançamento na conta corrente inválido
                fgAdicionaErro xmlErrosNegocio, 4030
            End If
        End If

        'Indicador de débito/crédito
        Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                If Val(objDomNode.Text) <> enumTipoDebitoCredito.Credito And _
                   Val(objDomNode.Text) <> enumTipoDebitoCredito.Debito Then
                    'Indicador de débito/crédito inválido
                    fgAdicionaErro xmlErrosNegocio, 4021
                End If
            End If
        End If

    End With

    fgConsisteLancamentoContaCorrenteOperacoesManuais = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    Exit Function

ErrorHandler:
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteLancamentoContaCorrenteOperacoesManuais", 0

End Function

Public Function fgConsisteMensagemBMC(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim datDataVigencia                         As Date
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40

Dim lngTipoMensagem                         As Long
Dim lngTipoSolicitacao                      As Long
Dim blnEntradaManual                        As Boolean

Dim objDomNode                              As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")

        If Not fgExisteTipoMensagem(.selectSingleNode("TP_MESG").Text, datDataVigencia) Then
            'Tipo de mensagem inválida
            fgAdicionaErro xmlErrosNegocio, 4003
        Else
            lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)
        End If

        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        End If

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If

            If Not blnEntradaManual Then
                If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                               objDomNode.Text) Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If

            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If

        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) = vbNullString Then
                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                   'Código do local de liquidação inválido
                   fgAdicionaErro xmlErrosNegocio, 4008
                End If
            Else
                If Not fgExisteLocalLiquidacao(glngCodigoEmpresaFusionada, _
                                               objDomNode.Text, _
                                               datDataVigencia) Then
                    'Código do local de liquidação inválido
                    fgAdicionaErro xmlErrosNegocio, 4008
                Else
                    If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.BMC Then
                        'Código do local de liquidação inválido
                        fgAdicionaErro xmlErrosNegocio, 4008
                    End If
                End If
            End If
        End If

        'Tipo Conta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If

        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(glngCodigoEmpresaFusionada, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If

        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(glngCodigoEmpresaFusionada, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If

        'Código do indexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If

        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If lngTipoMensagem = enumEvento.InformacoesCadastraisVeiculoLegal Then
                If Trim$(objDomNode.Text) = vbNullString Then
                        'Veiculo legal inválido
                        fgAdicionaErro xmlErrosNegocio, 4007
                End If
            Else
                If Not Trim$(objDomNode.Text) = vbNullString Then
                    If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                                objDomNode.Text, _
                                                datDataVigencia, _
                                                .selectSingleNode("CO_EMPR").Text) Then
                        'Veiculo legal inválido
                        fgAdicionaErro xmlErrosNegocio, 4007
                    End If
                End If
            End If
        End If

        'Indicador de débito/crédito
        Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                If Val(objDomNode.Text) <> enumTipoDebitoCredito.Credito And _
                   Val(objDomNode.Text) <> enumTipoDebitoCredito.Debito Then
                    'Indicador de débito/crédito inválido
                    fgAdicionaErro xmlErrosNegocio, 4021
                End If
            End If
        End If

        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            End If
        End If

        'Código do Produto Destino
        Set objDomNode = .selectSingleNode("CO_PROD_DEST")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            End If
        End If

        'Valor Financeiro
        Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal("0" & objDomNode.Text) <= 0 Then
                'Valor Financeiro deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4049
            End If
        End If

        'Valor Moeda estrangeira
        Set objDomNode = .selectSingleNode("VA_MOED_ESTR_BMC")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal("0" & objDomNode.Text) <= 0 Then
                'Valor Moeda Estrangeira Inválido
                fgAdicionaErro xmlErrosNegocio, 4227
            End If
        End If

        'Taxa Câmbio Negociada
        Set objDomNode = .selectSingleNode("PE_TAXA_NEGO")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal("0" & objDomNode.Text) <= 0 Then
                'Taxa Câmbio Negociada Inválida
                fgAdicionaErro xmlErrosNegocio, 4228
            End If
        End If

        'Nome Contraparte
        Set objDomNode = .selectSingleNode("NO_CNPT")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text = vbNullString Then
                If lngTipoMensagem = enumTipoMensagemLQS.RegistroOperacoesBMC Then
                    If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                        'Nome Contraparte deve ser preenchido
                        fgAdicionaErro xmlErrosNegocio, 4231
                    End If
                End If
            End If
        End If

        'Data da Liquidação / Data da Liquidação Moeda Nacional
        Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4068
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da Liquidação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4069
            End If
        End If

        'Data da Liquidação Moeda Estrangeira
        Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV_MOED_ESTR")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4068
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da Liquidação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4069
            End If
        End If

        'Data da Operação
        Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data de movimento inválida
                fgAdicionaErro xmlErrosNegocio, 4012
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da operação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4059
            End If
        End If
    
        Select Case lngTipoMensagem
            Case enumTipoMensagemLQS.RegistroOperacoesBMC
            
                'Data da Liquidação Moeda Nacional e Estrageira
                If Not .selectSingleNode("DT_LIQU_OPER_ATIV") Is Nothing And _
                   Not .selectSingleNode("DT_LIQU_OPER_ATIV_MOED_ESTR") Is Nothing Then
        
                    Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV")
                    If fgDtXML_To_Date(objDomNode.Text) < flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Liquidação Moeda Nacional inválida
                        fgAdicionaErro xmlErrosNegocio, 4224
                    End If
        
                    Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV_MOED_ESTR")
                    If fgDtXML_To_Date(objDomNode.Text) < flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Liquidação Moeda Estrangeira inválida
                        fgAdicionaErro xmlErrosNegocio, 4225
                    End If
        
                End If
        
            Case enumTipoMensagemLQS.LiquidacaoMultilateralBMC
            
                'Data da Liquidação / Data da Liquidação Moeda Nacional
                Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV")
                If Not objDomNode Is Nothing Then
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Liquidação inválida
                        fgAdicionaErro xmlErrosNegocio, 4068
                    End If
                End If
        
                'Data da Operação
                Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
                If Not objDomNode Is Nothing Then
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data de movimento inválida
                        fgAdicionaErro xmlErrosNegocio, 4012
                    End If
                End If
            
            Case Else
                
                'Data da Liquidação / Data da Liquidação Moeda Nacional
                Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV")
                If Not objDomNode Is Nothing Then
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        If lngTipoMensagem = enumTipoMensagemLQS.RegistroOperacoesRodaDolar Then
                            If fgDtXML_To_Date(objDomNode.Text) <= flDataHoraServidor(enumFormatoDataHora.Data) Then
                                'Data da Liquidação inválida
                                fgAdicionaErro xmlErrosNegocio, 4068
                            End If
                        Else
                            'Data da Liquidação inválida
                            fgAdicionaErro xmlErrosNegocio, 4068
                        End If
                    End If
                End If
        
                'Data da Liquidação Moeda Nacional e Estrageira
                If Not .selectSingleNode("DT_LIQU_OPER_ATIV") Is Nothing And _
                   Not .selectSingleNode("DT_LIQU_OPER_ATIV_MOED_ESTR") Is Nothing Then
        
                    Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV")
                    If fgDtXML_To_Date(objDomNode.Text) <= flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Liquidação Moeda Nacional inválida
                        fgAdicionaErro xmlErrosNegocio, 4224
                    End If
        
                    Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV_MOED_ESTR")
                    If fgDtXML_To_Date(objDomNode.Text) <= flDataHoraServidor(enumFormatoDataHora.Data) Then
                        'Data da Liquidação Moeda Estrangeira inválida
                        fgAdicionaErro xmlErrosNegocio, 4225
                    End If
        
                    If .selectSingleNode("DT_LIQU_OPER_ATIV").Text <> .selectSingleNode("DT_LIQU_OPER_ATIV_MOED_ESTR").Text Then
                        'Data da Liquidação Moeda Nacional deve ser igual à Data da Liquidação Moeda Estrangeira
                        fgAdicionaErro xmlErrosNegocio, 4226
                    End If
        
                End If
        
                'Data da Operação
                Set objDomNode = .selectSingleNode("DT_OPER_ATIV")
                If Not objDomNode Is Nothing Then
                    If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                        If lngTipoMensagem = enumTipoMensagemLQS.RegistroOperacoesRodaDolar Then
                            'Data de movimento inválida
                            fgAdicionaErro xmlErrosNegocio, 4012
                        End If
                    End If
                End If
                
        End Select

        'Tipo de Ativo
        Set objDomNode = .selectSingleNode("TP_ATIV")
        If Not objDomNode Is Nothing Then
            If Not fgExisteTipoAtivo(Val("0" & objDomNode.Text)) Then
                'Tipo de Ativo inexistente
                fgAdicionaErro xmlErrosNegocio, 4128
            End If
        End If

        'Identificador Título
        Set objDomNode = .selectSingleNode("NU_ATIV_MERC")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) = vbNullString Then
                'Identificador de Título inválido
                fgAdicionaErro xmlErrosNegocio, 4074
            End If
        End If

        'Quantidade do título
        Set objDomNode = .selectSingleNode("QT_ATIV_MERC_BMC")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) = 0 Then
                'Quantidade de Títulos deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4050
            End If
        End If

        'Tipo Transferência
        Set objDomNode = .selectSingleNode("TP_TRAF")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumTipoTransferenciaLDL.Depósito And _
               Val("0" & objDomNode.Text) <> enumTipoTransferenciaLDL.Retirada Then
                'Tipo de Transferência LDL inválido
                fgAdicionaErro xmlErrosNegocio, 4038
            End If
        End If

        'Finalidade Cobertura Conta
        Set objDomNode = .selectSingleNode("CO_FIND_COBE")
        If Not objDomNode Is Nothing Then
            If Not fgExisteFinalidadeCoberturaConta(Val("0" & objDomNode.Text)) Then
                'Finalidade Cobertura Conta inválida
                fgAdicionaErro xmlErrosNegocio, 4085
            End If
        End If

        'Tipo Requisição
        Set objDomNode = .selectSingleNode("TP_REQU_BMC")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal(objDomNode.Text) <> enumTipoRequisicaoBMA.Requisitado And _
               fgVlrXml_To_Decimal(objDomNode.Text) <> enumTipoRequisicaoBMA.Voluntario Then
                'Tipo Requisição inválido
                fgAdicionaErro xmlErrosNegocio, 4082
            End If
        End If

        'Natureza Alteração
        If lngTipoSolicitacao = enumTipoSolicitacao.Alteracao Then
            Set objDomNode = .selectSingleNode("CO_NATU_ALTE")
            If Not objDomNode Is Nothing Then
                If objDomNode.Text <> "ME" And _
                   objDomNode.Text <> "MN" Then
                    'Natureza Alteração deve ser igual a ME (Moeda Estrageira) ou MN (Moeda Nacional)
                    fgAdicionaErro xmlErrosNegocio, 4232
                End If
            End If
        End If

        'Código de moeda estrangeira
        Set objDomNode = .selectSingleNode("CO_MOED_ESTR")
        If Not objDomNode Is Nothing Then
            If Not fgExisteMoeda(Val("0" & objDomNode.Text)) Then
                'Código da moeda inválido
                fgAdicionaErro xmlErrosNegocio, 4117
            End If
        End If

        'Canal SISBACEN Corretora
        Set objDomNode = .selectSingleNode("CO_SISB_COTR")
        If Not objDomNode Is Nothing Then
            If Val(objDomNode.Text) = 0 Then
                'Código canal SISBACEN corretora inválido
                fgAdicionaErro xmlErrosNegocio, 4243
            End If
        End If

        'Canal Operação Interbancária
        Set objDomNode = .selectSingleNode("CO_CNAL_OPER_INTE")
        If Not objDomNode Is Nothing Then
            If Val(objDomNode.Text) <> 1 Then
                'Código canal operação interbancária inválido
                fgAdicionaErro xmlErrosNegocio, 4244
            End If
        End If

        'Código Contratação SISBACEN
        Set objDomNode = .selectSingleNode("CO_CNTR_SISB")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text = vbNullString Then
                 If lngTipoMensagem = enumTipoMensagemLQS.LiquidacaoMultilateralBMC Then
                    If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                        'Código Contratação SISBACEN Inválido
                        fgAdicionaErro xmlErrosNegocio, 4229
                    End If
                End If
            End If
        End If

        'Código Praça
        Set objDomNode = .selectSingleNode("CO_PRAC")
        If Not objDomNode Is Nothing Then
            If Val(objDomNode.Text) = 0 Then
                If lngTipoMensagem = enumTipoMensagemLQS.RegistroOperacoesRodaDolar Then
                    'Código Praça Inválido
                    fgAdicionaErro xmlErrosNegocio, 4230
                End If
            End If
        End If


        'Valida Tipo de Negociacao
        Set objDomNode = .selectSingleNode("TP_NEGO")
        If Not objDomNode Is Nothing Then
            If lngTipoMensagem = enumTipoMensagemLQS.RegistroOperacoesBMC Then
                If Val(objDomNode.Text) <> 1 _
                And Val(objDomNode.Text) <> 2 _
                And Val(objDomNode.Text) <> 3 Then
                    'Tipo de Negociacao invalido
                    fgAdicionaErro xmlErrosNegocio, 4293
                Else
                    'Se Tipo de Negociacao <> 3, o Numero Identificador Negociacao BMC deve estar preenchido
                    If Val(objDomNode.Text) <> 3 Then
                        'Valida NR_IDEF_NEGO_BMC
                        Set objDomNode = .selectSingleNode("NR_IDEF_NEGO_BMC")
                        If Not objDomNode Is Nothing Then
                            If Val(objDomNode.Text) = 0 Then
                                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                                    'O campo Numero Identificador Negociacao BMC e obrigatorio quando o
                                    'Tipo de Negociacao for 1 ou 2
                                    fgAdicionaErro xmlErrosNegocio, 4294
                                End If
                            End If
                        End If
                        'Valida CO_PRAC
                        Set objDomNode = .selectSingleNode("CO_PRAC")
                        If Not objDomNode Is Nothing Then
                            If Val(objDomNode.Text) = 0 Then
                                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                                    'O campo Codigo da Praca e obrigatorio quando o
                                    'Tipo de Negociacao for 1 ou 2
                                    fgAdicionaErro xmlErrosNegocio, 4297
                                End If
                            End If
                        End If
                    'Senao, se Tipo de Negociacao = 3, o Codigo Contratacao SISBACEN deve estar preenchido
                    Else
                        Set objDomNode = .selectSingleNode("CO_CNTR_SISB")
                        If Not objDomNode Is Nothing Then
                            If Val(objDomNode.Text) = 0 Then
                                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                                    'O campo Codigo Contratacao SISBACEN e obrigatorio quando o
                                    'Tipo de Negociacao for 3
                                    fgAdicionaErro xmlErrosNegocio, 4296
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Codigo SISBACEN IF Parte
        Set objDomNode = .selectSingleNode("CO_SISB_IF_PT")
        If Not objDomNode Is Nothing Then
            If Val(objDomNode.Text) = 0 Then
                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    'Codigo SISBACEN IF Parte invalido
                    fgAdicionaErro xmlErrosNegocio, 4298
                End If
            End If
        End If
        
        'Codigo SISBACEN IF Contraparte
        Set objDomNode = .selectSingleNode("CO_SISB_IF_CP")
        If Not objDomNode Is Nothing Then
            If Val(objDomNode.Text) = 0 Then
                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    'Codigo SISBACEN IF Contraparte invalido
                    fgAdicionaErro xmlErrosNegocio, 4299
                End If
            End If
        End If
        
        'Indicador Giro
        Set objDomNode = .selectSingleNode("IN_GIRO")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text = vbNullString Then
                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    'Indicador Giro invalido
                    fgAdicionaErro xmlErrosNegocio, 4300
                End If
            End If
        End If
        
        'Indicador Linha
        Set objDomNode = .selectSingleNode("IN_LINH")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text = vbNullString Then
                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    'Indicador Linha invalido
                    fgAdicionaErro xmlErrosNegocio, 4301
                End If
            End If
        End If

    End With

    fgConsisteMensagemBMC = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    Exit Function

ErrorHandler:
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteMensagemBMC", 0

End Function

Public Function fgConsisteMensagemBMF(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim datDataVigencia                         As Date
Dim lngCodigoEmpresaFusi                    As Long
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40

Dim lngTipoMensagem                         As Long
Dim lngTipoSolicitacao                      As Long
Dim lngTipoContraparte                      As Long
Dim lngFormaLiquidacao                      As Long
Dim lngTemp                                 As Long
Dim blnEntradaManual                        As Boolean
Dim intClienteQualificado                   As Integer

Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")

        If Not fgExisteTipoMensagem(.selectSingleNode("TP_MESG").Text, datDataVigencia) Then
            'Tipo de mensagem inválida
            fgAdicionaErro xmlErrosNegocio, 4003
        Else
            lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)
        End If

        If Not fgExisteEmpresa(.selectSingleNode("CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        End If
        
        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
        
            If Not blnEntradaManual Then
                If Not fgTipoSolicitacaoValido(lngTipoMensagem, _
                                               objDomNode.Text) Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If
        
        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Trim$(objDomNode.Text) = vbNullString Then
                If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                   'Código do local de liquidação inválido
                   fgAdicionaErro xmlErrosNegocio, 4008
                End If
            Else
                If Not fgExisteLocalLiquidacao(glngCodigoEmpresaFusionada, _
                                               objDomNode.Text, _
                                               datDataVigencia) Then
                    'Código do local de liquidação inválido
                    fgAdicionaErro xmlErrosNegocio, 4008
                Else
                    If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.BMD Then
                        'Código do local de liquidação inválido
                        fgAdicionaErro xmlErrosNegocio, 4008
                    End If
                End If
            End If
        End If

        'Tipo Conta
        Set objDomNode = .selectSingleNode("TP_CONT")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteTipoConta(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Tipo de Conta inválido
                    fgAdicionaErro xmlErrosNegocio, 4017
                End If
            End If
        End If

        'Código do Segmento
        Set objDomNode = .selectSingleNode("CO_SEGM")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteSegmento(glngCodigoEmpresaFusionada, _
                                        objDomNode.Text, _
                                        datDataVigencia) Then
                    'Segmento inválido
                    fgAdicionaErro xmlErrosNegocio, 4018
                End If
            End If
        End If

        'Código do Evento financeiro
        Set objDomNode = .selectSingleNode("CO_EVEN_FINC")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteEventoFinanceiro(glngCodigoEmpresaFusionada, _
                                                objDomNode.Text, _
                                                datDataVigencia) Then
                    'Evento financeiro inválido
                    fgAdicionaErro xmlErrosNegocio, 4019
                End If
            End If
        End If

        'Código do indexador
        Set objDomNode = .selectSingleNode("CO_INDX")
        If Not objDomNode Is Nothing Then
            If Not Val("0" & objDomNode.Text) = 0 Then
                If Not fgExisteIndexador(glngCodigoEmpresaFusionada, _
                                         objDomNode.Text, _
                                         datDataVigencia) Then
                    'Indexador inválido
                    fgAdicionaErro xmlErrosNegocio, 4020
                End If
            End If
        End If

        'Código do Veículo Legal
        Set objDomNode = .selectSingleNode("CO_VEIC_LEGA")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                If Not fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
                                            objDomNode.Text, _
                                            datDataVigencia, _
                                            .selectSingleNode("CO_EMPR").Text) Then
                    'Veiculo legal inválido
                    fgAdicionaErro xmlErrosNegocio, 4007
                End If
            End If
        End If

        'Identificador do Participante Câmara
        Set objDomNode = .selectSingleNode("CO_PARP_CAMR")
        If Not objDomNode Is Nothing Then
            If Not IsNumeric(objDomNode.Text) And Trim$(objDomNode.Text) <> vbNullString Then
                 'Identificador Participante Câmara inválido
                 fgAdicionaErro xmlErrosNegocio, 4189
            End If
        End If

        'Indicador de débito/crédito
        Set objDomNode = .selectSingleNode("IN_OPER_DEBT_CRED")
        If Not objDomNode Is Nothing Then
            If Not Trim$(objDomNode.Text) = vbNullString Then
                lngTemp = Val(objDomNode.Text)
                If lngTemp <> enumTipoDebitoCredito.Credito And _
                   lngTemp <> enumTipoDebitoCredito.Debito Then
                    'Indicador de débito/crédito inválido
                    fgAdicionaErro xmlErrosNegocio, 4021
                End If
            End If
        End If

        'Código do Produto
        Set objDomNode = .selectSingleNode("CO_PROD")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4011
                End If
            End If
        End If

        'Forma de Liquidação
        lngFormaLiquidacao = 0
        Set objDomNode = .selectSingleNode("CO_FORM_LIQU")
        If Not objDomNode Is Nothing Then
            lngTemp = Val("0" & objDomNode.Text)
            If lngTemp <> enumFormaLiquidacao.ContaCorrente And _
               lngTemp <> enumFormaLiquidacao.Contabil And _
               lngTemp <> enumFormaLiquidacao.Str Then
                'Forma de Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4026
            Else
                lngFormaLiquidacao = lngTemp
            End If
        End If

        'Número Operação
        Set objDomNode = .selectSingleNode("NU_COMD_OPER")
        If Not objDomNode Is Nothing Then
            If Len(objDomNode.Text) > 8 Or _
               (Not IsNumeric(objDomNode.Text) And Trim$(objDomNode.Text) <> vbNullString) Then
                'Número do comando inválido
                fgAdicionaErro xmlErrosNegocio, 4100
            End If
        End If

        Set objDomNode = .selectSingleNode("VA_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If fgVlrXml_To_Decimal("0" & objDomNode.Text) <= 0 Then
                'Valor Financeiro deve ser maior que zero
                fgAdicionaErro xmlErrosNegocio, 4049
            End If
        End If

        'Data da Liquidação
        Set objDomNode = .selectSingleNode("DT_LIQU_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Liquidação inválida
                fgAdicionaErro xmlErrosNegocio, 4068
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da Liquidação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4069
            ElseIf fgDtXML_To_Date(objDomNode.Text) < flDataHoraServidor(enumFormatoDataHora.Data) Then
                'Data da liquidação deve ser maior que hoje
                fgAdicionaErro xmlErrosNegocio, 4070
            ElseIf fgDtXML_To_Date(objDomNode.Text) > flDataHoraServidor(enumFormatoDataHora.Data) Then
                If lngTipoMensagem = enumTipoMensagemLQS.RegistroLiquidacaoMultilateralCBLC Then
                    'Data da Liquidação inválida
                    fgAdicionaErro xmlErrosNegocio, 4068
                End If
            End If
        End If

        'Pikachu
        'Esta condicao de compilação foi colocada para atender a versão de Homologação/Produção
        'Sem Conta Corrente
        #If ValidaCC = 1 Then

            'Código do Banco
            Set objDomNode = .selectSingleNode("CO_BANC")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                  (lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
                   lngTipoSolicitacao = enumTipoSolicitacao.Alteracao) Then
                    If Not fgExisteCodigoBanco(Val(objDomNode.Text)) Then
                        'Código do Banco inválido
                        fgAdicionaErro xmlErrosNegocio, 4027
                    End If
                End If
            End If

            'Código da Agência
            Set objDomNode = .selectSingleNode("CO_AGEN")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                    lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da Agência inválido
                        fgAdicionaErro xmlErrosNegocio, 4028
                    End If
                End If
            End If

            'Número da conta corrente
            Set objDomNode = .selectSingleNode("NU_CC")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                    lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4029
                    End If
                End If
            End If

            'Valor Lançamento Conta Corrente
            Set objDomNode = .selectSingleNode("VA_LANC_CC")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.ContaCorrente And _
                    lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                    If fgVlrXml_To_Decimal(objDomNode.Text) = 0 Then
                        'Valor do lançamento na conta corrente inválido
                        fgAdicionaErro xmlErrosNegocio, 4030
                    Else
                        If fgVlrXml_To_Decimal(objDomNode.Text) <= fgVlrXml_To_Decimal(.selectSingleNode("//VA_OPER_ATIV").Text) Then

                        Else
                            'Valor do lançamento na conta corrente inválido
                            fgAdicionaErro xmlErrosNegocio, 4030
                        End If
                    End If
                End If
            End If

            'Empresa Debitada
            Set objDomNode = .selectSingleNode("CO_EMPR_DEBT")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.Str Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da Empresa Debitada Inválido.
                        fgAdicionaErro xmlErrosNegocio, 4217
                    Else
                        If Not fgExisteEmpresa(objDomNode.Text, datDataVigencia) Then
                            'Código da Empresa Debitada Inválido.
                            fgAdicionaErro xmlErrosNegocio, 4217
                        End If
                    End If
                End If
            End If

            'Empresa Creditada
            Set objDomNode = .selectSingleNode("CO_EMPR_CRED")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.Str Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da Empresa Creditada Inválido.
                        fgAdicionaErro xmlErrosNegocio, 4218
                    Else
                        If Not fgExisteEmpresa(objDomNode.Text, datDataVigencia) Then
                            'Código da Empresa Debitada Inválido.
                            fgAdicionaErro xmlErrosNegocio, 4217
                        End If
                    End If
                End If
            End If

            'Número do Documento PZ
            Set objDomNode = .selectSingleNode("NU_DOCT_PZ")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.Str Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Número do Documento PZ obrigatório.
                        fgAdicionaErro xmlErrosNegocio, 4219
                    End If
                End If
            End If

            'Histórico
            Set objDomNode = .selectSingleNode("NU_DOCT_PZ")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.Str Then
                    If Trim(objDomNode.Text) = vbNullString Then
                        'Histórico obrigatório.
                        fgAdicionaErro xmlErrosNegocio, 4220
                    End If
                End If
            End If

            'Finalidade IF
            Set objDomNode = .selectSingleNode("NU_DOCT_PZ")
            If Not objDomNode Is Nothing Then
                If lngFormaLiquidacao = enumFormaLiquidacao.Str Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        'Código da Finalidade de lançamento obrigatório.
                        fgAdicionaErro xmlErrosNegocio, 4221
                    End If
                End If
            End If

        #End If

    End With

    fgConsisteMensagemBMF = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteMensagemBMF", 0

End Function

Public Function fgExisteGrupoLancFinanc(ByVal plngCodigoGrupo As Long) As Boolean

Dim strSQL                                  As String
Dim rsGrupoLancFinc                         As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = "select  CO_GRUP_LANC_FINC "
    strSQL = strSQL & " from a8.tb_grup_lanc_finc "
    strSQL = strSQL & " where CO_GRUP_LANC_FINC = '" & plngCodigoGrupo & "'"
    
    Set rsGrupoLancFinc = QuerySQL(strSQL)
    
    If rsGrupoLancFinc.RecordCount > 0 Then
        fgExisteGrupoLancFinanc = True
    End If
    Set rsGrupoLancFinc = Nothing
    Exit Function

ErrorHandler:
    Set rsGrupoLancFinc = Nothing
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA6A8ValidaRemessa", "fgExisteGrupoLancFinanc", lngCodigoErroNegocio, intNumeroSequencialErro)

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

'CCR
Public Function fgConsisteConsultaOperCCR(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim datDataVigencia                         As Date
Dim lngTipoMensagem                         As Long
Dim lngEmpresa                              As Long
Dim lngCodigoEmpresaFusi                    As Long
Dim lngTipoSolicitacao                      As Long
Dim blnEntradaManual                        As Boolean
Dim lngDebitoCredito                        As Long

Dim strCodReemb                             As String
Dim strTpDtCCR                              As String
Dim strDtIni                                As String
Dim strDtFim                                As String

Dim strTpOpComercExtr                       As String
Dim strTpComerc                             As String

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("//MESG")
    
        datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)
    
        lngTipoMensagem = fgVlrXml_To_Decimal(.selectSingleNode("//TP_MESG").Text)
        
        If Not fgExisteEmpresa(.selectSingleNode("//CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        Else
            lngEmpresa = glngCodigoEmpresa
            lngCodigoEmpresaFusi = glngCodigoEmpresaFusionada
        End If
        
        If Not fgExisteSistema(.selectSingleNode("//SG_SIST_ORIG").Text, glngCodigoEmpresa, datDataVigencia) Then
            'Sistema inexistente
            fgAdicionaErro xmlErrosNegocio, 4004
        End If

        If Not .selectSingleNode("SG_SIST_DEST").Text = "A8" Then
            'Sistema destino incorreto
            fgAdicionaErro xmlErrosNegocio, 4005
        End If

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("//TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
        
            If Not blnEntradaManual Then
                If objDomNode.Text <> enumTipoSolicitacao.Complementacao Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        Else
            'Tipo Solicitação inválido
            fgAdicionaErro xmlErrosNegocio, 4016
        End If

        'Data da Mensagem Invalida
        Set objDomNode = .selectSingleNode("//DT_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            ElseIf fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            End If
        End If
        
        'Hora de envio
        Set objDomNode = .selectSingleNode("//HO_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaHoraNumerica(fgCompletaString(objDomNode.Text, "0", 4, True)) Then
                'Hora da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4061
            End If
        End If

        'Veículo Legal
        If Not fgExisteVeiculoLegal(.selectSingleNode("//SG_SIST_ORIG").Text, _
                                    .selectSingleNode("//CO_VEIC_LEGA").Text, _
                                    datDataVigencia, _
                                    .selectSingleNode("//CO_EMPR").Text) Then
            'Veículo legal inexistente
            fgAdicionaErro xmlErrosNegocio, 4007
        End If
        
        'Data da operação
        Set objDomNode = .selectSingleNode("//DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            'Data da Operação não pode ser maior que hoje
            If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                'Data de movimento inválida
                fgAdicionaErro xmlErrosNegocio, 4012
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da operação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4059
            End If
        End If

        If Val(.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.ConsultaOperacaoCCR Then
        
            Set objDomNode = .selectSingleNode("//TpComerc")
            If Not objDomNode Is Nothing Then
                If Trim(objDomNode.Text) = vbNullString Then
                    'Tipo Comércio inválido.
                    fgAdicionaErro xmlErrosNegocio, 4400
                Else
                    If Not fgExisteDominioTagMensagemBACEN("TpComerc", Trim(objDomNode.Text)) Then
                        'Tipo Comércio inválido.
                        fgAdicionaErro xmlErrosNegocio, 4400
                    Else
                        strTpComerc = UCase(Trim(objDomNode.Text))
                    End If
                End If
            Else
                'Tipo Comércio inválido.
                fgAdicionaErro xmlErrosNegocio, 4400
            End If
            
            Set objDomNode = .selectSingleNode("//TpOpComercExtr")
            If Not objDomNode Is Nothing Then
                If Trim(objDomNode.Text) <> vbNullString Then
                    If Not fgExisteDominioTagMensagemBACEN("TpOpComercExtr", Trim(objDomNode.Text)) Then
                        'Tipo Operação Comércio Exterior inválido.
                        fgAdicionaErro xmlErrosNegocio, 4401
                    Else
                        strTpOpComercExtr = UCase(Trim(objDomNode.Text))
                    End If
                End If
            End If
            
            'EM Emissão
            'EC   Emissão da Contraparte
            'NE Negociação
            'NC   Negociação da Contraparte
            
            'RE Reembolso
            'ER   Estorno de Reembolso
            
            'DE   Débito do Exterior
            'RC Recolhimento
            'DR   Devolução de Recolhimento
            'ED   Estorno de Débito do Exterior

            
            If strTpComerc = "EX" Then
                If Not fgIN(strTpOpComercExtr, "EM", "EC", "NE", "NC", "RE", "ER", "") Then
                    'Tipo Operação Comércio Exterior inválido.
                    fgAdicionaErro xmlErrosNegocio, 4401
                End If
                
            ElseIf strTpComerc = "IM" Then
                If Not fgIN(strTpOpComercExtr, "EM", "EC", "NE", "NC", "DE", "RC", "DR", "ED", "") Then
                    'Tipo Operação Comércio Exterior inválido.
                    fgAdicionaErro xmlErrosNegocio, 4401
                End If
            End If
            
            strCodReemb = vbNullString
            strTpDtCCR = vbNullString
            
            If Not .selectSingleNode("//CodReemb") Is Nothing Then
                strCodReemb = .selectSingleNode("//CodReemb").Text
            End If
            
            If Not .selectSingleNode("//TpDtCCR") Is Nothing Then
                strTpDtCCR = .selectSingleNode("//TpDtCCR").Text
            End If
            
            If Not .selectSingleNode("//DtIni") Is Nothing Then
                strDtIni = .selectSingleNode("//DtIni").Text
            End If
            
            If Not .selectSingleNode("//DtFim") Is Nothing Then
                strDtFim = .selectSingleNode("//DtFim").Text
            End If
            
            
            If strCodReemb = vbNullString And strTpDtCCR = vbNullString Then
                'Código Reembolso ou Tipo Data CCR deve ser informado.
                fgAdicionaErro xmlErrosNegocio, 4402
            End If
            
            
            If strTpDtCCR <> vbNullString Then
                If Not fgExisteDominioTagMensagemBACEN("TpDtCCR", strTpDtCCR) Then
                    'Tipo Data CCR inválida.
                    fgAdicionaErro xmlErrosNegocio, 4403
                Else
                    If strDtIni = vbNullString Or strDtIni = vbNullString Then
                        'Data Início e Data Fim são obrigatórios.
                        fgAdicionaErro xmlErrosNegocio, 4404
                    End If
                    
                    If strDtIni <> vbNullString Then
                        If Not flValidaDataNumerica(strDtIni) Then
                            'Data de Inicio inválida
                            fgAdicionaErro xmlErrosNegocio, 4505
                        End If
                    End If
                    
                    If strDtFim <> vbNullString Then
                        If Not flValidaDataNumerica(strDtFim) Then
                            'Data Fim inválida
                            fgAdicionaErro xmlErrosNegocio, 4506
                        End If
                    End If
                    
                End If
            End If
        
        End If
    End With
    
    fgConsisteConsultaOperCCR = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteConsultaOperCCR", 0


End Function


Public Function fgConsisteEmissaoOperCCR(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim datDataVigencia                         As Date
Dim lngTipoMensagem                         As Long
Dim lngEmpresa                              As Long
Dim lngCodigoEmpresaFusi                    As Long
Dim lngTipoSolicitacao                      As Long
Dim blnEntradaManual                        As Boolean
Dim lngDebitoCredito                        As Long
Dim datDataMovimento                        As Date
Dim datDataEmissao                          As Date
Dim strIndicadorAlteracao                   As String

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)
    
    strIndicadorAlteracao = vbNullString
    lngTipoSolicitacao = 0
    
    With xmlDOMMensagem.selectSingleNode("//MESG")
    
        datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)
    
        lngTipoMensagem = fgVlrXml_To_Decimal(.selectSingleNode("//TP_MESG").Text)
        
        If Not fgExisteEmpresa(.selectSingleNode("//CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        Else
            lngEmpresa = glngCodigoEmpresa
            lngCodigoEmpresaFusi = glngCodigoEmpresaFusionada
        End If
        
        If Not fgExisteSistema(.selectSingleNode("//SG_SIST_ORIG").Text, glngCodigoEmpresa, datDataVigencia) Then
            'Sistema inexistente
            fgAdicionaErro xmlErrosNegocio, 4004
        End If

        If Not .selectSingleNode("SG_SIST_DEST").Text = "A8" Then
            'Sistema destino incorreto
            fgAdicionaErro xmlErrosNegocio, 4005
        End If

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("//TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
            
            'a.  Tipo de Solicitação - Deve ser igual a 1 - Inclusão ou 2 - Complementação;
            If Not blnEntradaManual Then
                If objDomNode.Text <> enumTipoSolicitacao.Inclusao And _
                   objDomNode.Text <> enumTipoSolicitacao.Complementacao Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        Else
            'Tipo Solicitação inválido
            fgAdicionaErro xmlErrosNegocio, 4016
        End If

        'Data da Mensagem Invalida
        Set objDomNode = .selectSingleNode("//DT_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            ElseIf fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            End If
        End If
        
        'Hora de envio
        Set objDomNode = .selectSingleNode("//HO_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaHoraNumerica(fgCompletaString(objDomNode.Text, "0", 4, True)) Then
                'Hora da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4061
            End If
        End If

        'Veículo Legal
        If Not fgExisteVeiculoLegal(.selectSingleNode("//SG_SIST_ORIG").Text, _
                                    .selectSingleNode("//CO_VEIC_LEGA").Text, _
                                    datDataVigencia, _
                                    .selectSingleNode("//CO_EMPR").Text) Then
            'Veículo legal inexistente
            fgAdicionaErro xmlErrosNegocio, 4007
        End If
        
        'Data da operação
        Set objDomNode = .selectSingleNode("//DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            'Data da Operação não pode ser maior que hoje
            If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                'Data de movimento inválida
                fgAdicionaErro xmlErrosNegocio, 4012
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da operação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4059
            Else
                datDataMovimento = fgDtXML_To_Date(objDomNode.Text)
            End If
        End If

        'g.  Tipo Comércio Exterior - Deve ser igual a "Im" - Importação ou "Ex" - Exportação;
        Set objDomNode = .selectSingleNode("//TpComerc")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) = vbNullString Then
                'Tipo Comércio inválido.
                fgAdicionaErro xmlErrosNegocio, 4400
            Else
                If Not fgExisteDominioTagMensagemBACEN("TpComerc", Trim(objDomNode.Text)) Then
                    'Tipo Comércio inválido.
                    fgAdicionaErro xmlErrosNegocio, 4400
                End If
            End If
        Else
            'Tipo Comércio inválido.
            fgAdicionaErro xmlErrosNegocio, 4400
        End If
        
        'h.  Tipo Manutenção - Deve ser igual a "I" - Inclusão, "A" - Alteração ou "E" - Exclusão;
        Set objDomNode = .selectSingleNode("//TpManut")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) = vbNullString Then
                'Tipo manutenção inválido.
                fgAdicionaErro xmlErrosNegocio, 4407
            Else
                If Not fgExisteDominioTagMensagemBACEN("IndrIEA", Trim(objDomNode.Text)) Then
                    'Tipo manutenção inválido.
                    fgAdicionaErro xmlErrosNegocio, 4407
                Else
                    strIndicadorAlteracao = objDomNode.Text
                End If
            End If
        Else
            'Tipo manutenção inválido.
            fgAdicionaErro xmlErrosNegocio, 4407
        End If
            
        'i.  Código Reembolso - Obrigatório quando Tipo Comércio igual a "Ex" - Exportação.
        'Não pode ser preenchido quando Tipo Comércio igual a "Im" - Importação e não pode ser alterado após a inclusão;
        Set objDomNode = .selectSingleNode("//CodReemb")
        If UCase(.selectSingleNode("//TpComerc").Text) = "EX" Then
            If Not objDomNode Is Nothing Then
                If Trim(objDomNode.Text) = vbNullString Then
                    'Código Reembolso não informado para Exportação
                    fgAdicionaErro xmlErrosNegocio, 4408
                End If
            Else
                'Código Reembolso não informado para Exportação
                fgAdicionaErro xmlErrosNegocio, 4408
            End If
        Else
            If Not objDomNode Is Nothing Then
                If strIndicadorAlteracao = "I" Then
                    If Trim(objDomNode.Text) <> vbNullString Then
                        'Código Reembolso não deve ser informado para Importação
                        fgAdicionaErro xmlErrosNegocio, 4409
                    End If
                Else
                    If Trim(objDomNode.Text) = vbNullString Then
                        'Código Reembolso obrigatório
                        fgAdicionaErro xmlErrosNegocio, 4434
                    End If
                End If
            End If
        End If
        
                    
        'j.  País Contraparte - Não pode ser o código do Brasil e não pode ser alterado após a inclusão;
        Set objDomNode = .selectSingleNode("//PaisCtrapart")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) = vbNullString Then
                'Código País inválido.
                fgAdicionaErro xmlErrosNegocio, 4410
            Else
                If Trim(objDomNode.Text) = vbNullString Then
                    'Código País inválido.
                    fgAdicionaErro xmlErrosNegocio, 4410
                Else
                    If Val("0" & Trim(objDomNode.Text)) = 3 Then
                       'Código País inválido.
                       fgAdicionaErro xmlErrosNegocio, 4410
                    Else
                       If Not fgQueryDominioInternoCCR("PaisCtrapart", Trim(objDomNode.Text)) Then
                            'Código País inválido.
                            fgAdicionaErro xmlErrosNegocio, 4410
                       End If
                    End If
                
                End If
            End If
        Else
            'Código País inválido.
            fgAdicionaErro xmlErrosNegocio, 4410
        End If
        
        
        'k.  País Origem Mercadoria - Não deve ser informado quando Tipo Comércio igual a "Ex" - Exportação
        'e não pode ser informado o código do Brasil;
        Set objDomNode = .selectSingleNode("//PaisOrigemMercdria")
        If UCase(.selectSingleNode("//TpComerc").Text) = "IM" Then
            If Not objDomNode Is Nothing Then
                If Trim(objDomNode.Text) <> vbNullString Then
                    If Val("0" & Trim(objDomNode.Text)) = 3 Then
                       'País Origem Mercadoria inválido.
                       fgAdicionaErro xmlErrosNegocio, 4412
                    Else
                       If Not fgQueryDominioInternoCCR("PaisOrigemMercdria", Trim(objDomNode.Text)) Then
                            'País Origem Mercadoria inválido.
                            fgAdicionaErro xmlErrosNegocio, 4412
                       End If
                    End If
                End If
            End If
        Else
            If Not objDomNode Is Nothing Then
                If Trim(objDomNode.Text) <> vbNullString Then
                    'País Origem Mercadoria não deve ser informado para Exportação
                    fgAdicionaErro xmlErrosNegocio, 4413
                End If
            End If
        End If
        
        'm.  Tipo Instrumento CCR - Deve ser CC, CD, LA, PA e CG e não pode ser alterado após a inclusão;
        Set objDomNode = .selectSingleNode("//TpInstntoCCR")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) = vbNullString Then
                'Tipo Instrumento CCR não informado.
                fgAdicionaErro xmlErrosNegocio, 4414
            Else
                If Not fgExisteDominioTagMensagemBACEN("TpInstntoCCR", Trim(objDomNode.Text)) Then
                    'Tipo Instrumento CCR inválido.
                    fgAdicionaErro xmlErrosNegocio, 4415
                End If
            End If
        Else
            'Tipo Instrumento CCR não informado.
            fgAdicionaErro xmlErrosNegocio, 4413
        End If
            
        'n.  Data Emissão - Não deve ser preenchida quando Tipo Comércio igual a "Im" - Importação
        'e é obrigatório quando Tipo Comércio igual a "Ex" - Exportação;
        Set objDomNode = .selectSingleNode("//DtEms")
        If UCase(.selectSingleNode("//TpComerc").Text) = "EX" Then
            If Not objDomNode Is Nothing Then
                If Trim(objDomNode.Text) = vbNullString Then
                    'Data Emissão não informado para Exportação.
                    fgAdicionaErro xmlErrosNegocio, 4416
                Else
                    If Not flValidaDataNumerica(objDomNode.Text) Then
                        'Data Emissão inválida
                        fgAdicionaErro xmlErrosNegocio, 4517
                    Else
                        datDataEmissao = fgDtXML_To_Date(objDomNode.Text)
                    End If
                End If
            Else
                'Data Emissão não informado para Exportação
                fgAdicionaErro xmlErrosNegocio, 4416
            End If
        Else
            If Not objDomNode Is Nothing Then
                If Trim(objDomNode.Text) <> vbNullString Then
                    'Data Emissão não deve ser informado para Importação.
                    fgAdicionaErro xmlErrosNegocio, 4418
                End If
            End If
        End If
        
        'o.  Data Expiração – Deve ser maior ou igual à Data de Emissão e a Data de Movimento;
        Set objDomNode = .selectSingleNode("//DtExprc")
        If Not objDomNode Is Nothing Then
            If Val("0" & Trim(objDomNode.Text)) = 0 Then
                'Data Expiração não informado.
                fgAdicionaErro xmlErrosNegocio, 4419
            Else
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data Expiração inválida
                    fgAdicionaErro xmlErrosNegocio, 4420
                End If
                
                If fgDtXML_To_Date(objDomNode.Text) < datDataEmissao Then
                    'Data Expiração inválida
                    fgAdicionaErro xmlErrosNegocio, 4420
                End If
                
                If fgDtXML_To_Date(objDomNode.Text) < datDataMovimento Then
                    'Data Expiração inválida.
                    fgAdicionaErro xmlErrosNegocio, 4420
                End If
            End If
        Else
            'Data Expiração não informado.
            fgAdicionaErro xmlErrosNegocio, 4419
        End If
            
                
        'p.  Valor Emissão – Deve ser maior ou igual à zero;
        Set objDomNode = .selectSingleNode("//VlrEms")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) = vbNullString Then
                'Valor Emissão não informado.
                fgAdicionaErro xmlErrosNegocio, 4421
            Else
                If fgVlrXml_To_Decimal(objDomNode.Text) < 0 Then
                    'Valor Emissão inválido.
                    fgAdicionaErro xmlErrosNegocio, 4422
                End If
            End If
        Else
            'Valor Emissão não informado.
            fgAdicionaErro xmlErrosNegocio, 4421
        End If
        
        'q.  Tipo Pessoa Importador ou Exportador;
        Set objDomNode = .selectSingleNode("//TpPessoaImptdr_Exptdr")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) = vbNullString Then
                'Tipo Pessoa Importador ou Exportador não informado.
                fgAdicionaErro xmlErrosNegocio, 4423
            Else
                If Not fgExisteDominioTagMensagemBACEN("TpPessoa", UCase(Trim(objDomNode.Text))) Then
                    'Tipo Pessoa Importador ou Exportador inválida.
                    fgAdicionaErro xmlErrosNegocio, 4424
                End If
            End If
        Else
            'Tipo Pessoa Importador ou Exportador não informado.
            fgAdicionaErro xmlErrosNegocio, 4423
        End If
        
        'r.  CNPJ ou CPF Pessoa Importador ou Exportador;
        Set objDomNode = .selectSingleNode("//CNPJ_CPFImptdr_Exptdr")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) = vbNullString Then
                'CNPJ ou CPF Importador ou Exportador não informado.
                fgAdicionaErro xmlErrosNegocio, 4425
            End If
        Else
            'CNPJ ou CPF Importador ou Exportador não informado.
            fgAdicionaErro xmlErrosNegocio, 4425
        End If
        
        's.  Indicador Operação acima 360 dias – Obrigatório quando Tipo Comércio igual a “Ex” – Exportação
        'e não deve ser informado quando Tipo Comércio igual a “Im” – Importação;
        Set objDomNode = .selectSingleNode("//IndrOpSup360Dia")
        
        If UCase(.selectSingleNode("//TpComerc").Text) = "EX" Then
            If Not objDomNode Is Nothing Then
                If Trim(objDomNode.Text) = vbNullString Then
                    'Indicador Operação acima 360 dias para Exportação
                    fgAdicionaErro xmlErrosNegocio, 4426
                Else
                    If Not fgExisteDominioTagMensagemBACEN("Indr", Trim(objDomNode.Text)) Then
                        'Tipo Pessoa Importador ou Exportador inválida.
                        fgAdicionaErro xmlErrosNegocio, 4428
                    End If
                End If
            Else
                'Indicador Operação acima 360 dias não informado para Exportação
                fgAdicionaErro xmlErrosNegocio, 4426
            End If
        Else
            If Not objDomNode Is Nothing Then
                If Trim(objDomNode.Text) <> vbNullString Then
                    'Indicador Operação acima 360 dias não deve ser informado para Importação
                    fgAdicionaErro xmlErrosNegocio, 4427
                End If
            End If
        End If
        
        If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
            
            If strIndicadorAlteracao = "A" Or _
               strIndicadorAlteracao = "B" Or _
               strIndicadorAlteracao = "E" Then
                
                lngCodigoErroNegocio = fgValidaSituacaoOperacaoCCR(.selectSingleNode("//CO_OPER_ATIV").Text, _
                                                                   enumTipoOperacaoLQS.EmissaoOperacaoCCR)
                
                If lngCodigoErroNegocio <> 0 Then
                    If blnEntradaManual Then
                        If lngCodigoErroNegocio <> 3120 Then
                            fgAdicionaErro xmlErrosNegocio, lngCodigoErroNegocio
                        End If
                    Else
                        fgAdicionaErro xmlErrosNegocio, lngCodigoErroNegocio
                    End If
                End If
            End If
            lngCodigoErroNegocio = 0
        End If
        
                
    End With
    
    fgConsisteEmissaoOperCCR = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteEmissaoOperCCR", 0


End Function

Public Function fgValidaSituacaoOperacaoCCR(ByVal pstrCO_OPER_ATIV As String, _
                                            ByVal plngTipoOperacao As Long) As Long

Dim strSQL                                  As String
Dim strWhereTipoOper                        As String
Dim rsQuery                                 As ADODB.Recordset
Dim blnValido                               As Boolean
Dim lngErro                                 As Long

On Error GoTo ErrorHandler
    
    strSQL = " SELECT NU_CTRL_MESG_SPB_ORIG, " & _
             "        TP_OPER,               " & _
             "        CO_ULTI_SITU_PROC,     " & _
             "        QT_ATIV_MERC           " & vbCrLf & _
             " FROM   A8.TB_OPER_ATIV        " & vbCrLf & _
             " WHERE  CO_OPER_ATIV = '" & pstrCO_OPER_ATIV & "'"

    strWhereTipoOper = " AND         TP_OPER =  " & plngTipoOperacao

    strSQL = strSQL & strWhereTipoOper
    
    Set rsQuery = QuerySQL(strSQL)

    lngErro = 0

    If Not rsQuery.EOF Then
        If rsQuery!CO_ULTI_SITU_PROC = enumStatusOperacao.PendenteDeAceite Or _
           rsQuery!CO_ULTI_SITU_PROC = enumStatusOperacao.PendenteDeRegistro Then
            'Operação não pode ser alterado ou exluído. Situação da operação não permite alteração.
            lngErro = 4429
        End If
    Else
        lngErro = 3120
    End If
    
    fgValidaSituacaoOperacaoCCR = lngErro
    
    Set rsQuery = Nothing

    Exit Function
ErrorHandler:

    Set rsQuery = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgValidaSituacaoOperacaoCCR", 0
End Function



Public Function fgConsisteNegociacaoCCR(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim datDataVigencia                         As Date
Dim lngTipoMensagem                         As Long
Dim lngEmpresa                              As Long
Dim lngCodigoEmpresaFusi                    As Long
Dim lngTipoSolicitacao                      As Long
Dim blnEntradaManual                        As Boolean
Dim lngDebitoCredito                        As Long
Dim datDataMovimento                        As Date
Dim datDataEmissao                          As Date
Dim strIndicadorAlteracao                   As String
Dim strValorNegociacao                      As String
Dim strCodReembolso                         As String
Dim strTipoInstrumento                      As String

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)
    
    strIndicadorAlteracao = vbNullString
    lngTipoSolicitacao = 0
    
    With xmlDOMMensagem.selectSingleNode("//MESG")
    
        datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)
    
        lngTipoMensagem = fgVlrXml_To_Decimal(.selectSingleNode("//TP_MESG").Text)
        
        If Not fgExisteEmpresa(.selectSingleNode("//CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        Else
            lngEmpresa = glngCodigoEmpresa
            lngCodigoEmpresaFusi = glngCodigoEmpresaFusionada
        End If
        
        If Not fgExisteSistema(.selectSingleNode("//SG_SIST_ORIG").Text, glngCodigoEmpresa, datDataVigencia) Then
            'Sistema inexistente
            fgAdicionaErro xmlErrosNegocio, 4004
        End If

        If Not .selectSingleNode("SG_SIST_DEST").Text = "A8" Then
            'Sistema destino incorreto
            fgAdicionaErro xmlErrosNegocio, 4005
        End If

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("//TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
            
            'a.  Tipo de Solicitação - Deve ser igual a 1 - Inclusão ou 2 - Complementação;
            If Not blnEntradaManual Then
                If objDomNode.Text <> enumTipoSolicitacao.Inclusao And _
                   objDomNode.Text <> enumTipoSolicitacao.Complementacao Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        Else
            'Tipo Solicitação inválido
            fgAdicionaErro xmlErrosNegocio, 4016
        End If

        'Data da Mensagem Invalida
        Set objDomNode = .selectSingleNode("//DT_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            ElseIf fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            End If
        End If
        
        'Hora de envio
        Set objDomNode = .selectSingleNode("//HO_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaHoraNumerica(fgCompletaString(objDomNode.Text, "0", 4, True)) Then
                'Hora da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4061
            End If
        End If

        'Veículo Legal
        If Not fgExisteVeiculoLegal(.selectSingleNode("//SG_SIST_ORIG").Text, _
                                    .selectSingleNode("//CO_VEIC_LEGA").Text, _
                                    datDataVigencia, _
                                    .selectSingleNode("//CO_EMPR").Text) Then
            'Veículo legal inexistente
            fgAdicionaErro xmlErrosNegocio, 4007
        End If
        
        'Data da operação
        Set objDomNode = .selectSingleNode("//DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            'Data da Operação não pode ser maior que hoje
            If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                'Data de movimento inválida
                fgAdicionaErro xmlErrosNegocio, 4012
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da operação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4059
            Else
                datDataMovimento = fgDtXML_To_Date(objDomNode.Text)
            End If
        End If

        'g.  Tipo Comércio Exterior - Deve ser igual a "Im" - Importação ou "Ex" - Exportação;
        Set objDomNode = .selectSingleNode("//TpComerc")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) = vbNullString Then
                'Tipo Comércio inválido.
                fgAdicionaErro xmlErrosNegocio, 4400
            Else
                If Not fgExisteDominioTagMensagemBACEN("TpComerc", Trim(objDomNode.Text)) Then
                    'Tipo Comércio inválido.
                    fgAdicionaErro xmlErrosNegocio, 4400
                End If
            End If
        Else
            'Tipo Comércio inválido.
            fgAdicionaErro xmlErrosNegocio, 4400
        End If
        
        'h.  Tipo Manutenção - Deve ser igual a "I" - Inclusão, "A" - Alteração ou "E" - Exclusão;
        Set objDomNode = .selectSingleNode("//TpManut")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) = vbNullString Then
                'Tipo manutenção inválido.
                fgAdicionaErro xmlErrosNegocio, 4407
            Else
                If Not fgExisteDominioTagMensagemBACEN("IndrIEA", Trim(objDomNode.Text)) Then
                    'Tipo manutenção inválido.
                    fgAdicionaErro xmlErrosNegocio, 4407
                Else
                    strIndicadorAlteracao = objDomNode.Text
                End If
            End If
        Else
            'Tipo manutenção inválido.
            fgAdicionaErro xmlErrosNegocio, 4407
        End If
            
        Set objDomNode = .selectSingleNode("//CodReemb")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) = vbNullString Then
                'Código Reembolso obrigatorio
                fgAdicionaErro xmlErrosNegocio, 4434
            Else
                strCodReembolso = Trim(objDomNode.Text)
            End If
        Else
            'Código Reembolso obrigatorio
            fgAdicionaErro xmlErrosNegocio, 4434
        End If
                    
        
        Set objDomNode = .selectSingleNode("//TpInstntoCCR")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) = vbNullString Then
                'Tipo Instrumento CCR não informado.
                fgAdicionaErro xmlErrosNegocio, 4414
            Else
                If Not fgExisteDominioTagMensagemBACEN("TpInstntoCCR", Trim(objDomNode.Text)) Then
                    'Tipo Instrumento CCR inválido.
                    fgAdicionaErro xmlErrosNegocio, 4415
                Else
                    strTipoInstrumento = UCase(Trim(objDomNode.Text))
                End If
            End If
        Else
            'Tipo Instrumento CCR não informado.
            fgAdicionaErro xmlErrosNegocio, 4413
        End If
            
        
        Set objDomNode = .selectSingleNode("//DtNegc")
        If Not objDomNode Is Nothing Then
            If Val("0" & Trim(objDomNode.Text)) = 0 Then
                'Data Negociação não informado.
                fgAdicionaErro xmlErrosNegocio, 4431
            Else
                If Not flValidaDataNumerica(objDomNode.Text) Then
                    'Data Negociação Inválida.
                    fgAdicionaErro xmlErrosNegocio, 4430
                End If
                
            End If
        Else
            'Data Negociação não informado.
            fgAdicionaErro xmlErrosNegocio, 4431
        End If
        
        Set objDomNode = .selectSingleNode("//VlrNegc")
        If Not objDomNode Is Nothing Then
            If Val("0" & Trim(objDomNode.Text)) = 0 Then
                strValorNegociacao = "0"
            Else
                strValorNegociacao = Trim(objDomNode.Text)
            End If
        Else
            strValorNegociacao = "0"
        End If
                
        'Retirado da Validacao, conforme solicitacao Aline.
        'Ivan - 11/08/2011
'        If Not fgValidaValorNegociacaoCCR(strCodReembolso, strValorNegociacao, strTipoInstrumento) Then
'            'Valor Negociação maior que Valor Emissão.
'            fgAdicionaErro xmlErrosNegocio, 4432
'        End If
        
        If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
            
            If strIndicadorAlteracao = "A" Or _
               strIndicadorAlteracao = "B" Or _
               strIndicadorAlteracao = "E" Then
                
                lngCodigoErroNegocio = fgValidaSituacaoOperacaoCCR(.selectSingleNode("//CO_OPER_ATIV").Text, _
                                                                   enumTipoOperacaoLQS.NegociacaoOperacaoCCR)
                
                If lngCodigoErroNegocio <> 0 Then
                    If blnEntradaManual Then
                        If lngCodigoErroNegocio <> 3120 Then
                            fgAdicionaErro xmlErrosNegocio, lngCodigoErroNegocio
                        End If
                    Else
                        fgAdicionaErro xmlErrosNegocio, lngCodigoErroNegocio
                    End If
                End If
            End If
            lngCodigoErroNegocio = 0
        End If
        
                
    End With
    
    fgConsisteNegociacaoCCR = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteNegociacaoCCR", 0


End Function



Public Function fgConsisteDevolucaoEstornoCCR(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim datDataVigencia                         As Date
Dim lngTipoMensagem                         As Long
Dim lngEmpresa                              As Long
Dim lngCodigoEmpresaFusi                    As Long
Dim lngTipoSolicitacao                      As Long
Dim blnEntradaManual                        As Boolean
Dim lngDebitoCredito                        As Long
Dim datDataMovimento                        As Date
Dim datDataEmissao                          As Date
Dim strIndicadorAlteracao                   As String

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)
    
    strIndicadorAlteracao = vbNullString
    lngTipoSolicitacao = 0
    
    With xmlDOMMensagem.selectSingleNode("//MESG")
    
        datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)
    
        lngTipoMensagem = fgVlrXml_To_Decimal(.selectSingleNode("//TP_MESG").Text)
        
        If Not fgExisteEmpresa(.selectSingleNode("//CO_EMPR").Text, datDataVigencia) Then
            'Empresa inválida
            fgAdicionaErro xmlErrosNegocio, 4006
        Else
            lngEmpresa = glngCodigoEmpresa
            lngCodigoEmpresaFusi = glngCodigoEmpresaFusionada
        End If
        
        If Not fgExisteSistema(.selectSingleNode("//SG_SIST_ORIG").Text, glngCodigoEmpresa, datDataVigencia) Then
            'Sistema inexistente
            fgAdicionaErro xmlErrosNegocio, 4004
        End If

        If Not .selectSingleNode("SG_SIST_DEST").Text = "A8" Then
            'Sistema destino incorreto
            fgAdicionaErro xmlErrosNegocio, 4005
        End If

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("//TP_SOLI")
        If Not objDomNode Is Nothing Then
            If Not .selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                blnEntradaManual = True
            Else
                blnEntradaManual = False
            End If
            
            'a.  Tipo de Solicitação - Deve ser igual a 1 - Inclusão ou 2 - Complementação;
            If Not blnEntradaManual Then
                If objDomNode.Text <> enumTipoSolicitacao.Inclusao And _
                   objDomNode.Text <> enumTipoSolicitacao.Complementacao Then
                    'Tipo Solicitação inválido
                    fgAdicionaErro xmlErrosNegocio, 4016
                End If
            End If
            
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        Else
            'Tipo Solicitação inválido
            fgAdicionaErro xmlErrosNegocio, 4016
        End If

        'Data da Mensagem Invalida
        Set objDomNode = .selectSingleNode("//DT_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaDataNumerica(objDomNode.Text) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            ElseIf fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                'Data da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4060
            End If
        End If
        
        'Hora de envio
        Set objDomNode = .selectSingleNode("//HO_MESG")
        If Not objDomNode Is Nothing Then
            If Not flValidaHoraNumerica(fgCompletaString(objDomNode.Text, "0", 4, True)) Then
                'Hora da Mensagem inválida
                fgAdicionaErro xmlErrosNegocio, 4061
            End If
        End If

        'Veículo Legal
        If Not fgExisteVeiculoLegal(.selectSingleNode("//SG_SIST_ORIG").Text, _
                                    .selectSingleNode("//CO_VEIC_LEGA").Text, _
                                    datDataVigencia, _
                                    .selectSingleNode("//CO_EMPR").Text) Then
            'Veículo legal inexistente
            fgAdicionaErro xmlErrosNegocio, 4007
        End If
        
        'Data da operação
        Set objDomNode = .selectSingleNode("//DT_OPER_ATIV")
        If Not objDomNode Is Nothing Then
            'Data da Operação não pode ser maior que hoje
            If fgDtXML_To_Date(objDomNode.Text) <> flDataHoraServidor(enumFormatoDataHora.Data) Then
                'Data de movimento inválida
                fgAdicionaErro xmlErrosNegocio, 4012
            ElseIf Not fgDiaUtil(fgDtXML_To_Date(objDomNode.Text)) Then
                'Data da operação não é um dia útil
                fgAdicionaErro xmlErrosNegocio, 4059
            Else
                datDataMovimento = fgDtXML_To_Date(objDomNode.Text)
            End If
        End If

        Set objDomNode = .selectSingleNode("//TpOpComercExtr")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) <> vbNullString Then
                If UCase(Trim(objDomNode.Text)) <> "DR" And _
                   UCase(Trim(objDomNode.Text)) <> "ER" Then
                    'Tipo Operação Comércio Exterior inválido.
                    fgAdicionaErro xmlErrosNegocio, 4401
                End If
            End If
        Else
            'Tipo Operação Comércio Exterior inválido.
            fgAdicionaErro xmlErrosNegocio, 4401
        End If

        
        Set objDomNode = .selectSingleNode("//CodReemb")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) = vbNullString Then
                'Código Reembolso obrigatorio
                fgAdicionaErro xmlErrosNegocio, 4434
            End If
        Else
            'Código Reembolso obrigatorio
            fgAdicionaErro xmlErrosNegocio, 4434
        End If
                    
        
        Set objDomNode = .selectSingleNode("//TpInstntoCCR")
        If Not objDomNode Is Nothing Then
            If Trim(objDomNode.Text) = vbNullString Then
                'Tipo Instrumento CCR não informado.
                fgAdicionaErro xmlErrosNegocio, 4414
            Else
                If Not fgExisteDominioTagMensagemBACEN("TpInstntoCCR", Trim(objDomNode.Text)) Then
                    'Tipo Instrumento CCR inválido.
                    fgAdicionaErro xmlErrosNegocio, 4415
                End If
            End If
        Else
            'Tipo Instrumento CCR não informado.
            fgAdicionaErro xmlErrosNegocio, 4413
        End If
            
        'VlrDevRecolht_EstReemb
                
        
                
    End With
    
    fgConsisteDevolucaoEstornoCCR = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteDevolucaoEstornoCCR", 0


End Function


Public Function fgValidaValorNegociacaoCCR(ByVal pstrCodReemb As String, _
                                           ByVal pstrValorNecociacao As String, _
                                           ByVal strTipoInstrumento) As Boolean

Dim strSQL                                  As String
Dim strWhereTipoOper                        As String
Dim rsQuery                                 As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT SUM(VA_OPER_ATIV) AS VA_OPER_ATIV " & _
             " FROM   A8.TB_OPER_ATIV        " & vbCrLf & _
             " WHERE  NU_CTRL_MESG_SPB_ORIG = '" & pstrCodReemb & "'" & _
             " AND    CO_ULTI_SITU_PROC NOT IN (" & enumStatusOperacao.Cancelada & ", " & _
                                                    enumStatusOperacao.Excluida & ", " & _
                                                    enumStatusOperacao.Inconsistencia & ", " & _
                                                    enumStatusOperacao.Rejeitada & ", " & _
                                                    enumStatusOperacao.RejeitadaLiquidacao & ") "

    Set rsQuery = QuerySQL(strSQL)

    If Not rsQuery.EOF Then
    
        If strTipoInstrumento = "CC" Or strTipoInstrumento = "CD" Then
            If fgVlrXml_To_Decimal(pstrValorNecociacao) > (rsQuery!VA_OPER_ATIV * fgVlrXml_To_Decimal("1.1")) Then
                fgValidaValorNegociacaoCCR = False
            Else
                fgValidaValorNegociacaoCCR = True
            End If
        Else
            If fgVlrXml_To_Decimal(pstrValorNecociacao) > rsQuery!VA_OPER_ATIV Then
                fgValidaValorNegociacaoCCR = False
            Else
                fgValidaValorNegociacaoCCR = True
            End If
        End If
    Else
        fgValidaValorNegociacaoCCR = True
    End If
    
    Set rsQuery = Nothing

    Exit Function
ErrorHandler:

    Set rsQuery = Nothing

    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgValidaSituacaoOperacaoCCR", 0
End Function

' Retorna data e hora do servidor.

Public Function flDataHoraServidor(ByVal pFormato As enumFormatoDataHoraAux) As Date

Dim strSQL                                   As String
Dim objRS                                    As ADODB.Recordset
Dim strRetorno                               As String
Dim strAux                                   As String

On Error GoTo ErrorHandler

    strSQL = "select to_char(sysdate,'yyyymmddhh24miss')  as DH_SERV from dual"

    Set objRS = QuerySQL(strSQL)
    
    strRetorno = objRS!DH_SERV
    
    Select Case pFormato
        Case enumFormatoDataHoraAux.DataAux
            
            strAux = Mid(strRetorno, 1, 8)
            flDataHoraServidor = fgDtXML_To_Date(strAux)
            
        Case enumFormatoDataHoraAux.HoraAux
            
            flDataHoraServidor = fgDtHrXml_To_Time(strRetorno)
        
        Case enumFormatoDataHoraAux.DataHoraAux
            
            flDataHoraServidor = fgDtHrStr_To_DateTime(strRetorno)
            
    End Select

    
    Set objRS = Nothing
    
    Exit Function
ErrorHandler:
    Set objRS = Nothing
    
  
    
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "flDataHoraServidor", 0
End Function

Public Function fgConsisteOperacaoInterbancaria(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngTipoSolicitacao                      As Long
Dim lngTipoMensagem                         As Long
Dim intTipoNegociacaoInterbancaria          As Integer
Dim datDataVigencia                         As Date

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")

        lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("//TP_SOLI")
        If Not objDomNode Is Nothing Then
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If
        
        'Tipo de Negociação Interbancária
        Set objDomNode = .selectSingleNode("//TP_NEGO_INTB")
        If Not objDomNode Is Nothing Then
            intTipoNegociacaoInterbancaria = Val("0" & objDomNode.Text)
        End If

        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.SSTR And _
               Val("0" & objDomNode.Text) <> enumLocalLiquidacao.BMC And _
               Val("0" & objDomNode.Text) <> enumLocalLiquidacao.PAG Then
                'Código do local de liquidação inválido
                fgAdicionaErro xmlErrosNegocio, 4008
            End If
        End If
        
        'Código do Produto Moeda Estrangeira
        Set objDomNode = .selectSingleNode("//CO_PROD_MOED_ESTR")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("//CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4506
                End If
            Else
                'Produto inválido
                fgAdicionaErro xmlErrosNegocio, 4506
            End If
        End If
        
        'Valida Tipo Operacao Cambio
        If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or lngTipoSolicitacao = enumTipoSolicitacao.Confirmacao Then
            Set objDomNode = .selectSingleNode("//TP_OPER_CAMB")
            If Not objDomNode Is Nothing Then
                If objDomNode.Text = vbNullString Then
                    fgAdicionaErro xmlErrosNegocio, 4523
                End If
            Else
                fgAdicionaErro xmlErrosNegocio, 4523
            End If
        End If

        If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or lngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then

            If intTipoNegociacaoInterbancaria = enumTipoNegociacaoInterbancaria.SemCamara Then

                'Valida CNPJ IF Compradora
                Set objDomNode = .selectSingleNode("//CO_CNPJ_IF_COMPR")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        fgAdicionaErro xmlErrosNegocio, 4472
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4472
                End If

                'Valida CNPJ IF Vendedora
                Set objDomNode = .selectSingleNode("//CO_CNPJ_IF_VENC")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        fgAdicionaErro xmlErrosNegocio, 4473
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4473
                End If

                'Valida Indicador Giro
                Set objDomNode = .selectSingleNode("//IN_GIRO")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text = vbNullString Then
                        fgAdicionaErro xmlErrosNegocio, 4475
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4475
                End If

                'Valida Indicador Linha
                Set objDomNode = .selectSingleNode("//IN_LINHA")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text = vbNullString Then
                        fgAdicionaErro xmlErrosNegocio, 4476
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4476
                End If

                'Valida Codigo Fato Natureza
                Set objDomNode = .selectSingleNode("//CO_FATO_NATU")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text = vbNullString Then
                        fgAdicionaErro xmlErrosNegocio, 4477
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4477
                End If

                'Valida Codigo Cliente Natureza
                Set objDomNode = .selectSingleNode("//CO_CLIE_NATU")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text = vbNullString Then
                        fgAdicionaErro xmlErrosNegocio, 4478
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4478
                End If

                'Valida Indicador Aval Natureza
                Set objDomNode = .selectSingleNode("//IN_AVAL_NATU")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text = vbNullString Then
                        fgAdicionaErro xmlErrosNegocio, 4479
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4479
                End If

                'Valida Codigo Pagador ou Recebedor Exterior
                Set objDomNode = .selectSingleNode("//CO_PAGA_RECB_EXTE")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text = vbNullString Then
                        fgAdicionaErro xmlErrosNegocio, 4480
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4480
                End If
                
                'Valida Codigo Grupo Natureza
                Set objDomNode = .selectSingleNode("//CO_GRUP_NATU")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text = vbNullString Then
                        fgAdicionaErro xmlErrosNegocio, 4481
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4481
                End If
                
                'Valida Código Forma Entrega Moeda
                Set objDomNode = .selectSingleNode("//CO_FORM_ENTR_MOED")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        fgAdicionaErro xmlErrosNegocio, 4505
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4505
                End If

            ElseIf intTipoNegociacaoInterbancaria = enumTipoNegociacaoInterbancaria.InterbancarioEletronico Then

                'Valida CNPJ Base Camara
                Set objDomNode = .selectSingleNode("//CO_CNPJ_BASE_CAMR")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text = vbNullString Then
                        fgAdicionaErro xmlErrosNegocio, 4483
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4483
                End If
                
                'Valida Chave Associacao Cambio
                Set objDomNode = .selectSingleNode("//ChACAM")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text = vbNullString Then
                        fgAdicionaErro xmlErrosNegocio, 4484
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4484
                End If

                'Valida CNPJ IF
                Set objDomNode = .selectSingleNode("//CO_CNPJ_IF")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        fgAdicionaErro xmlErrosNegocio, 4485
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4485
                End If

            ElseIf intTipoNegociacaoInterbancaria = enumTipoNegociacaoInterbancaria.SemTelaCega Then

                'Valida CNPJ IF Compradora
                Set objDomNode = .selectSingleNode("//CO_CNPJ_IF_COMPR")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        fgAdicionaErro xmlErrosNegocio, 4472
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4472
                End If
                
                'Valida CNPJ IF Vendedora
                Set objDomNode = .selectSingleNode("//CO_CNPJ_IF_VENC")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        fgAdicionaErro xmlErrosNegocio, 4473
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4473
                End If

                'Valida CNPJ Camara
                Set objDomNode = .selectSingleNode("//CO_CNPJ_CAM")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        fgAdicionaErro xmlErrosNegocio, 4474
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4474
                End If
                
                'Valida Indicador Giro
                Set objDomNode = .selectSingleNode("//IN_GIRO")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text = vbNullString Then
                        fgAdicionaErro xmlErrosNegocio, 4475
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4475
                End If
                
                'Valida Indicador Linha
                Set objDomNode = .selectSingleNode("//IN_LINHA")
                If Not objDomNode Is Nothing Then
                    If objDomNode.Text = vbNullString Then
                        fgAdicionaErro xmlErrosNegocio, 4476
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4476
                End If

            End If
        End If

    End With
    
    fgConsisteOperacaoInterbancaria = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteOperacaoInterbancaria", 0
End Function

Public Function fgConsisteOperacaoArbitragem(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngTipoSolicitacao                      As Long
Dim lngTipoMensagem                         As Long
Dim intTipoNegociacaoArbitragem             As Integer
Dim datDataVigencia                         As Date

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    datDataVigencia = flDataHoraServidor(enumFormatoDataHora.Data)

    With xmlDOMMensagem.selectSingleNode("/MESG")

        lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)

        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.BMC Then
                'Código do local de liquidação inválido
                fgAdicionaErro xmlErrosNegocio, 4008
            End If
        End If

        'Tipo de solicitação
        Set objDomNode = .selectSingleNode("//TP_SOLI")
        If Not objDomNode Is Nothing Then
            lngTipoSolicitacao = Val("0" & objDomNode.Text)
        End If
        
        'Tipo de Negociação Arbitragem
        Set objDomNode = .selectSingleNode("//TP_NEGO_ARBT")
        If Not objDomNode Is Nothing Then
            intTipoNegociacaoArbitragem = Val("0" & objDomNode.Text)
        End If

        'Valida Codigo de Produto Moeda Estrangeira
        Set objDomNode = .selectSingleNode("//CO_PROD_MOED_ESTR")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> 0 Then
                If Not fgExisteProduto(.selectSingleNode("//CO_EMPR").Text, _
                                       objDomNode.Text, _
                                       datDataVigencia) Then
                    'Produto inválido
                    fgAdicionaErro xmlErrosNegocio, 4506
                End If
            Else
                'Produto inválido
                fgAdicionaErro xmlErrosNegocio, 4506
            End If
        End If

        If lngTipoSolicitacao = enumTipoSolicitacao.Complementacao Or lngTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then

            If intTipoNegociacaoArbitragem = enumTipoNegociacaoArbitragem.ParceiroPais Then
            
                'Valida CNPJ IF
                Set objDomNode = .selectSingleNode("//CO_CNPJ_IF")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        fgAdicionaErro xmlErrosNegocio, 4485
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4485
                End If
            
                'Valida CNPJ IF Parceira
                Set objDomNode = .selectSingleNode("//CO_CNPJ_IF_PARC")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        fgAdicionaErro xmlErrosNegocio, 4488
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4488
                End If
            
'                'Valida Numero Sequencia Instrucao Pagamento
'                Set objDomNode = .selectSingleNode("//NU_SEQU_INST_PAGTO")
'                If Not objDomNode Is Nothing Then
'                    If Val("0" & objDomNode.Text) = 0 Then
'                        fgAdicionaErro xmlErrosNegocio, 4489
'                    End If
'                Else
'                    fgAdicionaErro xmlErrosNegocio, 4489
'                End If
            
            End If
            
            'Valida Grupo Contratacao
            Set objDomNode = .selectSingleNode("//GR_CONTR")
            If Not objDomNode Is Nothing Then
                If objDomNode.childNodes.length = 0 Then
                    fgAdicionaErro xmlErrosNegocio, 4499
                End If
            Else
                fgAdicionaErro xmlErrosNegocio, 4499
            End If

            'Valida Valor Moeda Nacional
            Set objDomNode = .selectSingleNode("//VA_MOED_NACIO")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) = 0 Then
                    fgAdicionaErro xmlErrosNegocio, 4490
                End If
            Else
                fgAdicionaErro xmlErrosNegocio, 4490
            End If
            
            'Valida Data Liquidacao
            Set objDomNode = .selectSingleNode("//DT_LIQU_OPER")
            If Not objDomNode Is Nothing Then
                If Val("0" & objDomNode.Text) = 0 Then
                    fgAdicionaErro xmlErrosNegocio, 4491
                End If
            Else
                fgAdicionaErro xmlErrosNegocio, 4491
            End If

            'Valida Codigo Fato Natureza
            Set objDomNode = .selectSingleNode("//CO_FATO_NATU")
            If Not objDomNode Is Nothing Then
                If objDomNode.Text = vbNullString Then
                    fgAdicionaErro xmlErrosNegocio, 4492
                End If
            Else
                fgAdicionaErro xmlErrosNegocio, 4492
            End If

            'Valida Codigo Cliente Natureza
            Set objDomNode = .selectSingleNode("//CO_CLIE_NATU")
            If Not objDomNode Is Nothing Then
                If objDomNode.Text = vbNullString Then
                    fgAdicionaErro xmlErrosNegocio, 4493
                End If
            Else
                fgAdicionaErro xmlErrosNegocio, 4493
            End If

            'Valida Indicador Aval Natureza
            Set objDomNode = .selectSingleNode("//IN_AVAL_NATU")
            If Not objDomNode Is Nothing Then
                If objDomNode.Text = vbNullString Then
                    fgAdicionaErro xmlErrosNegocio, 4494
                End If
            Else
                fgAdicionaErro xmlErrosNegocio, 4494
            End If

            'Valida Codigo Pagador ou Recebedor Exterior
            Set objDomNode = .selectSingleNode("//CO_PAGA_RECB_EXTE")
            If Not objDomNode Is Nothing Then
                If objDomNode.Text = vbNullString Then
                    fgAdicionaErro xmlErrosNegocio, 4495
                End If
            Else
                fgAdicionaErro xmlErrosNegocio, 4495
            End If
            
            'Valida Codigo Grupo Natureza
            Set objDomNode = .selectSingleNode("//CO_GRUP_NATU")
            If Not objDomNode Is Nothing Then
                If objDomNode.Text = vbNullString Then
                    fgAdicionaErro xmlErrosNegocio, 4496
                End If
            Else
                fgAdicionaErro xmlErrosNegocio, 4496
            End If

        ElseIf lngTipoSolicitacao = enumTipoSolicitacao.Confirmacao Then

'            'Valida Numero Sequencia Instrucao Pagamento
'            Set objDomNode = .selectSingleNode("//NU_SEQU_INST_PAGTO")
'            If Not objDomNode Is Nothing Then
'                If Val("0" & objDomNode.Text) = 0 Then
'                    fgAdicionaErro xmlErrosNegocio, 4489
'                End If
'            Else
'                fgAdicionaErro xmlErrosNegocio, 4489
'            End If

        End If

    End With
    
    fgConsisteOperacaoArbitragem = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteOperacaoArbitragem", 0
End Function

Public Function fgConsisteOperacaoLiquidacaoInterbancaria(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngTipoMensagem                         As Long

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    With xmlDOMMensagem.selectSingleNode("/MESG")

        lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)

        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.SSTR And _
               Val("0" & objDomNode.Text) <> enumLocalLiquidacao.BMC And _
               Val("0" & objDomNode.Text) <> enumLocalLiquidacao.PAG Then
                'Código do local de liquidação inválido
                fgAdicionaErro xmlErrosNegocio, 4008
            End If
        End If

        'Código do Produto
        Set objDomNode = .selectSingleNode("//TP_NEGO_CAML")
        If Not objDomNode Is Nothing Then
            If objDomNode.Text = enumTipoNegociacaoCambial.Interbancaria Then
                Set objDomNode = .selectSingleNode("//CO_PROD")
                If Not objDomNode Is Nothing Then
                    If Val("0" & objDomNode.Text) = 0 Then
                        fgAdicionaErro xmlErrosNegocio, 4070 'Produto inválido
                    End If
                Else
                    fgAdicionaErro xmlErrosNegocio, 4070 'Produto inválido
                End If
            End If
        End If
        
    End With
    
    fgConsisteOperacaoLiquidacaoInterbancaria = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteOperacaoLiquidacaoInterbancaria", 0
End Function


Public Function fgConsisteOperacaoConsultaContratosMercadoInterbancario(ByRef xmlDOMMensagem As MSXML2.DOMDocument40) As String

Dim xmlErrosNegocio                         As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngTipoMensagem                         As Long

On Error GoTo ErrorHandler

    Set xmlErrosNegocio = CreateObject("MSXML2.DOMDocument.4.0")

    With xmlDOMMensagem.selectSingleNode("/MESG")

        lngTipoMensagem = Val("0" & .selectSingleNode("TP_MESG").Text)

        'Local de Liquidação
        Set objDomNode = .selectSingleNode("CO_LOCA_LIQU")
        If Not objDomNode Is Nothing Then
            If Val("0" & objDomNode.Text) <> enumLocalLiquidacao.BMC Then
                'Código do local de liquidação inválido
                fgAdicionaErro xmlErrosNegocio, 4008
            End If
        End If
    
    End With
    
    fgConsisteOperacaoConsultaContratosMercadoInterbancario = xmlErrosNegocio.xml

    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing

Exit Function
ErrorHandler:
    Set xmlErrosNegocio = Nothing
    Set objDomNode = Nothing
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "fgConsisteOperacaoConsultaContratosMercadoInterbancario", 0
End Function




