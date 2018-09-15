VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcesso_Batch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Controle Transacional de acesso ao MQSeries/Oracle, distribuição das Mensagens aos objetos responsáveis pelo tratamento.

Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Private datDataHoraUltimaExecucao           As Date
Private llTempoVerifSTR0008R2               As Long

Implements COMSVCSLib.ObjectControl

Private ObjectContext                       As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
End Sub
Public Function ProcessarAgendamento(ByVal pstrXmlParam)
Dim xmlParametro                            As MSXML2.DOMDocument40
Dim strRetorno                              As String
Dim blnProcessar                            As Boolean
Dim strNomeObjeto                           As String
Dim strNomeFila                             As String

On Error GoTo ErrorHandler

    blnProcessar = True

    If Not fgVerificarBancoDados Then
        blnProcessar = False
    Else
        If Not fgVerificaMQSeries("QM.SLCC.01") Then
            blnProcessar = False
        End If
    End If
    
    If Not blnProcessar Then Exit Function
    
    Set xmlParametro = CreateObject("MSXML2.DOMDocument.4.0")
    xmlParametro.loadXML pstrXmlParam
    
    strNomeFila = Trim(xmlParametro.selectSingleNode("//IdProcesso").Text)
    
    Select Case UCase(strNomeFila)
        Case "CONCILIACAOLOTE"
            Call ConciliarLote
    End Select
    
    Set xmlParametro = Nothing
    
Exit Function
ErrorHandler:
    
    Set xmlParametro = Nothing
    
    fgGravaArquivo App.EXEName & "." & "Processar", Err.Description
End Function

Public Function Processar(ByVal pstrXmlParam As String, _
                          ByVal strDataHoraUltiExecucao As String)

Dim xmlParametro                            As MSXML2.DOMDocument40
Dim strRetorno                              As String
Dim strNomeFila                             As String

Dim blnProcessar                            As Boolean

On Error GoTo ErrorHandler

    blnProcessar = True

    If Not fgVerificarBancoDados Then
        blnProcessar = False
    Else
        If Not fgVerificaMQSeries("QM.SLCC.01") Then
            blnProcessar = False
        End If
    End If
    
    If Not blnProcessar Then Exit Function

    datDataHoraUltimaExecucao = fgDtHrStr_To_DateTime(strDataHoraUltiExecucao)

    Set xmlParametro = CreateObject("MSXML2.DOMDocument.4.0")
    xmlParametro.loadXML pstrXmlParam
    
    strNomeFila = Trim(xmlParametro.selectSingleNode("//Fila").Text)
    
    Select Case UCase(strNomeFila)
        
        Case "A8B.RENOVA_CAHE"
        
            Call flRenovaCache
            
        Case "A8B.AGENDAMENTO"
            
            Call LiberarAgendamento
            
        Case "A8B.LIQUIDACAO_TQ"
            
            Call LiquidarOperacaoTQ
            
        Case "A8B.CTRL_USUA_SIST"
            
            Call VerificaUsuariosInativos
            
        Case "A8B.PROC_MESG_STR"
            
            Call ProcessaMensagemSTR0008R2Comti
            
        Case "A8B.LIQUIDACAO_FUTURA"
            
            Call GerenciarFluxoOperacoesLiquidacaoFutura
            
        Case "A8B.CTRL_ACESSO"
            
            Call RemoveUsuarioCache
        
        Case "A8B.REPROC_CC"
            
            Call flReprocessarCC
        
    End Select

    Set xmlParametro = Nothing

    Exit Function
    
ErrorHandler:
    
    Set xmlParametro = Nothing
    
    fgGravaArquivo App.EXEName & "." & "Processar", Err.Description
    

End Function
Private Function ConciliarLote() As String
Dim objGerenciador                          As Object
Dim strRetorno                              As String
Dim strLogErro                              As String

On Error GoTo ErrorHandler
    Set objGerenciador = CreateObject("A8LQS.clsConciliacaoLote")
    strRetorno = objGerenciador.LiberarLoteParaAlcada()
    Set objGerenciador = Nothing
    
    If strRetorno <> vbNullString Then
        fgGravaArquivo "ConciliarLote", strRetorno
    End If
    
    Exit Function
    
ErrorHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    fgGravaArquivo App.EXEName & "." & "ConciliarLote", strLogErro

End Function


'-----------------------------------------------------------------------
'Chamada a rotina de liberação de mensagens agendadas
'-----------------------------------------------------------------------

Private Function LiberarAgendamento() As String

Dim objGerenciador                          As Object 'A7Server.clsGerenciadorAgendamento
Dim strRetorno                              As String
Dim strLogErro                              As String

On Error GoTo ErrorHandler
    
    Set objGerenciador = CreateObject("A7Server.clsGerenciadorAgendamento")
    objGerenciador.LiberarAgendamento "A8LQS.clsTRemessa"
    Set objGerenciador = Nothing
    
    LiberarAgendamento = vbNullString
    
    Exit Function
    
ErrorHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    fgGravaArquivo App.EXEName & "." & "LiberarAgendamento", strLogErro
    
    
End Function

'-----------------------------------------------------------------------
'Chamada a rotina de liquidação com o sistema TQ
'-----------------------------------------------------------------------

Private Function LiquidarOperacaoTQ() As String

Dim objGerenciador                          As Object 'A7Server.clsGerenciadorAgendamento
Dim strRetorno                              As String
Dim strLogErro                              As String

On Error GoTo ErrorHandler
    
    Set objGerenciador = CreateObject("A8LQS.clsLiquidacaoTQ")
    strRetorno = objGerenciador.LiquidarOperacoesTQ
    Set objGerenciador = Nothing
    
    If strRetorno <> vbNullString Then
        fgGravaArquivo "LiquidacaoTQ", strRetorno
    End If
    
    Exit Function
    
ErrorHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    fgGravaArquivo App.EXEName & "." & "LiquidacaoTQ", strLogErro
    
End Function

'-----------------------------------------------------------------------
'Chamada a rotina de Processamento Mensagem STR0008R2 - Conti
'-----------------------------------------------------------------------

Private Function ProcessaMensagemSTR0008R2Comti() As String

Dim objGerenciador                          As Object 'A8LQS.clsProcessoMensagemSTR
Dim strRetorno                              As String
Dim strLogErro                              As String
Dim objFuncao                               As Object 'A6A7A8.clsA6A7A8Funcoes
Dim llTempoVerifSTR0008R2                   As Long

Dim datDataHoraUltimaExecucaoSTR            As Date
Dim xmlControleHora                         As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler
            
    Set objFuncao = CreateObject("A6A7A8.clsA6A7A8Funcoes")
    llTempoVerifSTR0008R2 = Val(objFuncao.ObterValorParametrosGerais("CONSULTA_STR0008/PERIODICIDADE_EM_MINUTOS"))
    Set objFuncao = Nothing
    
    If llTempoVerifSTR0008R2 = 0 Then llTempoVerifSTR0008R2 = 15
    
    Set xmlControleHora = CreateObject("MSXML2.DOMDocument.4.0")
        
    If Dir(App.Path & "\xmlSolicitacaoConsultaPZ.xml") = vbNullString Then
        Call fgAppendNode(xmlControleHora, "", "ControleHoraConsultaPZ", "")
        Call fgAppendNode(xmlControleHora, "ControleHoraConsultaPZ", "Data", fgDt_To_Xml(fgDataHoraServidor(DataAux)))
        Call fgAppendNode(xmlControleHora, "ControleHoraConsultaPZ", "UltimaConsulta", "000000")
        
        xmlControleHora.Save App.Path & "\xmlSolicitacaoConsultaPZ.xml"
    Else
        xmlControleHora.Load App.Path & "\xmlSolicitacaoConsultaPZ.xml"
        
        If xmlControleHora.selectSingleNode("//Data").Text <> fgDt_To_Xml(fgDataHoraServidor(DataAux)) Then
            xmlControleHora.selectSingleNode("//Data").Text = fgDt_To_Xml(fgDataHoraServidor(DataAux))
            xmlControleHora.selectSingleNode("//UltimaConsulta").Text = "000000"
        End If
    End If
    
    datDataHoraUltimaExecucaoSTR = fgDtHrStr_To_DateTime(xmlControleHora.selectSingleNode("//Data").Text & _
                                                         xmlControleHora.selectSingleNode("//UltimaConsulta").Text)
    
    If DateDiff("n", datDataHoraUltimaExecucaoSTR, fgDataHoraServidor(DataHoraAux)) >= llTempoVerifSTR0008R2 Then
        
        xmlControleHora.selectSingleNode("//UltimaConsulta").Text = Right$(fgDtHr_To_Xml(fgDataHoraServidor(DataHoraAux)), 6)
        xmlControleHora.Save App.Path & "\xmlSolicitacaoConsultaPZ.xml"
        
        Set objGerenciador = CreateObject("A8LQS.clsProcessoMensagemSTR")
        strRetorno = objGerenciador.VerificaMensagemSTR0008R2Comti()
        Set objGerenciador = Nothing
    
    End If
    
    If strRetorno <> vbNullString Then
        fgGravaArquivo "ProcessaMensagemSTR0008R2Comti", strRetorno
    End If
    
    Exit Function
    
ErrorHandler:
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    Set objFuncao = Nothing
    
    fgGravaArquivo App.EXEName & "." & "ProcessaMensagemSTR0008R2Comti", strLogErro
    
End Function

'-----------------------------------------------------------------------
'Chamada a rotina de Envio e-mail Relatorio Usuários Inativos - SLCC
'-----------------------------------------------------------------------

Private Function VerificaUsuariosInativos() As String

Dim objGerenciador                          As Object 'A7Server.clsControleUsuarioSistema
Dim strRetorno                              As String
Dim strLogErro                              As String
Dim datDataUltiExec                         As Date
  
On Error GoTo ErrorHandler
    
    datDataUltiExec = DateSerial(Year(datDataHoraUltimaExecucao), Month(datDataHoraUltimaExecucao), Day(datDataHoraUltimaExecucao))
    
    If datDataUltiExec <> fgDataHoraServidor(DataAux) Then
        Set objGerenciador = CreateObject("A7Server.clsControleUsuarioSistema")
        strRetorno = objGerenciador.VerificaUsuarioInativo
        Set objGerenciador = Nothing
    End If
    
    If strRetorno <> vbNullString Then
        fgGravaArquivo "VerificaUsuariosInativos", strRetorno
    End If
    
    Exit Function
    
ErrorHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    fgGravaArquivo App.EXEName & "." & "VerificaUsuariosInativos", strLogErro
    
End Function

'-----------------------------------------------------------------------
'Chamada a rotina de liquidação futura D+1 (CETIP)
'-----------------------------------------------------------------------

Private Function GerenciarFluxoOperacoesLiquidacaoFutura() As String

Dim objGerenciador                          As Object 'A8LQS.CLSLIQUIDACAOFUTURA
Dim strRetorno                              As String
Dim strLogErro                              As String

On Error GoTo ErrorHandler
    
    If DatePart("H", Now) > 5 Then
        Set objGerenciador = CreateObject("A8LQS.clsLiquidacaoFutura")
        strRetorno = objGerenciador.GerenciarFluxoOperacoesLiquidacaoFutura
        Set objGerenciador = Nothing
    End If
    
    If strRetorno <> vbNullString Then
        fgGravaArquivo "GerenciarFluxoOperacoesLiquidacaoFutura", strRetorno
    End If
    
    Exit Function
    
ErrorHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    fgGravaArquivo App.EXEName & "." & "GerenciarFluxoOperacoesLiquidacaoFutura", strLogErro
    
End Function

'-----------------------------------------------------------------------
'Chamada a rotina de conciliacao de Roda de Dolar
'-----------------------------------------------------------------------

Private Function GerenciarFluxoConciliacaoRodaDolar() As String

Dim objGerenciador                          As Object
Dim strRetorno                              As String
Dim strLogErro                              As String

On Error GoTo ErrorHandler
    
    Set objGerenciador = CreateObject("")
    strRetorno = objGerenciador.GerenciarFluxoConciliacaoRodaDolar
    Set objGerenciador = Nothing
    
    If strRetorno <> vbNullString Then
        fgGravaArquivo "GerenciarFluxoConciliacaoRodaDolar", strRetorno
    End If
    
    Exit Function
    
ErrorHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    fgGravaArquivo App.EXEName & "." & "GerenciarFluxoConciliacaoRodaDolar", strLogErro
    
End Function

'-----------------------------------------------------------------------
'Chamada RemoveUsuarioCache
'-----------------------------------------------------------------------

Private Function RemoveUsuarioCache() As String

Dim objFuncao                               As Object 'A6A7A8.clsA6A7A8Funcoes

Dim objValidaRemessa                        As Object 'A6A8ValidaRemessa.clsValidaRemessa
Dim objParam                                As Object

Dim strLogErro                              As String

On Error GoTo ErrorHandler
    
        
    If UCase(flVerificaCacheUsuario) = "S" Then
    
        Set objValidaRemessa = CreateObject("A6A8ValidaRemessa.clsValidaRemessa")
        Call objValidaRemessa.RemoveUsuariosCache
        Set objValidaRemessa = Nothing
                                
        Set objParam = CreateObject("A7Server.clsParametrosGerais")
        Call objParam.AtualizaParametroCacheAcesso
        Set objParam = Nothing
    
    End If
    
    Exit Function
    
ErrorHandler:
    
    Set objFuncao = Nothing
    Set objValidaRemessa = Nothing
    Set objParam = Nothing
    
    strLogErro = Err.Description
    
    
    fgGravaArquivo App.EXEName & "." & "RemoveUsuarioCache", strLogErro
    
End Function

Private Function flVerificaCacheUsuario() As String

Dim objFuncao                               As Object 'A6A7A8.clsA6A7A8Funcoes

On Error GoTo ErrorHandler
            
    Set objFuncao = CreateObject("A6A7A8.clsA6A7A8Funcoes")
    flVerificaCacheUsuario = objFuncao.ObterValorParametrosGerais("CACHE_CONTROLE_ACESSO/PERIODICIDADE_EM_MINUTOS")
    Set objFuncao = Nothing

    Exit Function

ErrorHandler:
    flVerificaCacheUsuario = "N"
    Err.Clear

End Function

Private Sub flRenovaCache()

Dim objValidaRemessa                        As Object 'A6A8ValidaRemessa.clsValidaRemessa
Dim datDataUltiExec                         As Date

  
On Error GoTo ErrorHandler
    
    datDataUltiExec = DateSerial(Year(datDataHoraUltimaExecucao), Month(datDataHoraUltimaExecucao), Day(datDataHoraUltimaExecucao))
    
    If datDataUltiExec <> fgDataHoraServidor(DataAux) Then
        Set objValidaRemessa = CreateObject("A6A8ValidaRemessa.clsValidaRemessa")
        Call objValidaRemessa.RecarregarCache
        Set objValidaRemessa = Nothing
    End If
    
    Exit Sub

ErrorHandler:
    Set objValidaRemessa = Nothing
    
    Err.Clear

End Sub


Private Sub flReprocessarCC()

Dim objGerenciadorReprocCC                  As Object
Dim strLogErro                              As String

On Error GoTo ErrorHandler
    
    Set objGerenciadorReprocCC = CreateObject("A8LQS.clsGerenciadorReprocCC")
    Call objGerenciadorReprocCC.Processar
    Set objGerenciadorReprocCC = Nothing
    
    Exit Sub

ErrorHandler:
    strLogErro = Err.Description
    
    Set objGerenciadorReprocCC = Nothing
    
    fgGravaArquivo App.EXEName & "." & "flReprocessarCC", strLogErro
    

End Sub

