VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMovimentoSubReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA\r\nPacote         :\r\nClasse         : clsMovimentoSubReserva\r\nData Criação   : 14/07/2003\r\nObjetivo       :\r\n\r\nAnalista       : Carlos Fortes\r\n\r\nProgramador    : Carlos Fortes\r\nData           : 14/07/2003\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F105A9300D3"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela consulta e atualização
' do Cadastro de Movimento Sub-reserva.

Option Explicit

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

Implements COMSVCSLib.ObjectControl
Private objContext                           As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

' Este método é responsável pela atualização geral da tabela de movimento sub-reserva,
' e pelo retorno à camada intermediária.

Public Function Executar(ByRef pxmlMoviSubRes As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlMoviSubRes.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Excluir"
            Executar = Salvar(pxmlMoviSubRes)
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
    End Select

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        ''objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Apresenta o movimento de acordo com a posição ATUAL do caixa SubReserva,
' independente se este estiver ABERTO, FECHADO ou DISPONÍVEL

Public Function ObterMovimentoSubReserva(ByVal pxmlFiltros As MSXML2.DOMDocument40) As String

Dim xmlRetorno                              As MSXML2.DOMDocument40
Dim xmlPosicaoCaixaSubReserva               As MSXML2.DOMDocument40
Dim pxmlDomNode                             As MSXML2.IXMLDOMNode
Dim strSQL                                  As String
Dim objPosicaoCaixaSubReserva               As A6SubReserva.clsCaixaSubReserva

Dim strPosicaoCaixaSubReserva               As String
Dim lngEmpresa                              As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String
Dim intNivelAgrupamento                     As Integer
Dim strMovimentoCaixaSubReserva             As String

Dim strAgrupItemCaixa                       As String
Dim intCount                                As Integer
Dim blnRetornaPosicaoDisponivel             As Boolean

Dim strDataCaixaInformada                   As String
Dim strDataCaixaSubReserva                  As String
Dim strValorAberturaCaixaSubReserva         As String

Dim intTipoBackOffice                       As Integer

On Error GoTo ErrorHandler
    
    blnRetornaPosicaoDisponivel = True
    
    'Captura a Empresa, Veículo Legal, Sistema e Nível de Agrupamento do filtro
    If Not pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_VeiculoLegal/BancoLiquidante") Is Nothing Then
        lngEmpresa = pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_VeiculoLegal/BancoLiquidante").Text
    End If
    
    strCodigoVeiculoLegal = pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_VeiculoLegal/VeiculoLegal").Text
    strSiglaSistema = pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_VeiculoLegal/Sistema").Text
    intNivelAgrupamento = Val(pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_VeiculoLegal/NivelAgrupamento").Text)
    
    If Not pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_VeiculoLegal/DataCaixa") Is Nothing Then
        strDataCaixaInformada = pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_VeiculoLegal/DataCaixa").Text
        
        'Extrai apenas a data, pois a data informada vem formatada << TO_DATE('YYYYMMDD','YYYYMMDD') >>
        strDataCaixaInformada = Mid(strDataCaixaInformada, 10, 8)
    End If
    
    'Captura a Posição do Caixa SubReserva
    Set objPosicaoCaixaSubReserva = CreateObject("A6SubReserva.clsCaixaSubReserva")
    
    'Se a data do caixa para a consulta for informada, então...
    If strDataCaixaInformada <> vbNullString Then
        Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
        
        '...verifica a posição atual do caixa...
        Call xmlRetorno.loadXML(objPosicaoCaixaSubReserva.ObterPosicaoCaixaSubReserva(strCodigoVeiculoLegal, _
                                                                                      strSiglaSistema))
        
        '...em seguida, verifica se a data informada para a consulta é MENOR que a data ATUAL do caixa...
        If Val(strDataCaixaInformada) < Val(xmlRetorno.selectSingleNode("//DT_CAIX_SUB_RESE_ATUAL").Text) Then
            '...se sim, então se a posição para a data de consulta NÃO FOR ENCONTRADA,
            '   NÃO RETORNAR POSIÇÃO DISPONÍVEL
            blnRetornaPosicaoDisponivel = False
        End If
    End If
    
    strPosicaoCaixaSubReserva = objPosicaoCaixaSubReserva.ObterPosicaoCaixaSubReserva(strCodigoVeiculoLegal, _
                                                                                      strSiglaSistema, _
                                                                                      strDataCaixaInformada, _
                                                                                      blnRetornaPosicaoDisponivel)
    Set objPosicaoCaixaSubReserva = Nothing
    
    If strPosicaoCaixaSubReserva <> vbNullString Then
        'Captura Data do Caixa SubReserva e Valor de Abertura do Caixa
        Set xmlPosicaoCaixaSubReserva = CreateObject("MSXML2.DOMDocument.4.0")
        
        Call xmlPosicaoCaixaSubReserva.loadXML(strPosicaoCaixaSubReserva)
        
        For Each pxmlDomNode In xmlPosicaoCaixaSubReserva.documentElement.childNodes
            If strDataCaixaInformada = vbNullString Then
                strDataCaixaSubReserva = pxmlDomNode.selectSingleNode("DT_CAIX_SUB_RESE").Text
            Else
                strDataCaixaSubReserva = strDataCaixaInformada
            End If
            
            strValorAberturaCaixaSubReserva = Replace(pxmlDomNode.selectSingleNode("VA_UTLZ_ABER_CAIX_SUB_RESE").Text, _
                                                        ",", ".")
        Next
    Else
        strDataCaixaSubReserva = strDataCaixaInformada
    End If
    
    'Monta o agrupamento de acordo com o número de níveis enviado
    strAgrupItemCaixa = vbNullString
    For intCount = 1 To intNivelAgrupamento
        strAgrupItemCaixa = strAgrupItemCaixa & _
                            "|| A.CO_ITEM_CAIX_NIVE_" & Format$(intCount, "00") & " "
    Next
    
    strAgrupItemCaixa = Mid$(strAgrupItemCaixa, 4)
    
    'Seleciona o movimento a partir da Data do Caixa SubReserva
    strSQL = " Select   " & strAgrupItemCaixa & "        CO_ITEM_CAIX, " & vbNewLine & _
            "         ((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 1, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) - " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 1, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0)))) - " & vbNewLine & _
            "         ((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 4, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) - " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 4, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0)))) As VA_PREV,  " & vbNewLine & _
            "         ((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) - " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0)))) - " & vbNewLine & _
            "         ((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 5, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 5, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0)))) As VA_SOLI, " & vbNewLine & _
            "         ((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0)))) - " & vbNewLine & _
            "         ((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 6, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 6, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0)))) As VA_CONF, " & vbNewLine

    strSQL = strSQL & _
            "        (((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0)))) - " & vbNewLine & _
            "        (((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 5, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 5, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0)))) + " & vbNewLine & _
            "         ((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0))))) - " & vbNewLine & _
            "         ((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 6, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 6, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0))))) As VA_TOT_MOV, " & vbNewLine

    strSQL = strSQL & _
            "       ((((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0)))) - " & vbNewLine & _
            "       ((((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 5, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 5, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0)))) + " & vbNewLine & _
            "         ((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0)))))) - " & vbNewLine & _
            "         ((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 6, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 6, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0)))))) - " & vbNewLine & _
            "        (((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 1, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) - " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 1, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0))))) - " & vbNewLine & _
            "        (((SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 4, DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0),0)) - " & vbNewLine & _
            "           SUM(DECODE(A.CO_SITU_MOVI_CAIX_SUB_RESE, 4, DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0),0))))) As VA_VAR_MOV " & vbNewLine
    
    strSQL = strSQL & _
            " From      A6.TB_MOVI_CAIX_SUB_RESE        A," & vbNewLine & _
            "           A8.TB_VEIC_LEGA                 B " & vbNewLine & _
            " Where     A.CO_VEIC_LEGA                  = B.CO_VEIC_LEGA" & vbNewLine & _
            " And       A.SG_SIST                       = B.SG_SIST" & vbNewLine & _
            " And       A.DH_MOVI_CAIX_SUB_RESE         BETWEEN TO_DATE('" & strDataCaixaSubReserva & "000000', 'YYYYMMDDHH24MISS')" & vbNewLine & _
            "                                           And     TO_DATE('" & strDataCaixaSubReserva & "235959', 'YYYYMMDDHH24MISS')" & vbNewLine & _
            " And       B.CO_VEIC_LEGA                  = '" & strCodigoVeiculoLegal & "'" & vbNewLine & _
            " And       B.SG_SIST                       = '" & strSiglaSistema & "'" & vbNewLine
    
    If lngEmpresa <> 0 Then
        strSQL = strSQL & _
                 " And       A.CO_EMPR = " & lngEmpresa & vbNewLine
    End If
    
    intTipoBackOffice = 99
    If Not pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_VeiculoLegal/BackOfficePerfilGeral") Is Nothing Then
        intTipoBackOffice = pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_VeiculoLegal/BackOfficePerfilGeral").Text
    End If
    
    Select Case intTipoBackOffice
        Case 99
            strSQL = strSQL & _
                     fgSegregaDados("A8.TB_VEIC_LEGA", _
                                    False, "B", "B", False, False, True, False, True, False)
        Case Is > 0
            strSQL = strSQL & _
                     vbNewLine & " AND B.TP_BKOF = " & intTipoBackOffice & vbNewLine
    End Select

    strSQL = strSQL & _
            " Group By  " & strAgrupItemCaixa & vbNewLine & _
            " Order By  " & strAgrupItemCaixa & " ASC"

    strMovimentoCaixaSubReserva = fgQueryXMLLerTodos("SubReservaD0", strSQL, App.EXEName & "." & TypeName(Me))

    'Retorna um XML contendo:
    '       1ª seção        -> Posição do Caixa SubReserva
    '       2ª seção        -> Movimento do Caixa SubReserva
    Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlRetorno, "", "PosicaoMovimentoSubReserva", "")
    
    If Not xmlPosicaoCaixaSubReserva Is Nothing Then
        If Trim(xmlPosicaoCaixaSubReserva.xml) <> vbNullString Then
            Call fgAppendXML(xmlRetorno, "PosicaoMovimentoSubReserva", xmlPosicaoCaixaSubReserva.xml)
        End If
    End If
        
    If Trim(strMovimentoCaixaSubReserva) <> vbNullString Then
        Call fgAppendXML(xmlRetorno, "PosicaoMovimentoSubReserva", strMovimentoCaixaSubReserva)
    End If
    
    ObterMovimentoSubReserva = xmlRetorno.xml
    
    Set xmlRetorno = Nothing
    Set xmlPosicaoCaixaSubReserva = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set xmlRetorno = Nothing
    Set objPosicaoCaixaSubReserva = Nothing
    Set xmlPosicaoCaixaSubReserva = Nothing
    
    If Not objContext Is Nothing Then
        ''objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarMapaNavegacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura de propriedades da tabela de movimento sub-reserva,
' e pelo retorno à camada intermediária.

Public Function ObterPropriedades() As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_MOVI_CAIX_SUB_RESE, " & vbCrLf & _
             " CO_EMPR, " & vbCrLf & _
             " TP_OPER, " & vbCrLf & _
             " DH_MOVI_CAIX_SUB_RESE, " & vbCrLf & _
             " CO_LOCA_LIQU, " & vbCrLf & _
             " SG_SIST, " & vbCrLf & _
             " CO_VEIC_LEGA, " & vbCrLf & _
             " CO_ITEM_CAIX, " & vbCrLf & _
             " DE_TIPO_LIQU, " & vbCrLf & _
             " VA_MOVI_CAIX_SUB_RESE, " & vbCrLf & _
             " IN_MOVI_ENTR_SAID, " & vbCrLf & _
             " CO_CNPJ_CNPT, " & vbCrLf & _
             " NO_CNPT, " & vbCrLf & _
             " DE_ATIV, " & vbCrLf & _
             " NU_SEQU_OPER_ATIV, " & vbCrLf & _
             " CO_SITU_MOVI_CAIX_SUB_RESE, " & vbCrLf & _
             " TP_GERA_MOVI, " & vbCrLf & _
             " TX_MOTI_GERA_MOVI, " & vbCrLf & _
             " CO_ITEM_CAIX_NIVE_01, CO_ITEM_CAIX_NIVE_02, " & vbCrLf & _
             " CO_ITEM_CAIX_NIVE_03, CO_ITEM_CAIX_NIVE_04, " & vbCrLf & _
             " CO_ITEM_CAIX_NIVE_05, " & vbCrLf & _
             " CO_USUA_ULTI_ATLZ, " & vbCrLf & _
             " CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
             " DH_ULTI_ATLZ " & vbCrLf
    strSQL = strSQL & " FROM  A6.TB_MOVI_CAIX_SUB_RESE "

    ObterPropriedades = fgPropriedades("Grupo_MovimentoSubReserva", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        ''objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura de detalhes do movimento sub-reserva,
' e pelo retorno à camada intermediária.

Public Function ObterDetalheMovimento(ByVal pxmlFiltros As MSXML2.DOMDocument40) As String

Dim strSQL                                  As String
Dim lngEmpresa                              As Long
Dim strDataMov                              As String
Dim strItemCaixa                            As String
Dim strCodVeiculoLegal                      As String
Dim strSiglaSistema                         As String

    On Error GoTo ErrorHandler

    'Captura filtro informado (Empresa, Data do Movimento, Item de Caixa e Veículo Legal)
    If Not pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_DetalheMovimento/BancoLiquidante") Is Nothing Then
        lngEmpresa = pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_DetalheMovimento/BancoLiquidante").Text
    End If
    
    strDataMov = pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_DetalheMovimento/DataMovimento").Text
    strItemCaixa = pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_DetalheMovimento/ItemCaixa").Text
    strCodVeiculoLegal = pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_DetalheMovimento/VeiculoLegal").Text
    strSiglaSistema = pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_DetalheMovimento/Sistema").Text
    
    strSQL = " SELECT   A.SG_SIST, B.NO_REDU_EMPR, C.SG_LOCA_LIQU, A.DE_TIPO_LIQU,               " & vbNewLine & _
             "          DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_MOVI_CAIX_SUB_RESE, 0) AS ENTRADA,   " & vbNewLine & _
             "          DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_MOVI_CAIX_SUB_RESE, 0) AS SAIDA,     " & vbNewLine & _
             "          A.DE_ATIV, A.CO_CNPJ_CNPT, A.NO_CNPT, A.CO_SITU_MOVI_CAIX_SUB_RESE,      " & vbNewLine & _
             "          A.DH_MOVI_CAIX_SUB_RESE, A.DT_RETN_OPER                                  " & vbNewLine & _
             " FROM     A6.TB_MOVI_CAIX_SUB_RESE        A,                                       " & vbNewLine & _
             "          A8.TB_EMPRESA_HO                B,                                       " & vbNewLine & _
             "          A8.TB_LOCAL_LIQUIDACAO          C                                        " & vbNewLine & _
             " WHERE    A.CO_EMPR      = B.CO_EMPR                                               " & vbNewLine & _
             " AND      A.CO_LOCA_LIQU = C.CO_LOCA_LIQU                                          " & vbNewLine
             
    If lngEmpresa <> 0 Then
        strSQL = strSQL & _
             " AND      A.CO_EMPR      = " & lngEmpresa & "                                      " & vbNewLine
    End If
    
    strSQL = strSQL & _
             " AND      A.CO_VEIC_LEGA = '" & strCodVeiculoLegal & "'                            " & vbNewLine & _
             " AND      A.SG_SIST      = '" & strSiglaSistema & "'                               " & vbNewLine & _
             " AND      A.CO_ITEM_CAIX = '" & strItemCaixa & "'                                  " & vbNewLine & _
             " AND      C.CO_EMPR_FUSI = 1                                                       " & vbNewLine & _
             " AND      A.DH_MOVI_CAIX_SUB_RESE BETWEEN TO_DATE('" & strDataMov & "000000', 'YYYYMMDDHH24MISS') " & vbNewLine & _
             "                                  AND     TO_DATE('" & strDataMov & "235959', 'YYYYMMDDHH24MISS') " & vbNewLine & _
             " ORDER BY A.DH_MOVI_CAIX_SUB_RESE, A.DH_ULTI_ATLZ DESC "
    
    'Verificação da Base histórica
    If fgDtXML_To_Date(strDataMov) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        strSQL = Replace(strSQL, "A6.TB_MOVI_CAIX_SUB_RESE", "A6HIST.TB_MOVI_CAIX_SUB_RESE")
    End If
    
    ObterDetalheMovimento = fgQueryXMLLerTodos("DetalheMovimento", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        ''objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheMovimento Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Este método é responsável pela leitura do total de movimento por veículo legal,
' e pelo retorno à camada intermediária.

Public Function ObterTotalPorVeiculoLegal(ByVal pstrCodVeiculoLegal As String, _
                                          ByVal pstrSiglaSistema As String) As String

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String

    On Error GoTo ErrorHandler

    'Atualiza a tabela << Posição do Caixa Disponível >>, somente se NÃO existir posição para a chave:
    'Véiculo Legal + Sigla Sistema
    Call flAtualizaPosicaoDisponivel(, pstrCodVeiculoLegal, pstrSiglaSistema)
    
    strWhereSegregacao = fgSegregaDados("A8.TB_VEIC_LEGA", _
                                        False, _
                                        "B", _
                                        "B", _
                                        False, _
                                        False, _
                                        True, _
                                        False, _
                                        True, _
                                        False)
    
    strSQL = " SELECT   C.CO_LOCA_LIQU, C.CO_VEIC_LEGA, D.SG_LOCA_LIQU, C.SG_SIST,                                         " & vbNewLine & _
             "          NVL(SUM(DECODE(C.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(C.IN_MOVI_ENTR_SAID, 1, C.VA_MOVI_CAIX_SUB_RESE, 0), 0)), 0) - " & vbNewLine & _
             "          NVL(SUM(DECODE(C.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(C.IN_MOVI_ENTR_SAID, 2, C.VA_MOVI_CAIX_SUB_RESE, 0), 0)), 0) + " & vbNewLine & _
             "          NVL(SUM(DECODE(C.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(C.IN_MOVI_ENTR_SAID, 1, C.VA_MOVI_CAIX_SUB_RESE, 0), 0)), 0) - " & vbNewLine & _
             "          NVL(SUM(DECODE(C.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(C.IN_MOVI_ENTR_SAID, 2, C.VA_MOVI_CAIX_SUB_RESE, 0), 0)), 0) AS VA_MOVI_CAIX_SUB_RESE " & vbNewLine
             
    strSQL = strSQL & _
             " FROM     A6.TB_POSI_CAIX_DISP           A,                               " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                B,                               " & vbNewLine & _
             "          A6.TB_MOVI_CAIX_SUB_RESE       C,                               " & vbNewLine & _
             "          A8.TB_LOCAL_LIQUIDACAO            D                                " & vbNewLine & _
             " WHERE    A.SG_SIST                      = B.SG_SIST                      " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA                 = B.CO_VEIC_LEGA                 " & vbNewLine & _
             " AND      A.SG_SIST                      = C.SG_SIST                      " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA                 = C.CO_VEIC_LEGA                 " & vbNewLine & _
             " AND      A.DT_CAIX_DISP                 = C.DH_MOVI_CAIX_SUB_RESE        " & vbNewLine & _
             " AND      C.CO_LOCA_LIQU                 = D.CO_LOCA_LIQU                 " & vbNewLine & _
             " AND      D.CO_EMPR_FUSI                 = 1                              " & vbNewLine & _
             " AND      B.CO_VEIC_LEGA                 = '" & pstrCodVeiculoLegal & "'  " & vbNewLine & _
             " AND      B.SG_SIST                      = '" & pstrSiglaSistema & "'     " & vbNewLine

    If Trim(strWhereSegregacao) <> vbNullString Then
        strSQL = strSQL & _
            strWhereSegregacao & vbNewLine
    End If

    strSQL = strSQL & _
             " GROUP BY C.CO_LOCA_LIQU, C.CO_VEIC_LEGA, D.SG_LOCA_LIQU, C.SG_SIST       " & vbNewLine & _
             " ORDER BY C.CO_LOCA_LIQU, C.CO_VEIC_LEGA, C.SG_SIST, D.SG_LOCA_LIQU "

    ObterTotalPorVeiculoLegal = fgQueryXMLLerTodos("TotalVeiculoLegal", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTotalPorVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura do total de movimento por grupo de veículo legal,
' e pelo retorno à camada intermediária.

Public Function ObterTotalPorGrupoVeiculoLegal(ByVal pintCO_GRUP_VEIC_LEGA As Integer) As String

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String

    On Error GoTo ErrorHandler

    'Atualiza a tabela << Posição do Caixa Disponível >>, somente se NÃO existir posição para a chave:
    'Véiculo Legal + Sigla Sistema
    Call flAtualizaPosicaoDisponivel(pintCO_GRUP_VEIC_LEGA)
    
    strWhereSegregacao = fgSegregaDados("A8.TB_VEIC_LEGA", _
                                        False, _
                                        "B", _
                                        "B", _
                                        False, _
                                        False, _
                                        True, _
                                        False, _
                                        True, _
                                        False)
    
    strSQL = " SELECT   C.CO_LOCA_LIQU, D.SG_LOCA_LIQU, B.CO_GRUP_VEIC_LEGA,                                               " & vbNewLine & _
             "          NVL(SUM(DECODE(C.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(C.IN_MOVI_ENTR_SAID, 1, C.VA_MOVI_CAIX_SUB_RESE, 0), 0)), 0) - " & vbNewLine & _
             "          NVL(SUM(DECODE(C.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(C.IN_MOVI_ENTR_SAID, 2, C.VA_MOVI_CAIX_SUB_RESE, 0), 0)), 0) + " & vbNewLine & _
             "          NVL(SUM(DECODE(C.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(C.IN_MOVI_ENTR_SAID, 1, C.VA_MOVI_CAIX_SUB_RESE, 0), 0)), 0) - " & vbNewLine & _
             "          NVL(SUM(DECODE(C.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(C.IN_MOVI_ENTR_SAID, 2, C.VA_MOVI_CAIX_SUB_RESE, 0), 0)), 0) AS VA_MOVI_CAIX_SUB_RESE " & vbNewLine
             
    strSQL = strSQL & _
             " FROM     A6.TB_POSI_CAIX_DISP           A,                          " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                B,                          " & vbNewLine & _
             "          A6.TB_MOVI_CAIX_SUB_RESE       C,                          " & vbNewLine & _
             "          A8.TB_LOCAL_LIQUIDACAO            D                           " & vbNewLine & _
             " WHERE    A.SG_SIST                      = B.SG_SIST                 " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA                 = B.CO_VEIC_LEGA            " & vbNewLine & _
             " AND      A.SG_SIST                      = C.SG_SIST                 " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA                 = C.CO_VEIC_LEGA            " & vbNewLine & _
             " AND      A.DT_CAIX_DISP                 = C.DH_MOVI_CAIX_SUB_RESE   " & vbNewLine & _
             " AND      C.CO_LOCA_LIQU                 = D.CO_LOCA_LIQU            " & vbNewLine & _
             " AND      D.CO_EMPR_FUSI                 = 1                         " & vbNewLine & _
             " AND      B.CO_GRUP_VEIC_LEGA            = " & pintCO_GRUP_VEIC_LEGA & vbNewLine

    If Trim(strWhereSegregacao) <> vbNullString Then
        strSQL = strSQL & _
            strWhereSegregacao & vbNewLine
    End If

    strSQL = strSQL & _
             " GROUP BY C.CO_LOCA_LIQU, D.SG_LOCA_LIQU, B.CO_GRUP_VEIC_LEGA                                                " & vbNewLine & _
             " ORDER BY C.CO_LOCA_LIQU, B.CO_GRUP_VEIC_LEGA, D.SG_LOCA_LIQU  "
    
    ObterTotalPorGrupoVeiculoLegal = fgQueryXMLLerTodos("TotalGrupoVeiculoLegal", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTotalPorGrupoVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela definição da instrução sql, de acordo com o tipo de remessa a ser incluída.

Private Function flDefinirStrValues(ByVal pxmlMoviSubRes As MSXML2.DOMDocument40) As String
Dim strValues                               As String
Dim intTipoEvento                           As Integer
Dim strItemCaix                             As String
Dim lngCoMoviCaixSubRese                    As Long
Dim strDataMovimento                        As String
Dim strCNPJContraparte                      As String
Dim strNomeContraparte                      As String
Dim strAtivo                                As String
Dim strTipoLiqu                             As String

    With pxmlMoviSubRes.documentElement
        lngCoMoviCaixSubRese = fgExecuteSequence("A6.SQ_A6_CO_MOVI_CAIX_SUB_RESE")
        
        If Not .selectSingleNode("//TP_MESG") Is Nothing Then
            intTipoEvento = CInt(.selectSingleNode("TP_MESG").Text)
        Else
            intTipoEvento = 0
        End If
        
        If intTipoEvento = enumEvento.DetalheOperacaoCaixaFuturo Then
            strItemCaix = .selectSingleNode("CO_ITEM_CAIX_SUB_RESE").Text
            
            strValues = lngCoMoviCaixSubRese & ",                                  " & vbNewLine & _
                        .selectSingleNode("CO_EMPR").Text & ",                     " & vbNewLine & _
                        IIf(.selectSingleNode("TP_OPER").Text = vbNullString, "               NULL", .selectSingleNode("TP_OPER").Text) & ", " & vbNewLine & _
                        fgDtXML_To_Oracle(.selectSingleNode("DT_REME").Text) & ",  " & vbNewLine & _
                        IIf(.selectSingleNode("CO_LOCA_LIQU").Text = vbNullString, "          NULL", .selectSingleNode("CO_LOCA_LIQU").Text) & ", " & vbNewLine & _
                        " '" & .selectSingleNode("SG_SIST").Text & "',             " & vbNewLine & _
                        " '" & .selectSingleNode("//CO_VEIC_LEGA").Text & "',        " & vbNewLine & _
                        " '" & strItemCaix & "',                                   " & vbNewLine & _
                        IIf(.selectSingleNode("DE_TIPO_LIQU").Text = vbNullString, "          NULL", " '" & .selectSingleNode("DE_TIPO_LIQU").Text & "' ") & ", " & vbNewLine & _
                        fgVlrToDBServer(.selectSingleNode("VA_LIQU_OPER").Text) & ",                " & vbNewLine & _
                        .selectSingleNode("IN_MOVI_ENTR_SAID").Text & ",           " & vbNewLine & _
                        IIf(.selectSingleNode("CO_CNPJ_CNPT").Text = vbNullString, "          NULL", .selectSingleNode("CO_CNPJ_CNPT").Text) & ", " & vbNewLine & _
                        IIf(.selectSingleNode("NO_CNPT").Text = vbNullString, "               NULL", " '" & .selectSingleNode("NO_CNPT").Text & "' ") & ", " & vbNewLine & _
                        " '" & .selectSingleNode("DE_ATIV_MERC").Text & "',             " & vbNewLine & _
                        IIf(.selectSingleNode("NU_SEQU_OPER_ATIV").Text = vbNullString, "     NULL", .selectSingleNode("NU_SEQU_OPER_ATIV").Text) & ", " & vbNewLine & _
                        .selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text & ",  " & vbNewLine & _
                        .selectSingleNode("TP_GERA_MOVI").Text & ",                " & vbNewLine & _
                        IIf(.selectSingleNode("TX_MOTI_GERA_MOVI").Text = vbNullString, "     NULL", " '" & .selectSingleNode("TX_MOTI_GERA_MOVI").Text & "' ") & ", " & vbNewLine & _
                        " '" & Mid$(strItemCaix, 2, 3) & "', '" & Mid$(strItemCaix, 5, 3) & "', '" & Mid$(strItemCaix, 8, 3) & "', '" & Mid$(strItemCaix, 11, 3) & "', '" & Mid$(strItemCaix, 14, 3) & "', " & vbNewLine & _
                        " '" & fgUsuarioRede & "',                                 " & vbNewLine & _
                        " '" & fgEstacaoTrabalhoUsuario & "',                      " & vbNewLine & _
                        " SYSDATE)"
        
        ElseIf intTipoEvento = enumEvento.MovimentoSubReserva Then
            strItemCaix = .selectSingleNode("CO_ITEM_CAIX_SUB_RESE").Text
            
            If .selectSingleNode("DT_OPER_ATIV") Is Nothing Then
                If .selectSingleNode("DT_LIQU_OPER_ATIV") Is Nothing Then
                    strDataMovimento = .selectSingleNode("DT_MESG").Text
                Else
                    strDataMovimento = .selectSingleNode("DT_LIQU_OPER_ATIV").Text
                End If
            Else
                strDataMovimento = .selectSingleNode("DT_OPER_ATIV").Text
            End If
                
            If .selectSingleNode("CO_CNPJ_CNPT") Is Nothing Then
                strCNPJContraparte = "NULL"
            Else
                strCNPJContraparte = "'" & .selectSingleNode("CO_CNPJ_CNPT").Text & "'"
            End If
            
            If .selectSingleNode("NO_CNPT") Is Nothing Then
                strNomeContraparte = "NULL"
            Else
                strNomeContraparte = "'" & .selectSingleNode("NO_CNPT").Text & "'"
            End If
            
            If .selectSingleNode("DE_TIPO_LIQU").Text = vbNullString Then
                strTipoLiqu = "NULL"
            Else
                strTipoLiqu = "'" & .selectSingleNode("DE_TIPO_LIQU").Text & "'"
            End If
            
            If .selectSingleNode("CO_LOCA_LIQU").Text <> enumLocalLiquidacao.CETIP Then
                If .selectSingleNode("DE_ATIV_MERC") Is Nothing Then
                    strAtivo = "NULL"
                Else
                    strAtivo = "'" & .selectSingleNode("DE_ATIV_MERC").Text & "'"
                End If
            
            Else
                If .selectSingleNode("CO_SUB_TIPO_ATIV") Is Nothing Then
                    strAtivo = "NULL"
                Else
                    strAtivo = "'" & .selectSingleNode("CO_SUB_TIPO_ATIV").Text & "'"
                End If
            
            End If
            
            strValues = lngCoMoviCaixSubRese & ",                                      " & vbNewLine & _
                        .selectSingleNode("CO_EMPR").Text & ",                         " & vbNewLine & _
                        IIf(.selectSingleNode("TP_OPER").Text = vbNullString, "          NULL", .selectSingleNode("TP_OPER").Text) & ", " & vbNewLine & _
                        fgDtXML_To_Oracle(strDataMovimento) & ",                       " & vbNewLine & _
                        IIf(.selectSingleNode("CO_LOCA_LIQU").Text = vbNullString, "     NULL", .selectSingleNode("CO_LOCA_LIQU").Text) & ", " & vbNewLine & _
                        " '" & .selectSingleNode("SG_SIST_ORIG").Text & "',            " & vbNewLine & _
                        " '" & .selectSingleNode("//CO_VEIC_LEGA").Text & "',          " & vbNewLine & _
                        " '" & strItemCaix & "',                                       " & vbNewLine & _
                        strTipoLiqu & ",                                               " & vbNewLine & _
                        fgVlrToDBServer(.selectSingleNode("//VA_OPER_ATIV").Text) & ", " & vbNewLine & _
                        .selectSingleNode("IN_OPER_DEBT_CRED").Text & ",               " & vbNewLine & _
                        strCNPJContraparte & ",                                        " & vbNewLine & _
                        strNomeContraparte & ",                                        " & vbNewLine & _
                        strAtivo & ",                                                  " & vbNewLine & _
                        IIf(.selectSingleNode("NU_SEQU_OPER_ATIV").Text = vbNullString, "NULL", .selectSingleNode("NU_SEQU_OPER_ATIV").Text) & ", " & vbNewLine & _
                        .selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text & ",      " & vbNewLine & _
                        .selectSingleNode("TP_GERA_MOVI").Text & ",                    " & vbNewLine & _
                        IIf(.selectSingleNode("TX_MOTI_GERA_MOVI").Text = vbNullString, "NULL", " '" & .selectSingleNode("TX_MOTI_GERA_MOVI").Text & "' ") & ", " & vbNewLine & _
                        " '" & Mid$(strItemCaix, 2, 3) & "', '" & Mid$(strItemCaix, 5, 3) & "', '" & Mid$(strItemCaix, 8, 3) & "', '" & Mid$(strItemCaix, 11, 3) & "', '" & Mid$(strItemCaix, 14, 3) & "', " & vbNewLine & _
                        " '" & fgUsuarioRede & "',                                     " & vbNewLine & _
                        " '" & fgEstacaoTrabalhoUsuario & "',                          " & vbNewLine & _
                        " SYSDATE)"
            
        Else
            strItemCaix = .selectSingleNode("CO_ITEM_CAIX").Text
            
            If Not flDiaUtil(.selectSingleNode("DH_MOVI_CAIX_SUB_RESE").Text) Then
                'Data do movimento não é um dia útil
                lngCodigoErroNegocio = 2027
                GoTo ErrorHandler
            End If
            
            strValues = lngCoMoviCaixSubRese & ",                                               " & vbNewLine & _
                        .selectSingleNode("CO_EMPR").Text & ",                                  " & vbNewLine & _
                        IIf(.selectSingleNode("TP_OPER").Text = vbNullString, "          NULL", .selectSingleNode("TP_OPER").Text) & ", " & vbNewLine & _
                        fgDtXML_To_Oracle(.selectSingleNode("DH_MOVI_CAIX_SUB_RESE").Text) & ", " & vbNewLine & _
                        IIf(.selectSingleNode("CO_LOCA_LIQU").Text = vbNullString, "     NULL", .selectSingleNode("CO_LOCA_LIQU").Text) & ", " & vbNewLine & _
                        " '" & .selectSingleNode("SG_SIST").Text & "',                          " & vbNewLine & _
                        " '" & .selectSingleNode("//CO_VEIC_LEGA").Text & "',                     " & vbNewLine & _
                        " '" & strItemCaix & "',                                                " & vbNewLine & _
                        IIf(.selectSingleNode("DE_TIPO_LIQU").Text = vbNullString, "     NULL", " '" & .selectSingleNode("DE_TIPO_LIQU").Text & "' ") & ", " & vbNewLine & _
                        fgVlrToDBServer(.selectSingleNode("VA_MOVI_CAIX_SUB_RESE").Text) & ",                    " & vbNewLine & _
                        .selectSingleNode("IN_MOVI_ENTR_SAID").Text & ",                        " & vbNewLine & _
                        IIf(.selectSingleNode("CO_CNPJ_CNPT").Text = vbNullString, "     NULL", .selectSingleNode("CO_CNPJ_CNPT").Text) & ", " & vbNewLine & _
                        IIf(.selectSingleNode("NO_CNPT").Text = vbNullString, "          NULL", " '" & .selectSingleNode("NO_CNPT").Text & "' ") & ", " & vbNewLine & _
                        " '" & .selectSingleNode("DE_ATIV").Text & "',                          " & vbNewLine & _
                        IIf(.selectSingleNode("NU_SEQU_OPER_ATIV").Text = vbNullString, "NULL", .selectSingleNode("NU_SEQU_OPER_ATIV").Text) & ", " & vbNewLine & _
                        .selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text & ",               " & vbNewLine & _
                        .selectSingleNode("TP_GERA_MOVI").Text & ",                             " & vbNewLine & _
                        IIf(.selectSingleNode("TX_MOTI_GERA_MOVI").Text = vbNullString, "NULL", " '" & .selectSingleNode("TX_MOTI_GERA_MOVI").Text & "' ") & ", " & vbNewLine & _
                        " '" & Mid$(strItemCaix, 2, 3) & "', '" & Mid$(strItemCaix, 5, 3) & "', '" & Mid$(strItemCaix, 8, 3) & "', '" & Mid$(strItemCaix, 11, 3) & "', '" & Mid$(strItemCaix, 14, 3) & "', " & vbNewLine & _
                        " '" & fgUsuarioRede & "',                                              " & vbNewLine & _
                        " '" & fgEstacaoTrabalhoUsuario & "',                                   " & vbNewLine & _
                        " SYSDATE)"
        
        End If
    End With
    
    flDefinirStrValues = strValues
    
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flDefinirStrValues Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Este método é responsável pela inclusão de um lançamento na tabela de movimento sub-reserva.

Private Function flIncluir(ByVal pxmlMoviSubRes As MSXML2.DOMDocument40) As Boolean
Dim strSQL                                  As String

    On Error GoTo ErrorHandler

'    fgGravaArquivo "LOGA6_clsMovimentoSubReserva_flIncluir_", vbNullString
    
    strSQL = " INSERT INTO A6.TB_MOVI_CAIX_SUB_RESE ( " & vbNewLine & _
             " CO_MOVI_CAIX_SUB_RESE,                 " & vbNewLine & _
             " CO_EMPR,                               " & vbNewLine & _
             " TP_OPER,                               " & vbNewLine & _
             " DH_MOVI_CAIX_SUB_RESE,                 " & vbNewLine & _
             " CO_LOCA_LIQU,                          " & vbNewLine & _
             " SG_SIST,                               " & vbNewLine & _
             " CO_VEIC_LEGA,                          " & vbNewLine & _
             " CO_ITEM_CAIX,                          " & vbNewLine & _
             " DE_TIPO_LIQU,                          " & vbNewLine & _
             " VA_MOVI_CAIX_SUB_RESE,                 " & vbNewLine & _
             " IN_MOVI_ENTR_SAID,                     " & vbNewLine & _
             " CO_CNPJ_CNPT,                          " & vbNewLine & _
             " NO_CNPT,                               " & vbNewLine & _
             " DE_ATIV,                               " & vbNewLine & _
             " NU_SEQU_OPER_ATIV,                     " & vbNewLine & _
             " CO_SITU_MOVI_CAIX_SUB_RESE,            " & vbNewLine & _
             " TP_GERA_MOVI,                          " & vbNewLine & _
             " TX_MOTI_GERA_MOVI,                     " & vbNewLine & _
             " CO_ITEM_CAIX_NIVE_01, CO_ITEM_CAIX_NIVE_02, CO_ITEM_CAIX_NIVE_03, CO_ITEM_CAIX_NIVE_04, CO_ITEM_CAIX_NIVE_05, " & vbNewLine & _
             " CO_USUA_ULTI_ATLZ,                     " & vbNewLine & _
             " CO_ETCA_TRAB_ULTI_ATLZ,                " & vbNewLine & _
             " DH_ULTI_ATLZ                           " & vbNewLine & _
             " ) VALUES (                             "
             
    strSQL = strSQL & flDefinirStrValues(pxmlMoviSubRes)
    
    Call fgExecuteSQL(strSQL)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela atualização da situação do movimento sub-reserva,
' e pelo retorno à camada intermediária.

Public Function AtualizarSituacaoMovimento(ByVal pxmlMoviSubRes As MSXML2.DOMDocument40) As Long
Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    With pxmlMoviSubRes.documentElement
        'Update deve ser feito com NU_SEQU_OPER_ATIV e IN_MOVI_ENTR_SAID
        'por causa das operações de Cliente 1
        
        strSQL = " UPDATE A6.TB_MOVI_CAIX_SUB_RESE " & vbNewLine & _
                 " SET    CO_SITU_MOVI_CAIX_SUB_RESE = " & .selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text & vbNewLine & _
                 " WHERE  NU_SEQU_OPER_ATIV          = " & .selectSingleNode("NU_SEQU_OPER_ATIV").Text & vbNewLine & _
                 " AND    IN_MOVI_ENTR_SAID          = " & .selectSingleNode("IN_OPER_DEBT_CRED").Text & vbNewLine & _
                 " AND    CO_SITU_MOVI_CAIX_SUB_RESE = " & enumTipoMovimento.RealizadoSolicitado & vbNewLine & _
                 " AND    CO_MOVI_CAIX_SUB_RESE = (SELECT MAX(CO_MOVI_CAIX_SUB_RESE)   " & vbNewLine & _
                 "                                 FROM   A6.TB_MOVI_CAIX_SUB_RESE     " & vbNewLine & _
                 "                                 WHERE  NU_SEQU_OPER_ATIV          = " & .selectSingleNode("NU_SEQU_OPER_ATIV").Text & vbNewLine & _
                 "                                 AND    IN_MOVI_ENTR_SAID          = " & .selectSingleNode("IN_OPER_DEBT_CRED").Text & vbNewLine & _
                 "                                 AND    CO_SITU_MOVI_CAIX_SUB_RESE = " & enumTipoMovimento.RealizadoSolicitado & ")"
    
    End With
    
    AtualizarSituacaoMovimento = fgExecuteSQL(strSQL)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela atualização geral da tabela de movimento sub-reserva,
' e pelo retorno à camada intermediária.

Public Function Salvar(ByVal pxmlMoviSubRes As MSXML2.DOMDocument40) As Boolean
    On Error GoTo ErrorHandler

    flIncluir pxmlMoviSubRes

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura do resumo de movimentação do caixa sub-reserva,
' e pelo retorno à camada intermediária.

Public Function ObterResumoCaixaSubReserva(ByVal pxmlFiltros As MSXML2.DOMDocument40) As String

Dim strSQL                                  As String

Dim lngGrupoVeiculoLegal                    As Long
Dim strVeiculoLegal                         As String
Dim strSiglaSistema                         As String
Dim lngEmpresa                              As Long

Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim intTipoBackOffice                       As Integer

    On Error GoTo ErrorHandler

    'Captura o Grupo de Veículo Legal proveniente do filtro, se houver
    If Not pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_GrupoVeiculoLegal/GrupoVeiculoLegal") Is Nothing Then
        lngGrupoVeiculoLegal = Val(pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_GrupoVeiculoLegal/GrupoVeiculoLegal").Text)
    End If
    
    'Captura o Veículo Legal proveniente do filtro, se houver
    If Not pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_VeiculoLegal/VeiculoLegal") Is Nothing Then
        strVeiculoLegal = pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_VeiculoLegal/VeiculoLegal").Text
    End If
    
    'Captura a Sigla de Sistema proveniente do filtro, se houver
    If Not pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_Sistema/Sistema") Is Nothing Then
        strSiglaSistema = pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_Sistema/Sistema").Text
    End If
    
    'Captura a Empresa proveniente do filtro, se houver
    If Not pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_BancoLiquidante/BancoLiquidante") Is Nothing Then
        lngEmpresa = Val(pxmlFiltros.selectSingleNode("Repeat_Filtros/Grupo_BancoLiquidante/BancoLiquidante").Text)
    End If
    
    'Atualiza a tabela << Posição do Caixa Disponível >>, somente se NÃO existir posição para a chave:
    'Véiculo Legal + Sigla Sistema
    Call flAtualizaPosicaoDisponivel(lngGrupoVeiculoLegal, strVeiculoLegal)
    
    strSQL = " Select    B.CO_GRUP_VEIC_LEGA, C.NO_GRUP_VEIC_LEGA, A.CO_VEIC_LEGA, TRIM(A.SG_SIST) As SG_SIST, B.NO_VEIC_LEGA, A.DT_CAIX_DISP, A.CO_SITU_CAIX, A.VA_UTLZ_ABER_CAIX, " & vbNewLine & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 1, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) - " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 1, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) As VA_PREV,  " & vbNewLine & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) As VA_SOLI, " & vbNewLine & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) As VA_CONF, " & vbNewLine & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) + " & vbNewLine & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) As VA_TOT_MOV, " & vbNewLine & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) + " & vbNewLine & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) -  " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) - " & vbNewLine & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 1, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) - " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 1, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) As VA_VAR_MOV, " & vbNewLine

    strSQL = strSQL & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) - " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) As VA_TOT_SOLI, " & vbNewLine & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) - " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) As VA_TOT_CONF, " & vbNewLine & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) - " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) + " & vbNewLine & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) - " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) As VA_TOT_FECH, " & vbNewLine
            
    strSQL = strSQL & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) - " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 2, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) As VA_VAR_FECH_1, " & vbNewLine
             
    strSQL = strSQL & _
             "          (SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(D.IN_MOVI_ENTR_SAID, 1, D.VA_MOVI_CAIX_SUB_RESE, 0),0)) - " & vbNewLine & _
             "           SUM(DECODE(D.CO_SITU_MOVI_CAIX_SUB_RESE, 3, DECODE(D.IN_MOVI_ENTR_SAID, 2, D.VA_MOVI_CAIX_SUB_RESE, 0),0))) As VA_VAR_FECH_2 " & vbNewLine

    strSQL = strSQL & _
             " From      A6.TB_POSI_CAIX_DISP           A,                          " & vbNewLine & _
             "           A8.TB_VEIC_LEGA                B,                          " & vbNewLine & _
             "           A8.TB_GRUP_VEIC_LEGA           C,                          " & vbNewLine & _
             "           A6.TB_MOVI_CAIX_SUB_RESE       D                           " & vbNewLine & _
             " Where     A.SG_SIST                      = B.SG_SIST                 " & vbNewLine & _
             " And       A.CO_VEIC_LEGA                 = B.CO_VEIC_LEGA            " & vbNewLine & _
             " And       B.CO_GRUP_VEIC_LEGA            = C.CO_GRUP_VEIC_LEGA       " & vbNewLine & _
             " And       A.SG_SIST                      = D.SG_SIST                 " & vbNewLine & _
             " And       A.CO_VEIC_LEGA                 = D.CO_VEIC_LEGA            " & vbNewLine & _
             " And       A.DT_CAIX_DISP                 = D.DH_MOVI_CAIX_SUB_RESE   " & vbNewLine
    
    If lngGrupoVeiculoLegal <> 0 Then
        strSQL = strSQL & _
                 " And       C.CO_GRUP_VEIC_LEGA = " & lngGrupoVeiculoLegal & vbNewLine
    End If
    
    If strVeiculoLegal <> vbNullString Then
        strSQL = strSQL & _
                 " And       B.CO_VEIC_LEGA = '" & strVeiculoLegal & "'" & vbNewLine
    End If
    
    If strSiglaSistema <> vbNullString Then
        strSQL = strSQL & _
                 " And       B.SG_SIST = '" & strSiglaSistema & "'" & vbNewLine
    End If
    
    If lngEmpresa <> 0 Then
        strSQL = strSQL & _
                 " And       D.CO_EMPR = " & lngEmpresa & vbNewLine
    End If

    intTipoBackOffice = 99
    For Each xmlDomNode In pxmlFiltros.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
        intTipoBackOffice = Val(xmlDomNode.Text)
    Next
    
    Select Case intTipoBackOffice
        Case 99
            strSQL = strSQL & _
                     fgSegregaDados("A8.TB_VEIC_LEGA", _
                                    False, "B", "B", False, False, True, False, True, False)
        Case Is > 0
            strSQL = strSQL & _
                     vbNewLine & " AND B.TP_BKOF = " & intTipoBackOffice & vbNewLine
    End Select

    strSQL = strSQL & _
            " GROUP BY  B.CO_GRUP_VEIC_LEGA, C.NO_GRUP_VEIC_LEGA, A.CO_VEIC_LEGA, TRIM(A.SG_SIST), B.NO_VEIC_LEGA, A.DT_CAIX_DISP, A.CO_SITU_CAIX, A.VA_UTLZ_ABER_CAIX " & vbNewLine
            
    strSQL = strSQL & _
            " ORDER BY B.CO_GRUP_VEIC_LEGA, C.NO_GRUP_VEIC_LEGA, A.CO_VEIC_LEGA, TRIM(A.SG_SIST), B.NO_VEIC_LEGA"
    
    ObterResumoCaixaSubReserva = fgQueryXMLLerTodos("MovimentoGrupoSubReserva", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterResumoCaixaSubReserva Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Este método é responsável pela atualização da posição disponível do caixa sub-reserva.

Private Function flAtualizaPosicaoDisponivel(Optional ByVal plngGrupoVeiculoLegal As Long = 0, _
                                             Optional ByVal pstrVeiculoLegal As String = vbNullString, _
                                             Optional ByVal pstrSiglaSistema As String = vbNullString) As Boolean

Dim vntArray()                              As Variant
    
    On Error GoTo ErrorHandler

    'Atualiza os registros na tabela de apoio << A6.TB_POSI_CAIX_DISP >>,
    'se ainda não estiverem OK
    vntArray = Array(plngGrupoVeiculoLegal, pstrVeiculoLegal, pstrSiglaSistema)
    Call fgExecuteCMD("A6PROC.A6P_ATUALIZA_POSI_DISP", 0, vntArray)
    
    flAtualizaPosicaoDisponivel = True
    
    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAtualizaPosicaoDisponivel Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método verifica se uma determinada data, é dia útil ou não.

Private Function flDiaUtil(ByVal pstrData As String) As Boolean

Dim strSQL                                  As String
Dim rsData                                  As ADODB.Recordset
Dim datData                                 As Date

On Error GoTo ErrorHandler

    datData = fgDtXML_To_Date(pstrData)
    
    If Weekday(datData) = vbSunday Or Weekday(datData) = vbSaturday Then
        flDiaUtil = False
        Exit Function
    End If

    strSQL = " SELECT DT_FERI FROM A8.TB_FERIADO_HO WHERE DT_FERI = " & fgDtXML_To_Oracle(pstrData)

    Set rsData = fgQuerySQL(strSQL)

    flDiaUtil = rsData.RecordCount < 1

    Set rsData = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set rsData = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flDiaUtil Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
