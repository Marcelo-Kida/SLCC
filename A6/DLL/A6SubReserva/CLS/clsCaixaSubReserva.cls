VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaixaSubReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA\r\nPacote         :\r\nClasse         : clsCaixaSubReserva\r\nData Criação   : 14/07/2003\r\nObjetivo       :\r\n\r\nAnalista       : Carlos Fortes\r\n\r\nProgramador    : Carlos Fortes\r\nData           : 14/07/2003\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F105A0302A7"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela consulta e atualização
' do Movimento de Caixa Sub-reserva.

Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Implements COMSVCSLib.ObjectControl
Private objContext                          As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

' Este método é responsável pela leitura dos valores para abertura do caixa,
' e pelo retorno à camada intermediária.

Public Function ObterValoresAbertura(ByVal pxmlFiltros As MSXML2.DOMDocument40) As String

Dim xmlDomNode                              As MSXML2.IXMLDOMNode

Dim blnPrimeiro                             As Boolean
Dim strSQL                                  As String
Dim strWhere                                As String
Dim strSegregacao                           As String

Dim datDataDispDefault                      As Date
Dim datDataSaldo                            As Date
    
    On Error GoTo ErrorHandler

    datDataDispDefault = fgAdicionarDiasUteis(fgDataHoraServidor(DataAux) - 1, 1, Proximo)
    datDataSaldo = fgAdicionarDiasUteis(datDataDispDefault, 1, Anterior)
    
    strSQL = " SELECT B.CO_VEIC_LEGA, B.SG_SIST, B.NO_VEIC_LEGA, B.CO_GRUP_VEIC_LEGA,                           " & vbNewLine & _
             "        B.TP_BKOF , B.DT_CAIX_SUB_RESE, B.CO_SITU_CAIX_SUB_RESE, C.VA_SALD_FECH,                  " & vbNewLine & _
             "        C.DT_APUR_SALD, B.VA_FECH_CAIX_SUB_RESE, B.DT_ABER_CAIX                                   " & vbNewLine & _
             " FROM ((SELECT A.CO_VEIC_LEGA, A.SG_SIST, D.NO_VEIC_LEGA, D.CO_GRUP_VEIC_LEGA,                    " & vbNewLine & _
             "               D.TP_BKOF , A.DT_CAIX_SUB_RESE, A.CO_SITU_CAIX_SUB_RESE, A.VA_FECH_CAIX_SUB_RESE,  " & vbNewLine & _
             "               A8PROC.A8F_PROXIMO_DIA_UTIL(A.DT_CAIX_SUB_RESE) AS DT_ABER_CAIX                    " & vbNewLine & _
             "        FROM   A6.TB_POSI_CAIX_SUB_RESE A,                                                        " & vbNewLine & _
             "               A8.TB_VEIC_LEGA D                                                                  " & vbNewLine & _
             "        Where  D.CO_VEIC_LEGA   = A.CO_VEIC_LEGA                                                  " & vbNewLine & _
             "        AND    D.SG_SIST        = A.SG_SIST                                                       " & vbNewLine & _
             "        AND    DT_CAIX_SUB_RESE = (SELECT MAX(DT_CAIX_SUB_RESE)                                   " & vbNewLine & _
             "                                   From A6.TB_POSI_CAIX_SUB_RESE                                  " & vbNewLine & _
             "                                   Where CO_VEIC_LEGA = A.CO_VEIC_LEGA                            " & vbNewLine & _
             "                                   AND    SG_SIST      = A.SG_SIST)                               " & vbNewLine & _
             "        AND    A.CO_SITU_CAIX_SUB_RESE = " & enumEstadoCaixa.Fechado & ")                         "
    strSQL = strSQL & "" & _
             " Union All                                                                                        " & vbNewLine & _
             "       (SELECT v.CO_VEIC_LEGA, v.SG_SIST, v.NO_VEIC_LEGA, v.CO_GRUP_VEIC_LEGA, v.TP_BKOF,         " & vbNewLine & _
             "               TO_DATE('" & fgDt_To_Xml(datDataSaldo) & "', 'YYYYMMDD') AS DT_CAIX_SUB_RESE,      " & vbNewLine & _
             "               1 AS CO_SITU_CAIX_SUB_RESE, 0 AS VA_FECH_CAIX_SUB_RESE,                            " & vbNewLine & _
             "               TO_DATE('" & fgDt_To_Xml(datDataDispDefault) & "', 'YYYYMMDD') AS DT_ABER_CAIX     " & vbNewLine & _
             "        From A8.TB_VEIC_LEGA v                                                                    " & vbNewLine & _
             "        WHERE  NOT EXISTS  (SELECT     NULL                                                       " & vbNewLine & _
             "                            FROM       a6.tb_posi_caix_sub_rese p                                 " & vbNewLine & _
             "                            where      p.co_veic_lega = v.co_veic_lega                            " & vbNewLine & _
             "                            and        p.sg_sist      = v.sg_sist))) b,                           " & vbNewLine & _
             " A6.TB_SALD_FECH_VEIC_LEGA C                                                                      " & vbNewLine
             
    strSQL = strSQL & _
             " WHERE B.CO_VEIC_LEGA                           = C.CO_VEIC_LEGA (+)                              " & vbNewLine & _
             " AND   B.SG_SIST                                = C.SG_SIST      (+)                              " & vbNewLine & _
             " AND   to_date(B.DT_CAIX_SUB_RESE,'dd/mm/yyyy') = to_date(C.DT_APUR_SALD(+),'dd/mm/yyyy')         " & vbNewLine

    strWhere = vbNullString
    blnPrimeiro = True
    For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_VeiculoLegal/*")
        If blnPrimeiro Then
            strWhere = strWhere & " AND B.CO_VEIC_LEGA IN ('"
            strWhere = strWhere & Trim$(xmlDomNode.Text)
            blnPrimeiro = False
        Else
            strWhere = strWhere & "', '" & Trim$(xmlDomNode.Text)
        End If
    Next xmlDomNode
    If Not blnPrimeiro Then
        strWhere = strWhere & "')" & vbNewLine
    End If

    blnPrimeiro = True
    For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_GrupoVeiculoLegal/*")
        If blnPrimeiro Then
            strWhere = strWhere & " AND B.CO_GRUP_VEIC_LEGA IN ("
            strWhere = strWhere & xmlDomNode.Text
            blnPrimeiro = False
        Else
            strWhere = strWhere & ", " & xmlDomNode.Text
        End If
    Next xmlDomNode
    If Not blnPrimeiro Then
        strWhere = strWhere & ")" & vbNewLine
    End If

    strSegregacao = fgSegregaDados("A8.TB_VEIC_LEGA", _
                                   False, _
                                   "B", _
                                   "B", _
                                   False, _
                                   False, _
                                   True, _
                                   False, _
                                   True, _
                                   False)

    strSQL = strSQL & strWhere & strSegregacao & _
             " ORDER BY B.NO_VEIC_LEGA, B.CO_VEIC_LEGA"

    ObterValoresAbertura = fgQueryXMLLerTodos("CaixaSubReserva", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set xmlDomNode = Nothing

    Exit Function

ErrorHandler:
    Set xmlDomNode = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterValoresAbertura Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela abertura do caixa sub-reserva,
' e pelo retorno à camada intermediária.

Public Function AbrirCaixa(ByVal pxmlDomNode As MSXML2.IXMLDOMNode) As Boolean
Dim strSQL                                  As String
Dim vntValorAberturaCaixa                   As Variant
Dim xmlPosicaoCaixa                         As MSXML2.DOMDocument40
Dim datDataAberturaCaixa                    As Date
Dim intSituacaoCaixa                        As Integer

On Error GoTo ErrorHandler

    'Define o valor de Abertura do Caixa
    Select Case CLng(pxmlDomNode.selectSingleNode("TipoAbertura").Text)
        Case enumTipoAbertura.SaldoConta
            vntValorAberturaCaixa = CVar(pxmlDomNode.selectSingleNode("VA_MOVI_CAIX_SUB_RESE").Text)
        Case enumTipoAbertura.ValorD1
            vntValorAberturaCaixa = CVar(pxmlDomNode.selectSingleNode("VA_ABER_RECE_CAIX_SUB_RESE").Text)
        Case enumTipoAbertura.ValorInformado
            vntValorAberturaCaixa = CVar(pxmlDomNode.selectSingleNode("VA_ABER_INFO_CAIX_SUB_RESE").Text)
    End Select
        
    Set xmlPosicaoCaixa = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlPosicaoCaixa.loadXML(ObterPosicaoCaixaSubReserva(pxmlDomNode.selectSingleNode("CO_VEIC_LEGA").Text, _
                                                             pxmlDomNode.selectSingleNode("SG_SIST").Text))
                                                             
    datDataAberturaCaixa = fgDtXML_To_Date(xmlPosicaoCaixa.documentElement.selectSingleNode("//DT_CAIX_SUB_RESE_ATUAL").Text)
    intSituacaoCaixa = xmlPosicaoCaixa.documentElement.selectSingleNode("//CO_SITU_CAIX_SUB_RESE_ATUAL").Text
    
    If intSituacaoCaixa = enumEstadoCaixa.Fechado Then
        datDataAberturaCaixa = fgAdicionarDiasUteis(datDataAberturaCaixa, 1, Proximo)
    End If
    
    strSQL = "INSERT INTO " & vbCrLf & _
            "   A6.TB_POSI_CAIX_SUB_RESE " & vbCrLf & _
            "  (SG_SIST, " & vbCrLf & _
            "   DT_CAIX_SUB_RESE, " & vbCrLf & _
            "   CO_VEIC_LEGA, " & vbCrLf & _
            "   CO_SITU_CAIX_SUB_RESE, " & vbCrLf & _
            "   VA_UTLZ_ABER_CAIX_SUB_RESE, " & vbCrLf & _
            "   VA_ABER_RECE_CAIX_SUB_RESE, " & vbCrLf & _
            "   VA_MOVI_CAIX_SUB_RESE, " & vbCrLf & _
            "   VA_ABER_INFO_CAIX_SUB_RESE, " & vbCrLf & _
            "   VA_FECH_CAIX_SUB_RESE, " & vbCrLf & _
            "   CO_USUA_ULTI_ATLZ, " & vbCrLf & _
            "   CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
            "   DH_ULTI_ATLZ ) " & vbCrLf

    strSQL = strSQL & _
            "VALUES ( " & vbCrLf & _
            "'" & pxmlDomNode.selectSingleNode("SG_SIST").Text & "', " & vbCrLf & _
            fgDtXML_To_Oracle(fgDt_To_Xml(datDataAberturaCaixa)) & ", " & vbCrLf & _
            "'" & pxmlDomNode.selectSingleNode("CO_VEIC_LEGA").Text & "', " & vbCrLf & _
            enumEstadoCaixa.Aberto & ", " & vbCrLf & _
            fgVlrToDBServer(vntValorAberturaCaixa) & ", " & vbCrLf & _
            fgVlrToDBServer(pxmlDomNode.selectSingleNode("VA_ABER_RECE_CAIX_SUB_RESE").Text) & ", " & vbCrLf & _
            fgVlrToDBServer(pxmlDomNode.selectSingleNode("VA_MOVI_CAIX_SUB_RESE").Text) & ", " & vbCrLf & _
            fgVlrToDBServer(pxmlDomNode.selectSingleNode("VA_ABER_INFO_CAIX_SUB_RESE").Text) & ", " & vbCrLf & _
            "0, " & vbCrLf & _
            "'" & fgUsuarioRede & "' ," & vbCrLf & _
            "'" & fgEstacaoTrabalhoUsuario & "' ," & vbCrLf & _
            "SYSDATE)"

    fgExecuteSQL (strSQL)
    
    Call flAtualizaPosicaoDisponivel(pxmlDomNode.selectSingleNode("CO_VEIC_LEGA").Text, _
                                     pxmlDomNode.selectSingleNode("SG_SIST").Text, _
                                     datDataAberturaCaixa, _
                                     enumEstadoCaixa.Aberto, _
                                     vntValorAberturaCaixa)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AbrirCaixa Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura dos valores para o fechamento do caixa,
' e pelo retorno à camada intermediária.

Public Function ObterValoresFechamento(ByVal pxmlFiltros As MSXML2.DOMDocument40)

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim blnPrimeiro                              As Boolean
Dim strSQL                                   As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT A.CO_VEIC_LEGA, A.SG_SIST, C.NO_VEIC_LEGA, A.DT_CAIX_DISP, A.VA_UTLZ_ABER_CAIX, A.CO_SITU_CAIX," & vbNewLine & _
             "       (NVL(SUM(DECODE(B.IN_MOVI_ENTR_SAID, 1, B.VA_MOVI_CAIX_SUB_RESE, 0)),0) - " & vbNewLine & _
             "        NVL(SUM(DECODE(B.IN_MOVI_ENTR_SAID, 2, B.VA_MOVI_CAIX_SUB_RESE, 0)),0)) AS VA_MOVI_CAIX_SUB_RESE" & vbNewLine & _
             " FROM   A6.TB_POSI_CAIX_DISP     A," & vbNewLine & _
             "        A6.TB_MOVI_CAIX_SUB_RESE B," & vbNewLine & _
             "        A8.TB_VEIC_LEGA          C" & vbNewLine & _
             " WHERE  A.CO_VEIC_LEGA        =  B.CO_VEIC_LEGA           (+)" & vbNewLine & _
             " AND    A.SG_SIST             =  B.SG_SIST                (+)" & vbNewLine & _
             " AND    TO_CHAR(A.DT_CAIX_DISP, 'DD/MM/YYYY') =  TO_CHAR(B.DH_MOVI_CAIX_SUB_RESE (+), 'DD/MM/YYYY')" & vbNewLine & _
             " AND    A.CO_VEIC_LEGA = C.CO_VEIC_LEGA" & vbNewLine & _
             " AND    A.SG_SIST      = C.SG_SIST" & vbNewLine & _
             " AND    A.CO_SITU_CAIX = " & enumEstadoCaixa.Aberto & vbNewLine

    blnPrimeiro = True
    For Each objDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_VeiculoLegal/*")
        If blnPrimeiro Then
            strSQL = strSQL & "AND   A.CO_VEIC_LEGA IN ('" & Trim(objDomNode.Text)
            blnPrimeiro = False
        Else
            strSQL = strSQL & "', " & Trim(objDomNode.Text)
        End If
    Next objDomNode
    If blnPrimeiro = False Then
        strSQL = strSQL & "')" & vbNewLine
    End If

    blnPrimeiro = True
    If blnPrimeiro Then
        For Each objDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_GrupoVeiculoLegal/*")
            If blnPrimeiro Then
                strSQL = strSQL & " AND   A.CO_VEIC_LEGA IN (SELECT C.CO_VEIC_LEGA " & vbNewLine & _
                                  "                          FROM A8.TB_VEIC_LEGA C " & vbNewLine & _
                                  "                          WHERE  C.CO_GRUP_VEIC_LEGA IN ("
                strSQL = strSQL & CLng(objDomNode.Text)
                blnPrimeiro = False
            Else
                strSQL = strSQL & ", " & CLng(objDomNode.Text)
            End If
        Next objDomNode
        If blnPrimeiro = False Then
            strSQL = strSQL & "))" & vbNewLine
        End If
    End If

    strSQL = strSQL & _
             fgSegregaDados("A8.TB_VEIC_LEGA", _
                            False, _
                            "C", _
                            "C", _
                            False, _
                            False, _
                            True, _
                            False, _
                            True, _
                            False)
    
    strSQL = strSQL & " GROUP BY A.CO_VEIC_LEGA, A.SG_SIST, C.NO_VEIC_LEGA, A.DT_CAIX_DISP, A.VA_UTLZ_ABER_CAIX, A.CO_SITU_CAIX"
    strSQL = strSQL & " ORDER BY NO_VEIC_LEGA"

    ObterValoresFechamento = fgQueryXMLLerTodos("CaixaSubReserva", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterValoresFechamento Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pelo fechamento do caixa sub-reserva,
' e pelo retorno à camada intermediária.

Public Function FecharCaixa(ByVal pstrCodigoVeiculoLegal As String, _
                            ByVal pstrSiglaSistema As String) As Boolean

Dim strSQL                                  As String
Dim rsRetorno                               As ADODB.Recordset
Dim datDisponivel                           As Date

    On Error GoTo ErrorHandler

    strSQL = _
            "SELECT DT_CAIX_SUB_RESE, " & vbNewLine & _
            "       CO_VEIC_LEGA, " & vbNewLine & _
            "       SG_SIST, " & vbNewLine & _
            "       " & enumEstadoCaixa.Fechado & " AS CO_SITU_CAIX_SUB_RESE, " & vbNewLine & _
            "       VA_UTLZ_ABER_CAIX_SUB_RESE, " & vbNewLine & _
            "       VA_ABER_RECE_CAIX_SUB_RESE, " & vbNewLine & _
            "       (SELECT NVL(SUM(DECODE(B.IN_MOVI_ENTR_SAID, 1, B.VA_MOVI_CAIX_SUB_RESE, 0)),0) - " & vbNewLine & _
            "               NVL(SUM(DECODE(B.IN_MOVI_ENTR_SAID, 2, B.VA_MOVI_CAIX_SUB_RESE, 0)),0) " & vbNewLine & _
            "        FROM A6.TB_MOVI_CAIX_SUB_RESE B " & vbNewLine & _
            "        WHERE A.CO_VEIC_LEGA = B.CO_VEIC_LEGA " & vbNewLine & _
            "        AND   B.DH_MOVI_CAIX_SUB_RESE = TRUNC(A.DT_CAIX_SUB_RESE)) AS VA_MOVI_CAIX_SUB_RESE, " & vbNewLine & _
            "       VA_ABER_INFO_CAIX_SUB_RESE, " & vbNewLine & _
            "       (SELECT NVL(SUM(DECODE(B.IN_MOVI_ENTR_SAID, 1, B.VA_MOVI_CAIX_SUB_RESE, 0)),0) - " & vbNewLine & _
            "               NVL(SUM(DECODE(B.IN_MOVI_ENTR_SAID, 2, B.VA_MOVI_CAIX_SUB_RESE, 0)),0) " & vbNewLine & _
            "        FROM A6.TB_MOVI_CAIX_SUB_RESE B " & vbNewLine & _
            "        WHERE A.CO_VEIC_LEGA = B.CO_VEIC_LEGA " & vbNewLine & _
            "        AND   B.DH_MOVI_CAIX_SUB_RESE = TRUNC(A.DT_CAIX_SUB_RESE)) + VA_UTLZ_ABER_CAIX_SUB_RESE AS VA_FECH_CAIX_SUB_RESE, " & vbNewLine & _
            "       '" & fgUsuarioRede & "' AS CO_USUA_ULTI_ATLZ, " & vbNewLine & _
            "       '" & fgEstacaoTrabalhoUsuario & "'  AS CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
            "       SYSDATE AS DH_ULTI_ATLZ " & vbNewLine

    strSQL = strSQL & _
            "FROM   A6.TB_POSI_CAIX_SUB_RESE A " & vbNewLine & _
            "WHERE  A.CO_SITU_CAIX_SUB_RESE = " & enumEstadoCaixa.Aberto & vbNewLine & _
            "AND    A.CO_VEIC_LEGA = '" & pstrCodigoVeiculoLegal & "'" & vbNewLine & _
            "AND    A.SG_SIST      = '" & pstrSiglaSistema & "'" & vbNewLine & _
            "AND    A.DT_CAIX_SUB_RESE = (SELECT MAX(DT_CAIX_SUB_RESE) FROM A6.TB_POSI_CAIX_SUB_RESE B" & vbNewLine & _
            "                             WHERE  B.CO_VEIC_LEGA     = A.CO_VEIC_LEGA" & vbNewLine & _
            "                             AND    B.SG_SIST          = A.SG_SIST)"
            
    Set rsRetorno = fgQuerySQL(strSQL)
        
    If rsRetorno.EOF Then
        strSQL = Replace$(strSQL, "A6.", "A6HIST.")
        Set rsRetorno = fgQuerySQL(strSQL)
    End If
    
    strSQL = "INSERT INTO A6.TB_POSI_CAIX_SUB_RESE (" & vbNewLine & _
             "       DT_CAIX_SUB_RESE, " & vbNewLine & _
             "       CO_VEIC_LEGA, " & vbNewLine & _
             "       SG_SIST, " & vbNewLine & _
             "       CO_SITU_CAIX_SUB_RESE, " & vbNewLine & _
             "       VA_UTLZ_ABER_CAIX_SUB_RESE, " & vbNewLine & _
             "       VA_ABER_RECE_CAIX_SUB_RESE, " & vbNewLine & _
             "       VA_ABER_INFO_CAIX_SUB_RESE, " & vbNewLine & _
             "       VA_MOVI_CAIX_SUB_RESE, " & vbNewLine & _
             "       VA_FECH_CAIX_SUB_RESE, " & vbNewLine & _
             "       CO_USUA_ULTI_ATLZ, " & vbNewLine & _
             "       CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
             "       DH_ULTI_ATLZ) " & vbNewLine

    strSQL = strSQL & _
            "VALUES (" & fgDtXML_To_Oracle(fgDt_To_Xml(rsRetorno("DT_CAIX_SUB_RESE"))) & ", " & vbNewLine & _
            "        '" & rsRetorno("CO_VEIC_LEGA") & "',                                   " & vbNewLine & _
            "        '" & rsRetorno("SG_SIST") & "',                                        " & vbNewLine & _
            "        " & rsRetorno("CO_SITU_CAIX_SUB_RESE") & ",                            " & vbNewLine & _
            "        " & fgVlrToDBServer(rsRetorno("VA_UTLZ_ABER_CAIX_SUB_RESE")) & ",      " & vbNewLine & _
            "        " & fgVlrToDBServer(rsRetorno("VA_ABER_RECE_CAIX_SUB_RESE")) & ",      " & vbNewLine & _
            "        " & fgVlrToDBServer(rsRetorno("VA_ABER_INFO_CAIX_SUB_RESE")) & ",      " & vbNewLine & _
            "        " & fgVlrToDBServer(rsRetorno("VA_MOVI_CAIX_SUB_RESE")) & ",           " & vbNewLine & _
            "        " & fgVlrToDBServer(rsRetorno("VA_FECH_CAIX_SUB_RESE")) & ",           " & vbNewLine & _
            "        '" & rsRetorno("CO_USUA_ULTI_ATLZ") & "',                              " & vbNewLine & _
            "        '" & rsRetorno("CO_ETCA_TRAB_ULTI_ATLZ") & "',                         " & vbNewLine & _
            "        " & fgDtXML_To_Oracle(fgDt_To_Xml(rsRetorno("DT_CAIX_SUB_RESE"))) & ") "

    fgExecuteSQL (strSQL)
    
    datDisponivel = fgAdicionarDiasUteis(rsRetorno("DT_CAIX_SUB_RESE"), 1, Proximo)
    Call flAtualizaPosicaoDisponivel(pstrCodigoVeiculoLegal, _
                                     pstrSiglaSistema, _
                                     datDisponivel, _
                                     enumEstadoCaixa.Disponivel, _
                                     rsRetorno("VA_UTLZ_ABER_CAIX_SUB_RESE"))
    
    Set rsRetorno = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set rsRetorno = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "FecharCaixa Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura da posição atual do caixa sub-reserva,
' e pelo retorno à camada intermediária.

Public Function ObterPosicaoCaixaSubReserva(ByVal pstrCodigoVeiculoLegal As String, _
                                            ByVal pstrSiglaSistema As String, _
                                   Optional ByVal pstrDataCaixaSubReserva As String = "", _
                                   Optional ByVal pblnRetornaPosicaoDisponivel As Boolean = True) As String

On Error GoTo ErrorHandler

Dim xmlSubReserva                           As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim strSQL                                  As String
Dim strDataCaixaSubReserva                  As String
Dim strPosicaoCaixaSubReserva               As String

On Error GoTo ErrorHandler

    strSQL = " Select    DT_CAIX_SUB_RESE, " & vbNewLine & _
             "           CO_VEIC_LEGA, " & vbNewLine & _
             "           TRIM(SG_SIST)  As SG_SIST, " & vbNewLine & _
             "           CO_SITU_CAIX_SUB_RESE, " & vbNewLine & _
             "           VA_UTLZ_ABER_CAIX_SUB_RESE, " & vbNewLine & _
             "           VA_ABER_RECE_CAIX_SUB_RESE, " & vbNewLine & _
             "           VA_ABER_INFO_CAIX_SUB_RESE, " & vbNewLine & _
             "           VA_FECH_CAIX_SUB_RESE " & vbNewLine & _
             " From      A6.TB_POSI_CAIX_SUB_RESE " & vbNewLine & _
             " Where    (CO_SITU_CAIX_SUB_RESE, " & vbNewLine & _
             "           DT_CAIX_SUB_RESE) IN ( " & vbNewLine & _
             "               Select      MAX(CO_SITU_CAIX_SUB_RESE), " & vbNewLine & _
             "                           MAX(DT_CAIX_SUB_RESE) " & vbNewLine & _
             "               From        A6.TB_POSI_CAIX_SUB_RESE " & vbNewLine & _
             "               Where       DT_CAIX_SUB_RESE   IN ( " & vbNewLine & _
             "                               Select      MAX(DT_CAIX_SUB_RESE)       As DT_CAIX_SUB_RESE " & vbNewLine & _
             "                               From        A6.TB_POSI_CAIX_SUB_RESE " & vbNewLine & _
             "                               Where       CO_VEIC_LEGA   = '" & pstrCodigoVeiculoLegal & "'" & vbNewLine
             
    If pstrDataCaixaSubReserva <> vbNullString Then
        strSQL = strSQL & _
                 "                               And         DT_CAIX_SUB_RESE BETWEEN TO_DATE('" & pstrDataCaixaSubReserva & "000000', 'YYYYMMDDHH24MISS')" & vbNewLine & _
                 "                                                            And     TO_DATE('" & pstrDataCaixaSubReserva & "235959', 'YYYYMMDDHH24MISS')" & vbNewLine
    End If
    
    strSQL = strSQL & _
             "                               And         SG_SIST        = '" & pstrSiglaSistema & "')" & vbNewLine & _
             "               And         CO_VEIC_LEGA        = '" & pstrCodigoVeiculoLegal & "'" & vbNewLine & _
             "               And         SG_SIST             = '" & pstrSiglaSistema & "')" & vbNewLine & _
             " And       CO_VEIC_LEGA        = '" & pstrCodigoVeiculoLegal & "'" & vbNewLine & _
             " And       SG_SIST             = '" & pstrSiglaSistema & "'"

    
    'Verificação da Base histórica
    If pstrDataCaixaSubReserva <> vbNullString Then
        If fgDtXML_To_Date(pstrDataCaixaSubReserva) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
            strSQL = Replace(strSQL, "A6.TB_POSI_CAIX_SUB_RESE", "A6HIST.TB_POSI_CAIX_SUB_RESE")
        End If
    End If
    
    strPosicaoCaixaSubReserva = fgQueryXMLLerTodos("PosicaoCaixaSubReserva", strSQL, App.EXEName & "." & TypeName(Me))
    
    Set xmlSubReserva = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Trim(strPosicaoCaixaSubReserva) <> "" Then
        'Posição de Caixa encontrada
        '-------------------------------------------------------------------------------------------------------
        Call xmlSubReserva.loadXML(strPosicaoCaixaSubReserva)
        
        For Each xmlDomNode In xmlSubReserva.documentElement.childNodes
            'Existe Caixa Reserva
            If Val(xmlDomNode.selectSingleNode("CO_SITU_CAIX_SUB_RESE").Text) = enumEstadoCaixa.Aberto Then
                'Caixa Reserva Aberto Data do Caixa = Data do Caixa SubReserva
                
                'Cria Nodes de data da posição e situação atual
                Call fgAppendNode(xmlSubReserva, _
                                  "Grupo_PosicaoCaixaSubReserva", _
                                  "DT_CAIX_SUB_RESE_ATUAL", _
                                  xmlDomNode.selectSingleNode("DT_CAIX_SUB_RESE").Text)
            
                Call fgAppendNode(xmlSubReserva, _
                                  "Grupo_PosicaoCaixaSubReserva", _
                                  "CO_SITU_CAIX_SUB_RESE_ATUAL", _
                                  enumEstadoCaixa.Aberto)
                
            ElseIf Val(xmlDomNode.selectSingleNode("CO_SITU_CAIX_SUB_RESE").Text) = enumEstadoCaixa.Fechado Then
                'Caixa Reserva Fechado Data do Caixa = Próximo Dia Útil
                'a partir da Data do Caixa SubReserva
                strDataCaixaSubReserva = Format(fgAdicionarDiasUteis(fgDtXML_To_Date(xmlDomNode.selectSingleNode("DT_CAIX_SUB_RESE").Text), _
                                                                    1, _
                                                                    enumPaginacao.Proximo), _
                                                "yyyymmdd")
                
                'Cria Nodes de data da posição e situação atual
                Call fgAppendNode(xmlSubReserva, _
                                  "Grupo_PosicaoCaixaSubReserva", _
                                  "DT_CAIX_SUB_RESE_ATUAL", _
                                  xmlDomNode.selectSingleNode("DT_CAIX_SUB_RESE").Text)
            
                Call fgAppendNode(xmlSubReserva, _
                                  "Grupo_PosicaoCaixaSubReserva", _
                                  "CO_SITU_CAIX_SUB_RESE_ATUAL", _
                                  enumEstadoCaixa.Fechado)
                                  
                'Atribui o próximo dia útil ao Node de Data do Caixa SubReserva e
                'refaz o XML de retorno
                '   Remove o conteúdo do node original (DT_CAIX_SUB_RESE) e adiciona o novo conteúdo
                Call fgRemoveNode(xmlSubReserva, "DT_CAIX_SUB_RESE")
                Call fgAppendNode(xmlSubReserva, _
                                  "Grupo_PosicaoCaixaSubReserva", _
                                  "DT_CAIX_SUB_RESE", _
                                  strDataCaixaSubReserva)
            
                'Atribui a situação disponível ao Node de Situação do Caixa e
                'refaz o XML de retorno
                '   Remove o conteúdo do node original (DT_CAIX_SUB_RESE) e adiciona o novo conteúdo
                Call fgRemoveNode(xmlSubReserva, "CO_SITU_CAIX_SUB_RESE")
                Call fgAppendNode(xmlSubReserva, _
                                  "Grupo_PosicaoCaixaSubReserva", _
                                  "CO_SITU_CAIX_SUB_RESE", _
                                  enumEstadoCaixa.Disponivel)
            End If
            
            Exit For
        Next
    Else
        If Not pblnRetornaPosicaoDisponivel Then Exit Function
        
        'Posição de Caixa não encontrada.
        '------------------------------------------------------------------------------------------
        'Retorna a posição de caixa zerada,
        '          com data de abertura igual ao primeiro dia útil a partir da DATA DO SERVIDOR OU
        '          da DATA INFORMADA E situação do caixa igual a disponível
        If pstrDataCaixaSubReserva <> vbNullString Then
            strDataCaixaSubReserva = Format(fgAdicionarDiasUteis(fgDtXML_To_Date(pstrDataCaixaSubReserva), _
                                                                0, _
                                                                enumPaginacao.Proximo), _
                                                "yyyymmdd")
        Else
            strDataCaixaSubReserva = Format(fgAdicionarDiasUteis(fgDataHoraServidor(enumFormatoDataHora.Data), _
                                                                0, _
                                                                enumPaginacao.Proximo), _
                                                "yyyymmdd")
        End If
        
        'Prepara XML de retorno
        Call fgAppendNode(xmlSubReserva, "", "Repeat_PosicaoCaixaSubReserva", "")

        Call fgAppendNode(xmlSubReserva, _
                         "Repeat_PosicaoCaixaSubReserva", _
                         "Grupo_PosicaoCaixaSubReserva", "")
                         
        Call fgAppendNode(xmlSubReserva, _
                         "Grupo_PosicaoCaixaSubReserva", _
                         "DT_CAIX_SUB_RESE", strDataCaixaSubReserva)
        
        Call fgAppendNode(xmlSubReserva, _
                         "Grupo_PosicaoCaixaSubReserva", _
                         "CO_VEIC_LEGA", pstrCodigoVeiculoLegal)
                         
        Call fgAppendNode(xmlSubReserva, _
                         "Grupo_PosicaoCaixaSubReserva", _
                         "SG_SIST", Trim(UCase(pstrSiglaSistema)))
        
        Call fgAppendNode(xmlSubReserva, _
                         "Grupo_PosicaoCaixaSubReserva", _
                         "CO_SITU_CAIX_SUB_RESE", enumEstadoCaixa.Disponivel)
                         
        Call fgAppendNode(xmlSubReserva, _
                         "Grupo_PosicaoCaixaSubReserva", _
                         "VA_UTLZ_ABER_CAIX_SUB_RESE", 0)
                         
        Call fgAppendNode(xmlSubReserva, _
                         "Grupo_PosicaoCaixaSubReserva", _
                         "VA_ABER_RECE_CAIX_SUB_RESE", 0)
                         
        Call fgAppendNode(xmlSubReserva, _
                         "Grupo_PosicaoCaixaSubReserva", _
                         "VA_ABER_INFO_CAIX_SUB_RESE", 0)
                         
        Call fgAppendNode(xmlSubReserva, _
                         "Grupo_PosicaoCaixaSubReserva", _
                         "VA_FECH_CAIX_SUB_RESE", 0)
                         
        'Cria Nodes de data da posição e situação atual
        Call fgAppendNode(xmlSubReserva, _
                          "Grupo_PosicaoCaixaSubReserva", _
                          "DT_CAIX_SUB_RESE_ATUAL", _
                          strDataCaixaSubReserva)
        
        Call fgAppendNode(xmlSubReserva, _
                          "Grupo_PosicaoCaixaSubReserva", _
                          "CO_SITU_CAIX_SUB_RESE_ATUAL", _
                          enumEstadoCaixa.Disponivel)
    End If
    
    ObterPosicaoCaixaSubReserva = xmlSubReserva.xml
    
    Set xmlSubReserva = Nothing
    
    Exit Function

ErrorHandler:
    Set xmlSubReserva = Nothing
        
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPosicaoCaixaSubReserva Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Este método é responsável pela atualização da posição disponível para o caixa sub-reserva,
' e pelo retorno à camada intermediária.

Private Function flAtualizaPosicaoDisponivel(ByVal pstrCodigoVeiculoLegal As String, _
                                             ByVal pstrSiglaSistema As String, _
                                             ByVal pdatDataAberturaCaixa As Date, _
                                             ByVal pintSituacaoCaixa As Integer, _
                                             ByVal pvntValorAberturaCaixa As Variant)

Dim strSQL                                  As String
    
    On Error GoTo ErrorHandler
    
    strSQL = "UPDATE  A6.TB_POSI_CAIX_DISP                                                    " & vbNewLine & _
             "SET     CO_SITU_CAIX        = " & pintSituacaoCaixa & ",                        " & vbNewLine & _
             "        VA_UTLZ_ABER_CAIX   = " & fgVlrToDBServer(pvntValorAberturaCaixa) & ",  " & vbNewLine & _
             "        DT_CAIX_DISP        = " & _
                                        fgDtXML_To_Oracle(fgDt_To_Xml(pdatDataAberturaCaixa)) & vbNewLine & _
             "WHERE   SG_SIST             = '" & pstrSiglaSistema & "'                        " & vbNewLine & _
             "AND     CO_VEIC_LEGA        = '" & pstrCodigoVeiculoLegal & "'"
                                        
    If fgExecuteSQL(strSQL) = 0 Then
        strSQL = "INSERT INTO A6.TB_POSI_CAIX_DISP (                                                  " & vbNewLine & _
                 "        SG_SIST,                                                                    " & vbNewLine & _
                 "        CO_VEIC_LEGA,                                                               " & vbNewLine & _
                 "        DT_CAIX_DISP,                                                               " & vbNewLine & _
                 "        CO_SITU_CAIX,                                                               " & vbNewLine & _
                 "        VA_UTLZ_ABER_CAIX)                                                          " & vbNewLine & _
                 "VALUES ('" & pstrSiglaSistema & "',                                                 " & vbNewLine & _
                 "        '" & pstrCodigoVeiculoLegal & "',                                           " & vbNewLine & _
                 "        " & fgDtXML_To_Oracle(fgDt_To_Xml(pdatDataAberturaCaixa)) & ",              " & vbNewLine & _
                 "        " & pintSituacaoCaixa & ",                                                  " & vbNewLine & _
                 "        " & fgVlrToDBServer(pvntValorAberturaCaixa) & ")"
    
        Call fgExecuteSQL(strSQL)
    End If
    
    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAtualizaPosicaoDisponivel Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Este método é responsável pela leitura do histórico da posição do caixa,
' e pelo retorno à camada intermediária.

Public Function ObterHistoricoPosicaoCaixa(ByVal pxmlFiltros As String) As String


Dim xmlFiltro                               As DOMDocument40
Dim strSQL                                  As String
Dim intTipoBackOffice                       As Integer
    
On Error GoTo ErrorHandler

    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    xmlFiltro.loadXML pxmlFiltros
    
    strSQL = "SELECT b.no_veic_lega, " & vbCrLf & _
                    "a.dt_caix_sub_rese, " & vbCrLf & _
                    "a.co_situ_caix_sub_rese, " & vbCrLf & _
                    "DECODE(a.va_utlz_aber_caix_sub_rese, 0, DECODE(a.va_aber_rece_caix_sub_rese, 0, a.va_aber_info_caix_sub_rese, a.va_aber_rece_caix_sub_rese), a.va_utlz_aber_caix_sub_rese) VALOR_ABERTURA, " & vbCrLf & _
                    "a.va_movi_caix_sub_rese, " & vbCrLf & _
                    "a.va_fech_caix_sub_rese " & vbCrLf & _
               "FROM a6.tb_posi_caix_sub_rese A, " & vbCrLf & _
                    "a8.tb_veic_lega b " & vbCrLf & _
              "WHERE a.co_veic_lega = b.co_veic_lega " & vbCrLf
    
    strSQL = strSQL & "AND a.DT_CAIX_SUB_RESE " & _
                      "BETWEEN " & xmlFiltro.selectSingleNode("//Grupo_Data/DataIni").Text & _
                      " AND " & xmlFiltro.selectSingleNode("//Grupo_Data/DataFim").Text
    
    intTipoBackOffice = 99
    If Not xmlFiltro.selectSingleNode("//Grupo_BackOfficePerfilGeral/BackOfficePerfilGeral") Is Nothing Then
        intTipoBackOffice = xmlFiltro.selectSingleNode("//Grupo_BackOfficePerfilGeral/BackOfficePerfilGeral").Text
    End If
    
    Select Case intTipoBackOffice
        Case 99
            strSQL = strSQL & fgSegregaDados("A8.TB_VEIC_LEGA", False, "a", "b", False, , , False, , False)
        Case Is > 0
            strSQL = strSQL & _
                     vbNewLine & " AND B.TP_BKOF = " & intTipoBackOffice & vbNewLine
    End Select

    If Not xmlFiltro.selectSingleNode("//Grupo_VeiculoLegal/VeiculoLegal") Is Nothing Then
        strSQL = strSQL & "AND a.co_veic_lega = '" & _
                 xmlFiltro.selectSingleNode("//Grupo_VeiculoLegal/VeiculoLegal").Text & "'"
    Else
        If Not xmlFiltro.selectSingleNode("//Grupo_GrupoVeiculoLegal/GrupoVeiculoLegal") Is Nothing Then
            strSQL = strSQL & "AND b.co_grup_veic_lega = '" & _
                     xmlFiltro.selectSingleNode("//Grupo_GrupoVeiculoLegal/GrupoVeiculoLegal").Text & "'"
        End If
    End If
    
    'Verificação para a base historica
    If fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("//DataFim").Text, 10, 8)) >= (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        If fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("//DataIni").Text, 10, 8)) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
            strSQL = strSQL & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     Replace(strSQL, "a6.tb_posi_caix_sub_rese", "a6hist.tb_posi_caix_sub_rese")
        End If
    Else
        strSQL = Replace(strSQL, "a6.tb_posi_caix_sub_rese", "a6hist.tb_posi_caix_sub_rese")
    End If
    
    strSQL = strSQL & " ORDER BY no_veic_lega, dt_caix_sub_rese"
    
    Set xmlFiltro = Nothing

    ObterHistoricoPosicaoCaixa = fgQueryXMLLerTodos("HistoricoPosicaoCaixa", strSQL, App.EXEName & "." & TypeName(Me))
    
    Exit Function
ErrorHandler:
    Set xmlFiltro = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterHistoricoPosicaoCaixa Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
