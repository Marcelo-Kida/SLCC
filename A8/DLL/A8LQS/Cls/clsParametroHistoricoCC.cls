VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParametroHistoricoCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelos métodos de controle do parâmetro de históricos de Conta Corrente.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private objContext                       As COMSVCSLib.ObjectContext
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

'Inclui uma parametrização de histórico de conta corrente
Public Function flIncluir(ByRef pxmlParametroHistoricoCC As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim strCodSubTipo                           As String
Dim strFinalidadeTED                        As String
Dim strCodProd                              As String

On Error GoTo ErrorHandler

    flConsiste pxmlParametroHistoricoCC

    strSQL = " INSERT INTO A8.TB_PARM_HIST_CC (  " & vbNewLine & _
             " SG_SIST,                          " & vbNewLine & _
             " CO_EMPR,                          " & vbNewLine & _
             " TP_BKOF,                          " & vbNewLine & _
             " TP_OPER,                          " & vbNewLine & _
             " IN_LANC_DEBT_CRED,                " & vbNewLine & _
             " CO_HIST_CC,                       " & vbNewLine & _
             " CO_SUB_TIPO_ATIV,                 " & vbNewLine & _
             " CO_USUA_ULTI_ATLZ,                " & vbNewLine & _
             " CO_ETCA_TRAB_ULTI_ATLZ,           " & vbNewLine & _
             " DH_ULTI_ATLZ,                     " & vbNewLine & _
             " CD_FIND_TED,                      " & vbNewLine & _
             " CD_SUB_PROD,                      " & vbNewLine & _
             " TP_CNAL_VEND                      " & vbNewLine & _
             " ) VALUES (                        " & vbNewLine

    With pxmlParametroHistoricoCC.documentElement
    
        strCodSubTipo = .selectSingleNode("CO_SUB_TIPO_ATIV").Text
        If strCodSubTipo = vbNullString Then
            strCodSubTipo = "0"
        End If
        
        strFinalidadeTED = .selectSingleNode("CD_FIND_TED").Text
        If strFinalidadeTED = vbNullString Then
            strFinalidadeTED = "0"
        End If
        
        strSQL = strSQL & _
                 "'" & .selectSingleNode("SG_SIST").Text & "'                , " & vbNewLine & _
                 CLng("0" & .selectSingleNode("CO_EMPR").Text) & "           , " & vbNewLine & _
                 CLng("0" & .selectSingleNode("TP_BKOF").Text) & "           , " & vbNewLine & _
                 CLng("0" & .selectSingleNode("TP_OPER").Text) & "           , " & vbNewLine & _
                 CLng("0" & .selectSingleNode("IN_LANC_DEBT_CRED").Text) & " , " & vbNewLine & _
                 CLng("0" & .selectSingleNode("CO_HIST_CC").Text) & "        , " & vbNewLine & _
                 "'" & strCodSubTipo & "'                                    , " & vbNewLine & _
                 "'" & fgUsuarioRede & "'                                    , " & vbNewLine & _
                 "'" & fgEstacaoTrabalhoUsuario & "'                         , " & vbNewLine & _
                 " SYSDATE                                                   , " & vbNewLine & _
                 CLng("0" & strFinalidadeTED) & "                            , " & vbNewLine
                 If Not .selectSingleNode("CD_SUB_PROD") Is Nothing Then
                     strSQL = strSQL & "'" & .selectSingleNode("CD_SUB_PROD").Text & "' , " & vbNewLine
                 Else
                     strSQL = strSQL & " '0' , " & vbNewLine
                 End If
                 strSQL = strSQL & CLng("0" & .selectSingleNode("TP_CNAL_VEND").Text) & " ) "
    End With

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

'Altera uma parametrização de histórico de conta corrente
Private Function flAlterar(ByRef pxmlParametroHistoricoCC As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    flConsiste pxmlParametroHistoricoCC

    With pxmlParametroHistoricoCC.documentElement
        strSQL = " UPDATE A8.TB_PARM_HIST_CC                                                                       " & vbNewLine & _
                 " SET    CO_HIST_CC = " & CLng("0" & .selectSingleNode("CO_HIST_CC").Text) & " ,                  " & vbNewLine & _
                 "        CO_USUA_ULTI_ATLZ = '" & fgUsuarioRede & "'                           ,                  " & vbNewLine & _
                 "        CO_ETCA_TRAB_ULTI_ATLZ = '" & fgEstacaoTrabalhoUsuario & "'           ,                  " & vbNewLine & _
                 "        DH_ULTI_ATLZ           =  SYSDATE                                                        " & vbNewLine & _
                 " WHERE  SG_SIST           = '" & .selectSingleNode("SG_SIST").Text & "'                          " & vbNewLine & _
                 " AND    CO_EMPR           =  " & .selectSingleNode("CO_EMPR").Text & "                           " & vbNewLine & _
                 " AND    TP_BKOF           =  " & .selectSingleNode("TP_BKOF").Text & "                           " & vbNewLine & _
                 " AND    TP_OPER           =  " & .selectSingleNode("TP_OPER").Text & "                           " & vbNewLine & _
                 " AND    IN_LANC_DEBT_CRED =  " & .selectSingleNode("IN_LANC_DEBT_CRED").Text & "                 " & vbNewLine & _
                 " AND    DH_ULTI_ATLZ      =  " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text) & " " & vbNewLine & _
                 " AND    CO_SUB_TIPO_ATIV  = '" & .selectSingleNode("CO_SUB_TIPO_ATIV").Text & "'                 " & vbNewLine & _
                 " AND    CD_SUB_PROD       = '" & .selectSingleNode("CD_SUB_PROD").Text & "'                      " & vbNewLine & _
                 " AND    CD_FIND_TED       =  " & .selectSingleNode("CD_FIND_TED").Text & "                       " & vbNewLine & _
                 " AND    TP_CNAL_VEND      =  " & .selectSingleNode("TP_CNAL_VEND").Text
    End With
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Salva uma alteração de parametrização de histórico de conta corrente, seja de inclusão, alteração ou exclusão
Public Function Salvar(ByVal pxmlParametroHistoricoCC As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    If pxmlParametroHistoricoCC.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pxmlParametroHistoricoCC
    ElseIf pxmlParametroHistoricoCC.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        flAlterar pxmlParametroHistoricoCC
    ElseIf pxmlParametroHistoricoCC.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        flExcluir pxmlParametroHistoricoCC
    Else
        ' 7 - Operação Inválida
        lngCodigoErroNegocio = 7
        GoTo ErrorHandler
    End If
        
    Salvar = True
      
    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Exclui uma parametrização de histórico de conta corrente
Private Function flExcluir(ByRef pxmlParametroHistoricoCC As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    With pxmlParametroHistoricoCC.documentElement

        strSQL = " DELETE " & strSQL & _
                 " FROM A8.TB_PARM_HIST_CC " & _
                 " WHERE  SG_SIST           = '" & .selectSingleNode("SG_SIST").Text & "'                          " & vbNewLine & _
                 " AND    CO_EMPR           =  " & .selectSingleNode("CO_EMPR").Text & "                           " & vbNewLine & _
                 " AND    TP_BKOF           =  " & .selectSingleNode("TP_BKOF").Text & "                           " & vbNewLine & _
                 " AND    TP_OPER           =  " & .selectSingleNode("TP_OPER").Text & "                           " & vbNewLine & _
                 " AND    IN_LANC_DEBT_CRED =  " & .selectSingleNode("IN_LANC_DEBT_CRED").Text & "                 " & vbNewLine & _
                 " AND    DH_ULTI_ATLZ      =  " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text) & " " & vbNewLine & _
                 " AND    CO_SUB_TIPO_ATIV  = '" & .selectSingleNode("CO_SUB_TIPO_ATIV").Text & "'  " & vbNewLine & _
                 " AND    CD_FIND_TED       = " & .selectSingleNode("CD_FIND_TED").Text & vbNewLine & _
                 " AND    CD_SUB_PROD       = '" & .selectSingleNode("CD_SUB_PROD").Text & "' " & vbNewLine & _
                 " AND    TP_CNAL_VEND      = " & .selectSingleNode("TP_CNAL_VEND").Text
                 
    End With

    If fgExecuteSQL(strSQL) = 0 Then
        '4002 - Registro Inexistente ou excluído por outro usuário.
        lngCodigoErroNegocio = 4002
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter todas as parametrizações de histórico de conta corrente
Public Function ObterPropriedades() As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT SG_SIST,             " & vbNewLine & _
             "        CO_EMPR,             " & vbNewLine & _
             "        TP_BKOF,             " & vbNewLine & _
             "        TP_OPER,             " & vbNewLine & _
             "        IN_LANC_DEBT_CRED,   " & vbNewLine & _
             "        CO_HIST_CC,          " & vbNewLine & _
             "        DH_ULTI_ATLZ,        " & vbNewLine & _
             "        CO_SUB_TIPO_ATIV,    " & vbNewLine & _
             "        CD_FIND_TED,         " & vbNewLine & _
             "        CD_SUB_PROD,         " & vbNewLine & _
             "        TP_CNAL_VEND         " & vbNewLine & _
             " FROM   A8.TB_PARM_HIST_CC   "

    ObterPropriedades = fgPropriedades("Grupo_ParametroHistoricoCC", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter todas as parametrizações de histórico de conta corrente, utilizando a segregação de acesso
Public Function LerTodos() As String

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT SG_SIST,             " & vbNewLine & _
             "        A.CO_EMPR,             " & vbNewLine & _
             "        TP_BKOF,             " & vbNewLine & _
             "        TP_OPER,             " & vbNewLine & _
             "        IN_LANC_DEBT_CRED,   " & vbNewLine & _
             "        CO_HIST_CC,          " & vbNewLine & _
             "        A.DH_ULTI_ATLZ,        " & vbNewLine & _
             "        CO_SUB_TIPO_ATIV,    " & vbNewLine & _
             "        CD_FIND_TED,         " & vbNewLine & _
             "        CD_SUB_PROD,         " & vbNewLine & _
             "        TP_CNAL_VEND         " & vbNewLine & _
             " FROM   A8.TB_PARM_HIST_CC   A," & vbNewLine & _
             "        A8.TB_EMPRESA_HO     B " & vbNewLine
    
    'RATS - 935 - inclusão da vigência da tabela A8.TB_EMPRESA_HO
    
    strWhereSegregacao = fgSegregaDados("A8.TB_PARM_HIST_CC", _
                                        False, _
                                        "A8.TB_PARM_HIST_CC", _
                                        "A", _
                                        True, _
                                        False, _
                                        True, _
                                        False, _
                                        False, _
                                        False)
                                            
    strSQL = strSQL & strWhereSegregacao
    
    strSQL = strSQL & _
                    " AND   SG_SIST = 'BG' " & _
                    " AND   A.CO_EMPR = B.CO_EMPR " & _
                    " AND  (DT_FIM_VIGE    IS NULL " & _
                    "  OR   DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")" & _
                    " AND   DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                    " ORDER BY SG_SIST,CO_EMPR,TP_OPER "

    LerTodos = fgQueryXMLLerTodos("ParametroHistoricoCC", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter dados de uma parametrização de histórico de conta corrente
Public Function Ler(ByVal pstrSiglaSistema As String, _
                    ByVal plngCodigoEmpresa As Long, _
                    ByVal plngTipoBackOffice As Long, _
                    ByVal plngTipoOperacao As Long, _
                    ByVal plngDebitoCredito As Long, _
                    ByVal pstrSubTipoAtivo As String, _
           Optional ByVal plngFinalidadeTED As Long = 0, _
           Optional ByVal plngCodProduto As String = "0", _
           Optional ByVal plngTipoCanalVenda As Long = 0) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT   SG_SIST,                " & vbNewLine & _
             "          CO_EMPR,                " & vbNewLine & _
             "          TP_BKOF,                " & vbNewLine & _
             "          TP_OPER,                " & vbNewLine & _
             "          IN_LANC_DEBT_CRED,      " & vbNewLine & _
             "          CO_HIST_CC,             " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ,      " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
             "          DH_ULTI_ATLZ,           " & vbNewLine & _
             "          CO_SUB_TIPO_ATIV,       " & vbNewLine & _
             "          CD_FIND_TED,            " & vbNewLine & _
             "          CD_SUB_PROD,            " & vbNewLine & _
             "          TP_CNAL_VEND            " & vbNewLine & _
             " FROM     A8.TB_PARM_HIST_CC      " & vbNewLine & _
             " WHERE    SG_SIST             =  '" & pstrSiglaSistema & "'" & vbNewLine & _
             " AND      CO_EMPR             =   " & plngCodigoEmpresa & vbNewLine & _
             " AND      TP_BKOF             =   " & plngTipoBackOffice & vbNewLine & _
             " AND      TP_OPER             =   " & plngTipoOperacao & vbNewLine & _
             " AND      IN_LANC_DEBT_CRED   =   " & plngDebitoCredito & vbNewLine
             
    If pstrSubTipoAtivo = vbNullString Then
        strSQL = strSQL & _
             " AND     (CO_SUB_TIPO_ATIV    =  '" & pstrSubTipoAtivo & "'" & _
             " OR       CO_SUB_TIPO_ATIV    =  '0'" & _
             " OR       CO_SUB_TIPO_ATIV    IS  NULL)"
    Else
        strSQL = strSQL & _
             " AND      CO_SUB_TIPO_ATIV    =  '" & pstrSubTipoAtivo & "'"
    End If
    
    If plngFinalidadeTED = 0 Then
        strSQL = strSQL & _
            " AND       (CD_FIND_TED = 0 OR CD_FIND_TED IS NULL) "
    Else
        strSQL = strSQL & _
            " AND       CD_FIND_TED = " & plngFinalidadeTED & " "
    End If
    
    If plngCodProduto = "0" Or plngCodProduto = "" Then
        strSQL = strSQL & _
            " AND       (CD_SUB_PROD = '0'  OR CD_SUB_PROD IS NULL OR CD_SUB_PROD=' ' ) "
    Else
        strSQL = strSQL & _
            " AND       CD_SUB_PROD = '" & plngCodProduto & "' "
    End If
    
    If plngTipoCanalVenda = 0 Then
        strSQL = strSQL & _
            " AND TP_CNAL_VEND = 0 "
    Else
        strSQL = strSQL & _
            " AND TP_CNAL_VEND = " & plngTipoCanalVenda
    End If
    
    Ler = fgQueryXMLLer("ParametroHistoricoCC", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Executa uma tarefa em parametrização de histórico de conta corrente, seja de leitura ou alteração
Public Function Executar(ByVal pxmlParametroHistoricoCC As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlParametroHistoricoCC.documentElement.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Excluir", "Alterar"
            Salvar pxmlParametroHistoricoCC
        Case "Ler"
            With pxmlParametroHistoricoCC.documentElement
                Executar = Ler(.selectSingleNode("SG_SIST").Text, _
                               CLng("0" & .selectSingleNode("CO_EMPR").Text), _
                               CLng("0" & .selectSingleNode("TP_BKOF").Text), _
                               CLng("0" & .selectSingleNode("TP_OPER").Text), _
                               CLng("0" & .selectSingleNode("IN_LANC_DEBT_CRED").Text), _
                               .selectSingleNode("CO_SUB_TIPO_ATIV").Text, _
                               CLng("0" & .selectSingleNode("CD_FIND_TED").Text), _
                               .selectSingleNode("CD_SUB_PROD").Text, _
                               CLng("0" & .selectSingleNode("TP_CNAL_VEND").Text) _
                               )
            End With
        Case "LerTodos"
            Executar = LerTodos()
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
    End Select
    
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Valida dados de uma parametrização de histórico de conta corrente
Private Function flConsiste(ByRef pxmlParametroHistoricoCC As MSXML2.DOMDocument40) As String

On Error GoTo ErrorHandler

    With pxmlParametroHistoricoCC.documentElement
         If CLng(.selectSingleNode("CO_EMPR").Text) = enumCodigoEmpresa.Banespa Then
             If Len(Trim(.selectSingleNode("CO_HIST_CC").Text)) > 4 Then
                ' 3048 - Tamanho máximo do Código do Histórico para o Banespa é 4 posições.
                lngCodigoErroNegocio = 3048
                GoTo ErrorHandler
             End If
         End If
    End With

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConsiste Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
