VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParmHistCntaCntb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Objeto responsável pelo armazenamento dos métodos da tabela TB_PARM_HIST_CNTA_CNTB

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

'Incluir nova parametrização de Conta Contábil
Public Function flIncluir(ByRef pxmlParmHistCntaCntb As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim lngRetorno                              As Long

On Error GoTo ErrorHandler
    
    lngRetorno = flValidaRegras(pxmlParmHistCntaCntb)
    If lngRetorno <> 0 Then
        lngCodigoErroNegocio = lngRetorno
        GoTo ErrorHandler
    End If



    strSQL = " INSERT INTO A8.TB_PARM_HIST_CNTA_CNTB (  " & vbNewLine & _
             " SG_SIST,                          " & vbNewLine & _
             " CO_EMPR,                          " & vbNewLine & _
             " TP_BKOF,                          " & vbNewLine & _
             " CO_LOCA_LIQU,                     " & vbNewLine & _
             " IN_LANC_DEBT_CRED,                " & vbNewLine & _
             " CO_CNTA_DEBT,                     " & vbNewLine & _
             " CO_CNTA_CRED,                     " & vbNewLine & _
             " CO_HIST_CNTA_CNTB,                " & vbNewLine & _
             " DE_HIST_CNTA_CNTB,                " & vbNewLine & _
             " CO_CENT_DEST,                     " & vbNewLine & _
             " CO_USUA_ULTI_ATLZ,                " & vbNewLine & _
             " CO_ETCA_TRAB_ULTI_ATLZ,           " & vbNewLine & _
             " DH_ULTI_ATLZ,                      " & vbNewLine & _
             " TP_OPER                           " & vbNewLine & _
             " ) VALUES (                        " & vbNewLine

    With pxmlParmHistCntaCntb.documentElement
        strSQL = strSQL & _
             "'" & .selectSingleNode("SG_SIST").Text & "'       , " & vbNewLine & _
             CLng("0" & .selectSingleNode("CO_EMPR").Text) & " , " & vbNewLine & _
             CLng("0" & .selectSingleNode("TP_BKOF").Text) & " , " & vbNewLine & _
             CLng("0" & .selectSingleNode("CO_LOCA_LIQU").Text) & " , " & vbNewLine & _
             CLng("0" & .selectSingleNode("IN_LANC_DEBT_CRED").Text) & " , " & vbNewLine & _
             CLng("0" & .selectSingleNode("CO_CNTA_DEBT").Text) & "      , " & vbNewLine & _
             CLng("0" & .selectSingleNode("CO_CNTA_CRED").Text) & "      , " & vbNewLine & _
             "'" & .selectSingleNode("CO_HIST_CNTA_CNTB").Text & "'      , " & vbNewLine & _
             "'" & .selectSingleNode("DE_HIST_CNTA_CNTB").Text & "'      , " & vbNewLine & _
             " " & .selectSingleNode("CO_CENT_DEST").Text & "            , " & vbNewLine & _
             "'" & fgUsuarioRede & "'                         , " & vbNewLine & _
             "'" & fgEstacaoTrabalhoUsuario & "'              , " & vbNewLine & _
             " SYSDATE, " & vbNewLine & _
             CLng("0" & .selectSingleNode("TP_OPER").Text) & " ) "
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

'Alterar parametrização de conta contábil
Private Function flAlterar(ByRef pxmlParmHistCntaCntb As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim lngRetorno                              As Long

On Error GoTo ErrorHandler
    
    lngRetorno = flValidaRegras(pxmlParmHistCntaCntb)
    If lngRetorno <> 0 Then
        lngCodigoErroNegocio = lngRetorno
        GoTo ErrorHandler
    End If
    
    With pxmlParmHistCntaCntb.documentElement
    strSQL = " UPDATE A8.TB_PARM_HIST_CNTA_CNTB      " & vbNewLine & _
             " SET    CO_CNTA_DEBT = " & CLng("0" & .selectSingleNode("CO_CNTA_DEBT").Text) & " , " & vbNewLine & _
             "        CO_CNTA_CRED = " & CLng("0" & .selectSingleNode("CO_CNTA_CRED").Text) & " , " & vbNewLine & _
             "        CO_HIST_CNTA_CNTB = '" & .selectSingleNode("CO_HIST_CNTA_CNTB").Text & "' , " & vbNewLine & _
             "        DE_HIST_CNTA_CNTB = '" & .selectSingleNode("DE_HIST_CNTA_CNTB").Text & "'           , " & vbNewLine & _
             "        CO_CENT_DEST = " & CLng("0" & .selectSingleNode("CO_CENT_DEST").Text) & " , " & vbNewLine & _
             "        CO_USUA_ULTI_ATLZ = '" & fgUsuarioRede & "'                           , " & vbNewLine & _
             "        CO_ETCA_TRAB_ULTI_ATLZ = '" & fgEstacaoTrabalhoUsuario & "'           , " & vbNewLine & _
             "        DH_ULTI_ATLZ           =  SYSDATE                                       " & vbNewLine & _
             " WHERE  SG_SIST = '" & .selectSingleNode("SG_SIST").Text & "'                   " & vbNewLine & _
             " AND    CO_EMPR = " & .selectSingleNode("CO_EMPR").Text & "                     " & vbNewLine & _
             " AND    TP_BKOF = " & .selectSingleNode("TP_BKOF").Text & "                     " & vbNewLine & _
             " AND    CO_LOCA_LIQU = " & .selectSingleNode("CO_LOCA_LIQU").Text & "                     " & vbNewLine & _
             " AND    IN_LANC_DEBT_CRED = " & .selectSingleNode("IN_LANC_DEBT_CRED").Text & " " & vbNewLine & _
             " AND    DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text)
             
    strSQL = strSQL & " AND TP_OPER = " & .selectSingleNode("TP_OPER").Text & " "
    
            
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

'Executa a operação de inclusão, alteração ou exclusão de conta contábil
Public Function Salvar(ByVal pxmlParmHistCntaCntb As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    If pxmlParmHistCntaCntb.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pxmlParmHistCntaCntb
    ElseIf pxmlParmHistCntaCntb.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        flAlterar pxmlParmHistCntaCntb
    ElseIf pxmlParmHistCntaCntb.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        flExcluir pxmlParmHistCntaCntb
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

'Exclui uma parametrização de conta contábil
Private Function flExcluir(ByRef pxmlParmHistCntaCntb As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    With pxmlParmHistCntaCntb.documentElement

        strSQL = " DELETE " & strSQL & _
                 " FROM A8.TB_PARM_HIST_CNTA_CNTB " & _
                 " WHERE  SG_SIST = '" & .selectSingleNode("SG_SIST").Text & "'                   " & vbNewLine & _
                 " AND    CO_EMPR = " & .selectSingleNode("CO_EMPR").Text & "                     " & vbNewLine & _
                 " AND    TP_BKOF = " & .selectSingleNode("TP_BKOF").Text & "                     " & vbNewLine & _
                 " AND    CO_LOCA_LIQU = " & .selectSingleNode("CO_LOCA_LIQU").Text & "                     " & vbNewLine & _
                 " AND    IN_LANC_DEBT_CRED = " & .selectSingleNode("IN_LANC_DEBT_CRED").Text & " " & vbNewLine & _
                 " AND    DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text) & " " & vbNewLine & _
                 " AND    TP_OPER = " & .selectSingleNode("TP_OPER").Text
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

'Obter parametrizações de conta contabil
Public Function ObterPropriedades() As String

Dim strSQL                                  As String


    On Error GoTo ErrorHandler

    strSQL = " SELECT SG_SIST,             " & vbNewLine & _
             "        CO_EMPR,             " & vbNewLine & _
             "        TP_BKOF,             " & vbNewLine & _
             "        CO_LOCA_LIQU,        " & vbNewLine & _
             "        IN_LANC_DEBT_CRED,   " & vbNewLine & _
             "        CO_CNTA_DEBT,        " & vbNewLine & _
             "        CO_CNTA_CRED,        " & vbNewLine & _
             "        CO_HIST_CNTA_CNTB,   " & vbNewLine & _
             "        DE_HIST_CNTA_CNTB,   " & vbNewLine & _
             "        CO_CENT_DEST,        " & vbNewLine & _
             "        DH_ULTI_ATLZ,        " & vbNewLine & _
             "        TP_OPER              " & vbNewLine & _
             " FROM   A8.TB_PARM_HIST_CNTA_CNTB  "

    ObterPropriedades = fgPropriedades("Grupo_ParmHistCntaCntb", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Ler todas as parametrizações de conta contabil
Public Function LerTodos() As String

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String



On Error GoTo ErrorHandler

    strSQL = " SELECT A.SG_SIST,                    " & vbNewLine & _
             "        NO_SIST,                    " & vbNewLine & _
             "        A.CO_EMPR,                    " & vbNewLine & _
             "        TP_BKOF,                    " & vbNewLine & _
             "        CO_LOCA_LIQU,               " & vbNewLine & _
             "        IN_LANC_DEBT_CRED,          " & vbNewLine & _
             "        CO_CNTA_DEBT,               " & vbNewLine & _
             "        CO_CNTA_CRED,               " & vbNewLine & _
             "        CO_HIST_CNTA_CNTB,          " & vbNewLine & _
             "        DE_HIST_CNTA_CNTB,          " & vbNewLine & _
             "        CO_CENT_DEST,               " & vbNewLine & _
             "        A.DH_ULTI_ATLZ,             " & vbNewLine & _
             "        A.TP_OPER                   " & vbNewLine & _
             " FROM   A8.TB_PARM_HIST_CNTA_CNTB A," & vbNewLine & _
             "        A7.TB_SIST                B," & vbNewLine & _
             "        A8.TB_EMPRESA_HO          C " & vbNewLine & _
             " WHERE  A.CO_EMPR = B.CO_EMPR       " & vbNewLine & _
             " AND    A.CO_EMPR = C.CO_EMPR       " & vbNewLine & _
             " AND    A.SG_SIST = B.SG_SIST       " & vbNewLine

    strWhereSegregacao = fgSegregaDados("A8.TB_PARM_HIST_CNTA_CNTB", _
                                        False, _
                                        "A", _
                                        "A", _
                                        False, _
                                        False, _
                                        True, _
                                        False, _
                                        False, _
                                        False)
                                            
    strSQL = strSQL & strWhereSegregacao
                                            
    strSQL = strSQL & _
                    " AND  (C.DT_FIM_VIGE    IS NULL " & vbNewLine & _
                    "  OR   C.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")" & vbNewLine & _
                    " AND   C.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & vbNewLine

    strSQL = strSQL & " order by A.CO_EMPR,SG_SIST,CO_LOCA_LIQU"

        LerTodos = fgQueryXMLLerTodos("ParmHistCntaCntb", strSQL, App.EXEName & "." & TypeName(Me))

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

'Ler detalhes de uma parametrização de conta contábil
Public Function Ler(ByVal pstrSiglaSistema As String, _
                    ByVal plngCodigoEmpresa As Long, _
                    ByVal plngTipoBackOffice As Long, _
                    ByVal plngLocalLiquidacao As Long, _
                    ByVal plngDebitoCredito As Long, _
                    Optional ByVal plngTipoOperacao As Long = 0) As String


Dim strSQL                                  As String



    On Error GoTo ErrorHandler

    strSQL = " SELECT   SG_SIST,                   " & vbNewLine & _
             "          CO_EMPR,                   " & vbNewLine & _
             "          TP_BKOF,                   " & vbNewLine & _
             "          CO_LOCA_LIQU,              " & vbNewLine & _
             "          IN_LANC_DEBT_CRED,         " & vbNewLine & _
             "          CO_CNTA_DEBT,              " & vbNewLine & _
             "          CO_CNTA_CRED,              " & vbNewLine & _
             "          CO_HIST_CNTA_CNTB,         " & vbNewLine & _
             "          DE_HIST_CNTA_CNTB,         " & vbNewLine & _
             "          CO_CENT_DEST,              " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ,         " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ,    " & vbNewLine & _
             "          DH_ULTI_ATLZ,              " & vbNewLine & _
             "          TP_OPER                    " & vbNewLine & _
             " FROM     A8.TB_PARM_HIST_CNTA_CNTB  " & vbNewLine & _
             " WHERE    SG_SIST             =   '" & pstrSiglaSistema & "'" & vbNewLine & _
             " AND      CO_EMPR             =   " & plngCodigoEmpresa & vbNewLine & _
             " AND      TP_BKOF             =   " & plngTipoBackOffice & vbNewLine & _
             " AND      CO_LOCA_LIQU        =   " & plngLocalLiquidacao & vbNewLine & _
             " AND      IN_LANC_DEBT_CRED   =   " & plngDebitoCredito & vbNewLine & _
             " AND      TP_OPER = " & plngTipoOperacao
    
    Ler = fgQueryXMLLer("ParmHistCntaCntb", strSQL, App.EXEName & "." & TypeName(Me))

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

'Executa uma operação em conta contábil, seja de leitura ou manutenção
Public Function Executar(ByVal pxmlParmHistCntaCntb As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlParmHistCntaCntb.documentElement.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Excluir", "Alterar"
            Salvar pxmlParmHistCntaCntb
        Case "Ler"
            With pxmlParmHistCntaCntb.documentElement
                Executar = Ler(.selectSingleNode("SG_SIST").Text, _
                               CLng("0" & .selectSingleNode("CO_EMPR").Text), _
                               CLng("0" & .selectSingleNode("TP_BKOF").Text), _
                               CLng("0" & .selectSingleNode("CO_LOCA_LIQU").Text), _
                               CLng("0" & .selectSingleNode("IN_LANC_DEBT_CRED").Text), _
                               CLng("0" & .selectSingleNode("TP_OPER").Text))
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

'Valida as regras de uma parametrização de conta contábil
Private Function flValidaRegras(ByRef pxmlParmHistCntaCntb As MSXML2.DOMDocument40) As Long

Dim strSQL                                  As String
Dim rsParam                                 As ADODB.Recordset
Dim intIndicadorDebitoCredito               As Integer

On Error GoTo ErrorHandler
    
    'Pikachu - 09/06/2004
    'Desabilitada a validação de Codigo Conta Contabil Debito e Credito
    'para a chave :  Empresa, Tipo Operação e Tipo Lançamento Contábil.
    flValidaRegras = 0
    
    Exit Function
    
    With pxmlParmHistCntaCntb.documentElement
    
        If Val(.selectSingleNode("IN_LANC_DEBT_CRED").Text) = enumTipoDebitoCreditoEstorno.EstornoCredito Or _
           Val(.selectSingleNode("IN_LANC_DEBT_CRED").Text) = enumTipoDebitoCreditoEstorno.EstornoDebito Then
           flValidaRegras = 0
           Exit Function
        End If
        
        'Pesquisa se já existe um cadastramento para a Contra-Partida do lançamento
        intIndicadorDebitoCredito = IIf(Val(.selectSingleNode("IN_LANC_DEBT_CRED").Text) = 1, 2, 1)
        
        strSQL = " SELECT   CO_CNTA_DEBT,           " & vbNewLine & _
                 "          CO_CNTA_CRED            " & vbNewLine & _
                 " FROM     A8.TB_PARM_HIST_CNTA_CNTB  " & vbNewLine & _
                 " WHERE    SG_SIST             =   '" & .selectSingleNode("SG_SIST").Text & "'" & vbNewLine & _
                 " AND      CO_EMPR             =   " & .selectSingleNode("CO_EMPR").Text & vbNewLine & _
                 " AND      TP_BKOF             =   " & .selectSingleNode("TP_BKOF").Text & vbNewLine & _
                 " AND      CO_LOCA_LIQU        =   " & .selectSingleNode("CO_LOCA_LIQU").Text & vbNewLine & _
                 " AND      IN_LANC_DEBT_CRED   =   " & intIndicadorDebitoCredito
        
        Set rsParam = fgQuerySQL(strSQL)
        
        If Not rsParam.EOF Then
            If rsParam("CO_CNTA_DEBT") <> Val("0" & .selectSingleNode("CO_CNTA_DEBT").Text) Or _
               rsParam("CO_CNTA_CRED") <> Val("0" & .selectSingleNode("CO_CNTA_CRED").Text) Then
                '3047 - Contas Débito e Crédito devem ser iguais para a chave
                '      (Empresa, Tipo Operação e Tipo Lançamento Contábil).
                flValidaRegras = 3047
                
                rsParam.Close
                Set rsParam = Nothing
                
                Exit Function
            End If
        End If
    End With
    
    'Resultado OK
    flValidaRegras = 0
    
    rsParam.Close
    Set rsParam = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    Set rsParam = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flValidaRegras Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
