VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Objeto responsável pela manutenção e reperação de registros da tabela de tipo de operação (A8.TB_TIPO_OPER)


Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private ObjectContext                       As COMSVCSLib.ObjectContext
Implements COMSVCSLib.ObjectControl

Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
End Sub

'Recuperar registros da tabela de tipo de operação

Public Function LerTodos(Optional ByVal pstrIndicadorVigencia As String = "S", _
                         Optional ByVal plngCodigoOperacaoSelic As Long = 0) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " Select    a.TP_OPER,                " & vbNewLine & _
             "           a.NO_TIPO_OPER,           " & vbNewLine & _
             "           a.TP_MESG_RECB_INTE,      " & vbNewLine & _
             "           a.TP_MESG_RETN_INTE,      " & vbNewLine & _
             "           a.CO_MESG_SPB_REGT_OPER,  " & vbNewLine & _
             "           a.CO_OPER_SELIC,      " & vbNewLine & _
             "           a.DT_INIC_VIGE,           " & vbNewLine & _
             "           a.DT_FIM_VIGE,            " & vbNewLine & _
             "           a.CO_LOCA_LIQU,           " & vbNewLine & _
             "           b.SG_LOCA_LIQU,           " & vbNewLine & _
             "           a.CO_USUA_ULTI_ATLZ,      " & vbNewLine & _
             "           a.CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
             "           a.DH_ULTI_ATLZ            " & vbNewLine & _
             " From      A8.TB_TIPO_OPER        a, " & vbNewLine & _
             "           A8.TB_LOCAL_LIQUIDACAO b  " & vbNewLine

    strSQL = strSQL & _
            " WHERE      a.CO_LOCA_LIQU = b.CO_LOCA_LIQU (+) " & _
            " AND        1              = b.CO_EMPR_FUSI (+) "

    If pstrIndicadorVigencia = "S" Then
        strSQL = strSQL & " AND     (  a.DT_FIM_VIGE IS  NULL     " & vbNewLine & _
                          " OR         a.DT_FIM_VIGE >=  SYSDATE) "
    End If

    If plngCodigoOperacaoSelic <> 0 Then
        strSQL = strSQL & " And     a.CO_OPER_SELIC   =   " & Trim(plngCodigoOperacaoSelic)
    End If

    strSQL = strSQL & " Order By  a.TP_OPER ASC"

    LerTodos = fgQueryXMLLerTodos("TipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter os nomes da colunas da tabela de tipo de operação

Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   TP_OPER, " & _
            "           NO_TIPO_OPER, " & _
            "           TP_MESG_RECB_INTE, " & _
            "           TP_MESG_RETN_INTE, " & _
            "           CO_MESG_SPB_REGT_OPER, " & _
            "           DT_INIC_VIGE, " & _
            "           DT_FIM_VIGE, " & _
            "           CO_USUA_ULTI_ATLZ, " & _
            "           CO_ETCA_TRAB_ULTI_ATLZ, " & _
            "           DH_ULTI_ATLZ, " & _
            "           '' AS TP_VIGE " & _
            " FROM      A8.TB_TIPO_OPER "
    
    ObterPropriedades = fgPropriedades("Grupo_TipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Gerenciar as operações de Inclusção,Alteração,Exclusão,Ler e LerTodos da tabela de tipo de operação

Public Function Executar(ByVal pxmlTipoOperacao As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = vbNullString
    
    If Not pxmlTipoOperacao.selectSingleNode("//@Operacao") Is Nothing Then
        strOperacao = pxmlTipoOperacao.selectSingleNode("//@Operacao").Text
    End If

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pxmlTipoOperacao)
        Case "Ler"
            Executar = Ler(pxmlTipoOperacao.documentElement.selectSingleNode("TP_OPER").Text)
        Case "LerTodos"
            If pxmlTipoOperacao.selectSingleNode("//TP_VIGE") Is Nothing Then
                Executar = LerTodos()
            Else
                Executar = LerTodos(pxmlTipoOperacao.selectSingleNode("//TP_VIGE").Text)
            End If
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

'Incluir um registro na tabela de tipo de operação

Private Function flIncluir(ByRef pxmlTipoOperacao As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A8.TB_TIPO_OPER (" & _
             " TP_OPER, " & _
             " NO_TIPO_OPER, " & _
             " TP_MESG_RECB_INTE, " & _
             " TP_MESG_RETN_INTE, " & _
             " CO_MESG_SPB_REGT_OPER, " & _
             " DT_INIC_VIGE, " & _
             " DT_FIM_VIGE, " & _
             " CO_USUA_ULTI_ATLZ, " & _
             " CO_ETCA_TRAB_ULTI_ATLZ, " & _
             "  DH_ULTI_ATLZ " & _
             " ) VALUES ( " & _
             CLng(pxmlTipoOperacao.documentElement.selectSingleNode("TP_OPER").Text) & "," & _
             "'" & pxmlTipoOperacao.documentElement.selectSingleNode("NO_TIPO_OPER").Text & "'," & _
             "'" & CLng(pxmlTipoOperacao.documentElement.selectSingleNode("TP_MESG_RECB_INTE").Text) & "'," & _
             CLng(pxmlTipoOperacao.documentElement.selectSingleNode("TP_MESG_RETN_INTE").Text) & "," & _
             "'" & pxmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text & "'," & _
             fgDtXML_To_Oracle(pxmlTipoOperacao.documentElement.selectSingleNode("DT_INIC_VIGE").Text) & "," & _
             fgDtXML_To_Oracle(pxmlTipoOperacao.documentElement.selectSingleNode("DT_FIM_VIGE").Text) & "," & _
             "'" & fgObterUsuarioRede & "'," & _
             "'" & fgObterEstacaoTrabalhoUsuario & "'," & _
             "SYSDATE)"

    fgExecuteSQL strSQL
     
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Alterar um registro na tabela de tipo de operação

Private Function flAlterar(ByRef pxmlTipoOperacao As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler
    
    strSQL = " UPDATE A8.TB_TIPO_OPER " & _
            "  SET    NO_TIPO_OPER           = '" & pxmlTipoOperacao.documentElement.selectSingleNode("NO_TIPO_OPER").Text & "'," & _
            "         TP_MESG_RECB_INTE      = '" & CLng(pxmlTipoOperacao.documentElement.selectSingleNode("TP_MESG_RECB_INTE").Text) & "'," & _
            "         TP_MESG_RETN_INTE      =  " & CLng(pxmlTipoOperacao.documentElement.selectSingleNode("TP_MESG_RETN_INTE").Text) & "," & _
            "         CO_MESG_SPB_REGT_OPER  = '" & pxmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text & "'," & _
            "         DT_INIC_VIGE           =  " & fgDtXML_To_Oracle(pxmlTipoOperacao.documentElement.selectSingleNode("DT_INIC_VIGE").Text) & "," & _
            "         DT_FIM_VIGE            =  " & fgDtXML_To_Oracle(pxmlTipoOperacao.documentElement.selectSingleNode("DT_FIM_VIGE").Text) & "," & _
            "         CO_USUA_ULTI_ATLZ      = '" & fgObterUsuarioRede & "'," & _
            "         CO_ETCA_TRAB_ULTI_ATLZ = '" & fgObterEstacaoTrabalhoUsuario & "'," & _
            "         DH_ULTI_ATLZ           =      SYSDATE " & _
            " WHERE   TP_OPER                =  " & CLng(pxmlTipoOperacao.documentElement.selectSingleNode("TP_OPER").Text) & _
            " AND     DH_ULTI_ATLZ           =  " & fgDtHrXML_To_Oracle(pxmlTipoOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)
    
    If fgExecuteSQL(strSQL) = 0 Then
        'Registro foi alterado por outro usuário
        lngCodigoErroNegocio = 900
        GoTo ErrorHandler
    End If
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuparar um registro na tabela de tipo de operação para um tipo de mensagem de recebimento (Tipo de mensagem - A7)

Public Function LerPorTipoMensagemRecebimento(ByVal plngTipoMensagem As Long) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   TP_OPER, " & _
            "           NO_TIPO_OPER, " & _
            "           TP_MESG_RECB_INTE, " & _
            "           TP_MESG_RETN_INTE, " & _
            "           CO_MESG_SPB_REGT_OPER, " & _
            "           DT_INIC_VIGE, " & _
            "           DT_FIM_VIGE, " & _
            "           CO_LOCA_LIQU, " & _
            "           CO_USUA_ULTI_ATLZ, " & _
            "           CO_ETCA_TRAB_ULTI_ATLZ, " & _
            "           DH_ULTI_ATLZ " & _
            " FROM      A8.TB_TIPO_OPER " & _
            " WHERE     TP_MESG_RECB_INTE = '" & plngTipoMensagem & "'"
    
    LerPorTipoMensagemRecebimento = fgQueryXMLLer("Grupo_TipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerPorTipoMensagemRecebimento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar um registro na tabela de tipo de operação

Public Function Ler(ByVal plngTipoOperacao As Long) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   TP_OPER, " & _
            "           NO_TIPO_OPER, " & _
            "           TP_MESG_RECB_INTE, " & _
            "           TP_MESG_RETN_INTE, " & _
            "           CO_MESG_SPB_REGT_OPER, " & _
            "           DT_INIC_VIGE, " & _
            "           DT_FIM_VIGE, " & _
            "           CO_LOCA_LIQU, " & _
            "           CO_USUA_ULTI_ATLZ, " & _
            "           CO_ETCA_TRAB_ULTI_ATLZ, " & _
            "           DH_ULTI_ATLZ " & _
            " FROM      A8.TB_TIPO_OPER " & _
            " WHERE     TP_OPER = " & plngTipoOperacao
    
    Ler = fgQueryXMLLer("Grupo_TipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Gerenciar as operações de Inclusão,Alteração e Exclusão

Public Function Salvar(ByVal pxmlTipoOperacao As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    flRegrasNegocio pxmlTipoOperacao
            
    If pxmlTipoOperacao.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pxmlTipoOperacao
    ElseIf pxmlTipoOperacao.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        flAlterar pxmlTipoOperacao
    ElseIf pxmlTipoOperacao.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        flExcluir pxmlTipoOperacao
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

'regra de negócio

Private Function flRegrasNegocio(ByRef pxmlTipoOperacao As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler

    If pxmlTipoOperacao.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        
        If flVerificaTipoOperacao(CLng(pxmlTipoOperacao.documentElement.selectSingleNode("TP_OPER").Text)) Then
            'Atributo já está cadastrado
            lngCodigoErroNegocio = 501
            GoTo ErrorHandler
        End If
        
    ElseIf pxmlTipoOperacao.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        
    End If

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flRegrasNegocio Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verifica se existe um tipo de operação

Private Function flVerificaTipoOperacao(ByVal plngTipoOperacao As Long) As Boolean

Dim strSQL                                   As String
Dim objRS                                    As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT  TP_OPER " & _
             "   FROM  A8.TB_TIPO_OPER " & _
             "  WHERE  TP_OPER = " & plngTipoOperacao
    
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.EOF Then
        flVerificaTipoOperacao = False
    Else
        flVerificaTipoOperacao = True
    End If
    
    objRS.Close

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificaTipoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Excluir um registro na tabela de tipo de operação

Private Function flExcluir(ByRef pxmlTipoOperacao As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " DELETE FROM A8.TB_TIPO_OPER " & _
             "  WHERE   TP_OPER            = " & CLng(pxmlTipoOperacao.documentElement.selectSingleNode("TP_OPER").Text)

    If fgExecuteSQL(strSQL) = 0 Then
        'Registro foi exccluido por outro usuário
        lngCodigoErroNegocio = 901
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o tipo de operação pela remessa
Public Function ObterTipoOperacaoPorMsgBUS(ByRef pxmlDomRemessa As MSXML2.DOMDocument40, _
                                           ByRef plngTipoOperacao As Long, _
                                           ByRef pstrCodigoMensagemSPB As String, _
                                           ByRef pstrMensagemRetornoLegado As String, _
                                  Optional ByVal pintAcaoControleInterno As Integer = 0) As Boolean

Dim strSQL                                  As String
Dim strXML                                  As String
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode

Dim intLayoutEntrada                        As Integer

    On Error GoTo ErrorHandler
    
    intLayoutEntrada = 0
    If Not pxmlDomRemessa.documentElement.selectSingleNode("TP_MESG") Is Nothing Then
        intLayoutEntrada = Val(pxmlDomRemessa.documentElement.selectSingleNode("TP_MESG").Text)
    End If
    
    Select Case intLayoutEntrada
        Case enumTipoMensagemBUS.EnvioPagDespesas
        
            Call ObterTipoOperacaoPagDespesas(pxmlDomRemessa, plngTipoOperacao, pstrCodigoMensagemSPB, pstrMensagemRetornoLegado)
            Exit Function
            
        Case enumTipoMensagemBUS.OperacoesCorretoras
        
            Call ObterTipoOperacaoCorretora(pxmlDomRemessa, plngTipoOperacao, pstrCodigoMensagemSPB, pstrMensagemRetornoLegado)
            Exit Function

        Case enumTipoMensagemBUS.MovimentacaoInstrumentoFinanceiro
        
            Call ObterTipoOperacaoMovInstrFinanceiro(pxmlDomRemessa, plngTipoOperacao, pstrCodigoMensagemSPB, pstrMensagemRetornoLegado)
            Exit Function
    
        Case enumTipoMensagemBUS.AntecipacaoResgateContratoDerivativoCETIP, _
             enumTipoMensagemBUS.LancamentoPU_CETIP, _
             enumTipoMensagemBUS.EnvioTEDClientes

            Call ObterTipoOperacaoNovoGenerico(pxmlDomRemessa, plngTipoOperacao, pstrCodigoMensagemSPB, pstrMensagemRetornoLegado)
            Exit Function
            
        Case enumTipoMensagemBUS.RegistroContratoSWAP_CETIP21
             
            Call ObterTipoOperacaoRegraFixa(intLayoutEntrada, pxmlDomRemessa, plngTipoOperacao, pstrCodigoMensagemSPB, pstrMensagemRetornoLegado)
            Exit Function
            
        Case enumTipoMensagemBUS.RegistroOperacaoInterbancaria, _
             enumTipoMensagemBUS.RegistroOperacaoArbitragem
            
            Call ObterTipoOperacaoInteligente(pxmlDomRemessa, plngTipoOperacao, pstrCodigoMensagemSPB, pstrMensagemRetornoLegado)
            Exit Function
    End Select

    If pxmlDomRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMA Or _
       pxmlDomRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMD Or _
       pxmlDomRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMC Then
        
        ObterTipoOperacaoBMAPorMsgBUS pxmlDomRemessa, _
                                      plngTipoOperacao, _
                                      pstrCodigoMensagemSPB, _
                                      pstrMensagemRetornoLegado, _
                                      pintAcaoControleInterno
        Exit Function
    
    End If

    strSQL = " Select   A.TP_OPER, " & vbCrLf & _
             "          A.CO_MESG_SPB_REGT_OPER, " & vbCrLf & _
             "          A.TP_MESG_RETN_INTE, " & vbCrLf & _
             "          B.NO_ATRB_MESG, " & vbCrLf & _
             "          B.DE_CNTD_ATRB " & vbCrLf & _
             " From     A8.TB_TIPO_OPER             A, " & vbCrLf & _
             "          A8.TB_TIPO_OPER_CNTD_ATRB   B  " & vbCrLf & _
             " Where    A.TP_OPER            =  B.TP_OPER(+) " & vbCrLf & _
             " And      A.TP_MESG_RECB_INTE  = '" & intLayoutEntrada & "'" & vbCrLf & _
             " And      A.DT_INIC_VIGE      <=  " & fgDataHoraServidor_To_Oracle & vbCrLf & _
             " And     (A.DT_FIM_VIGE       IS  NULL " & vbCrLf & _
             " Or       A.DT_FIM_VIGE       >=  " & fgDataHoraServidor_To_Oracle & ")" & vbCrLf & _
             " Order By A.TP_OPER "

    strXML = fgQueryXMLLerTodos("TipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Trim(strXML) = vbNullString Then
        'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
        lngCodigoErroNegocio = 3003
        GoTo ErrorHandler
    End If

    plngTipoOperacao = 0

    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlTipoOperacao.loadXML strXML
    For Each objNode In xmlTipoOperacao.documentElement.selectNodes("//Grupo_TipoOperacao")
        If objNode.selectSingleNode("DE_CNTD_ATRB").Text <> vbNullString Then
            If objNode.selectSingleNode("DE_CNTD_ATRB").Text = _
                pxmlDomRemessa.documentElement.selectSingleNode(objNode.selectSingleNode("NO_ATRB_MESG").Text).Text Then
                plngTipoOperacao = CLng(objNode.selectSingleNode("TP_OPER").Text)
                pstrCodigoMensagemSPB = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
                Exit For
            End If
        Else
            'Regra específica para Operação Compromissada Retorno
            If objNode.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaVolta Or _
                objNode.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaVoltaConciliacao Then
                If pxmlDomRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = _
                    pxmlDomRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text Then
                    plngTipoOperacao = enumTipoOperacaoLQS.CompromissadaVoltaConciliacao
                    pstrCodigoMensagemSPB = ""
                    pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
                    Exit For
                Else
                    plngTipoOperacao = enumTipoOperacaoLQS.CompromissadaVolta
                    pstrCodigoMensagemSPB = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                    pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
                    Exit For
                End If
            Else
                plngTipoOperacao = CLng(objNode.selectSingleNode("TP_OPER").Text)
                pstrCodigoMensagemSPB = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
                Exit For
            End If
        End If
    Next

    If plngTipoOperacao = 0 Then
        'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
        lngCodigoErroNegocio = 3003
        GoTo ErrorHandler
    End If

VerificaConciliacao:
    If pxmlDomRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
        If intLayoutEntrada <> enumTipoMensagemBUS.OperacoesCorretoras Then
            flVerificarConciliacaoCETIP pxmlDomRemessa, plngTipoOperacao
        End If
    End If

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoOperacaoPorMsgBUS Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o tipo de operação pela remessa
Public Function ObterTipoOperacaoBMAPorMsgBUS(ByRef pxmlDomRemessa As MSXML2.DOMDocument40, _
                                              ByRef plngTipoOperacao As Long, _
                                              ByRef pstrCodigoMensagemSPB As String, _
                                              ByRef pstrMensagemRetornoLegado As String, _
                                     Optional ByVal pintAcaoControleInterno As Integer = 0) As Boolean

Dim strSQL                                  As String
Dim strXML                                  As String
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode
Dim objNodeAux                              As MSXML2.IXMLDOMNode
Dim strNomeTagAnt                           As String
Dim blnSegundaTag                           As Boolean
Dim lngTipoOperacaoAnt                      As Long

On Error GoTo ErrorHandler

    strSQL = " Select   A.TP_OPER, " & vbCrLf & _
             "          A.CO_MESG_SPB_REGT_OPER, " & vbCrLf & _
             "          A.TP_MESG_RETN_INTE, " & vbCrLf & _
             "          B.NO_ATRB_MESG, " & vbCrLf & _
             "          B.DE_CNTD_ATRB " & vbCrLf & _
             " From     A8.TB_TIPO_OPER             A, " & vbCrLf & _
             "          A8.TB_TIPO_OPER_CNTD_ATRB   B  " & vbCrLf & _
             " Where    A.TP_OPER            =  B.TP_OPER(+) " & vbCrLf & _
             " And      A.TP_MESG_RECB_INTE  =  '" & CLng(pxmlDomRemessa.documentElement.selectSingleNode("TP_MESG").Text) & "'" & vbCrLf & _
             " And      A.DT_INIC_VIGE      <=  " & fgDataHoraServidor_To_Oracle & vbCrLf & _
             " And     (A.DT_FIM_VIGE       IS  NULL " & vbCrLf & _
             " Or       A.DT_FIM_VIGE       >=  " & fgDataHoraServidor_To_Oracle & ")" & vbCrLf & _
             " Order By A.TP_OPER "

    strXML = fgQueryXMLLerTodos("TipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Trim(strXML) = vbNullString Then
        If pintAcaoControleInterno = enumAcaoControleInterno.ErrorHandlerBypass Then
            Exit Function
        End If
        'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
        lngCodigoErroNegocio = 3003
        GoTo ErrorHandler
    End If

    plngTipoOperacao = 0

    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlTipoOperacao.loadXML strXML
    blnSegundaTag = False
    For Each objNode In xmlTipoOperacao.documentElement.selectNodes("//Grupo_TipoOperacao")
        If objNode.selectSingleNode("DE_CNTD_ATRB").Text <> vbNullString Then
            If objNode.selectSingleNode("DE_CNTD_ATRB").Text = _
                pxmlDomRemessa.documentElement.selectSingleNode(objNode.selectSingleNode("NO_ATRB_MESG").Text).Text Then
                strNomeTagAnt = objNode.selectSingleNode("NO_ATRB_MESG").Text
                lngTipoOperacaoAnt = CLng(objNode.selectSingleNode("TP_OPER").Text)
                For Each objNodeAux In xmlTipoOperacao.documentElement.selectNodes("//Grupo_TipoOperacao")
                    If strNomeTagAnt <> objNodeAux.selectSingleNode("NO_ATRB_MESG").Text Then
                        blnSegundaTag = True
                        If objNodeAux.selectSingleNode("DE_CNTD_ATRB").Text = _
                            pxmlDomRemessa.documentElement.selectSingleNode(objNodeAux.selectSingleNode("NO_ATRB_MESG").Text).Text And _
                            lngTipoOperacaoAnt = CLng(objNodeAux.selectSingleNode("TP_OPER").Text) Then
                            plngTipoOperacao = CLng(objNodeAux.selectSingleNode("TP_OPER").Text)
                            pstrCodigoMensagemSPB = objNodeAux.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                            pstrMensagemRetornoLegado = objNodeAux.selectSingleNode("TP_MESG_RETN_INTE").Text
                            Exit For
                        End If
                    End If
                Next
                If blnSegundaTag = False Then
                    plngTipoOperacao = CLng(objNode.selectSingleNode("TP_OPER").Text)
                    pstrCodigoMensagemSPB = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                    pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
                    
                    If plngTipoOperacao = enumTipoOperacaoLQS.TransferenciaBMA Then
                        
                        If fgCompletaString(Right$(pxmlDomRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text, 9), "0", 9, True) <> _
                           fgCompletaString(Right$(pxmlDomRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU_DEST").Text, 9), "0", 9, True) Then
                            
                            If fgVerificaContaMesmaCamara(pxmlDomRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text, _
                                                          pxmlDomRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU_DEST").Text) Then
                                
                                plngTipoOperacao = enumTipoOperacaoLQS.TransferenciaBMA
                            Else
                                plngTipoOperacao = enumTipoOperacaoLQS.MovimentacaoEntreCamarasBMA
                            End If
                            
                        End If
                    
                    ElseIf plngTipoOperacao = enumTipoOperacaoLQS.MovimentacaoEntreCamarasBMA Then
                        
                        If fgCompletaString(Right$(pxmlDomRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text, 9), "0", 9, True) = _
                           fgCompletaString(Right$(pxmlDomRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU_DEST").Text, 9), "0", 9, True) Then
                            plngTipoOperacao = enumTipoOperacaoLQS.TransferenciaBMA
                        End If
                        
                    End If
                
                End If
                If plngTipoOperacao > 0 Then
                    Exit For
                End If
            End If
        Else
            
            plngTipoOperacao = CLng(objNode.selectSingleNode("TP_OPER").Text)
            pstrCodigoMensagemSPB = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
            pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
            
            Exit For
        End If
    Next

    'Esta fixo pois entrou depois no layout 32, como tem outras operações que utiliza o mesmo e todas é verificado
    'por duas tag para achar a operação não alterei a rotina no momento, senão teria que retestar muita coisa e não tem o tempo necessario
    Select Case plngTipoOperacao
        Case enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA, enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA
            If pxmlDomRemessa.documentElement.selectSingleNode("CO_SUB_TIPO_NEGO").Text = "2" Then
                plngTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaTermo
                pstrCodigoMensagemSPB = vbNullString
            ElseIf pxmlDomRemessa.documentElement.selectSingleNode("CO_SUB_TIPO_NEGO").Text = "3" Then
                plngTipoOperacao = enumTipoOperacaoLQS.CompromissadaGenericaAVista
                pstrCodigoMensagemSPB = vbNullString
            ElseIf pxmlDomRemessa.documentElement.selectSingleNode("CO_SUB_TIPO_NEGO").Text = "4" Then
                plngTipoOperacao = enumTipoOperacaoLQS.CompromissadaGenericaATermo
                pstrCodigoMensagemSPB = vbNullString
            End If
    End Select

    If plngTipoOperacao = 0 Then
        If pintAcaoControleInterno = enumAcaoControleInterno.ErrorHandlerBypass Then
            Exit Function
        End If
        'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
        lngCodigoErroNegocio = 3003
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoOperacaoBMAPorMsgBUS Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o tipo de operação pela remessa
Public Function ObterTipoOperacaoPagDespesas(ByRef pxmlDomRemessa As MSXML2.DOMDocument40, _
                                           ByRef plngTipoOperacao As Long, _
                                           ByRef pstrCodigoMensagemSPB As String, _
                                           ByRef pstrMensagemRetornoLegado As String) As Boolean

Dim strSQL                                  As String
Dim strXML                                  As String
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode
Dim objNodeAux                              As MSXML2.IXMLDOMNode
Dim objNodeAux2                             As MSXML2.IXMLDOMNode
Dim blnSegundaTag                           As Boolean
Dim blnTerceiraTag                          As Boolean
Dim lngTipoOperacaoAnt                      As Long
Dim strNomeTagAnt                           As String
Dim lngTipoOperacaoAnt2                     As Long
Dim strNomeTagAnt2                          As String
Dim strNomeTagAnt3                          As String
Dim lngTipoOperacaoAnt3                     As Long
'

On Error GoTo ErrorHandler

    strSQL = " Select   A.TP_OPER, " & vbCrLf & _
             "          A.CO_MESG_SPB_REGT_OPER, " & vbCrLf & _
             "          A.TP_MESG_RETN_INTE, " & vbCrLf & _
             "          B.NO_ATRB_MESG, " & vbCrLf & _
             "          B.DE_CNTD_ATRB " & vbCrLf & _
             " From     A8.TB_TIPO_OPER             A, " & vbCrLf & _
             "          A8.TB_TIPO_OPER_CNTD_ATRB   B  " & vbCrLf & _
             " Where    A.TP_OPER            =  B.TP_OPER(+) " & vbCrLf & _
             " And      A.TP_MESG_RECB_INTE  =  '" & CLng(pxmlDomRemessa.documentElement.selectSingleNode("TP_MESG").Text) & "'" & vbCrLf & _
             " And      A.DT_INIC_VIGE      <=  " & fgDataHoraServidor_To_Oracle & vbCrLf & _
             " And     (A.DT_FIM_VIGE       IS  NULL " & vbCrLf & _
             " Or       A.DT_FIM_VIGE       >=  " & fgDataHoraServidor_To_Oracle & ")" & vbCrLf & _
             " Order By A.TP_OPER"

    strXML = fgQueryXMLLerTodos("TipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Trim(strXML) = vbNullString Then
        'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
        lngCodigoErroNegocio = 3003
        GoTo ErrorHandler
    End If

    plngTipoOperacao = 0

    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlTipoOperacao.loadXML strXML
    blnSegundaTag = False
    strNomeTagAnt = ""
    strNomeTagAnt2 = ""
    For Each objNode In xmlTipoOperacao.documentElement.selectNodes("//Grupo_TipoOperacao")
        If objNode.selectSingleNode("DE_CNTD_ATRB").Text <> vbNullString Then
            If objNode.selectSingleNode("DE_CNTD_ATRB").Text = _
                pxmlDomRemessa.documentElement.selectSingleNode(objNode.selectSingleNode("NO_ATRB_MESG").Text).Text Then

                strNomeTagAnt = objNode.selectSingleNode("NO_ATRB_MESG").Text
                lngTipoOperacaoAnt = CLng(objNode.selectSingleNode("TP_OPER").Text)

                'If strNomeTagAnt = "CO_FORM_LIQU" And CLng(objNode.selectSingleNode("DE_CNTD_ATRB").Text) <> 4 Then
                '    plngTipoOperacao = CLng(objNode.selectSingleNode("TP_OPER").Text)
                '    pstrCodigoMensagemSPB = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                '    pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
                '    Exit For
                'End If

                For Each objNodeAux In xmlTipoOperacao.documentElement.selectNodes("//Grupo_TipoOperacao")
                    If strNomeTagAnt <> objNodeAux.selectSingleNode("NO_ATRB_MESG").Text Then
                        blnSegundaTag = True

                        strNomeTagAnt2 = objNodeAux.selectSingleNode("NO_ATRB_MESG").Text
                        lngTipoOperacaoAnt2 = CLng(objNodeAux.selectSingleNode("TP_OPER").Text)

                        If objNodeAux.selectSingleNode("DE_CNTD_ATRB").Text = _
                            pxmlDomRemessa.documentElement.selectSingleNode(objNodeAux.selectSingleNode("NO_ATRB_MESG").Text).Text And _
                            lngTipoOperacaoAnt = CLng(objNodeAux.selectSingleNode("TP_OPER").Text) Then


                            For Each objNodeAux2 In xmlTipoOperacao.documentElement.selectNodes("//Grupo_TipoOperacao")
                                                                   
                                If strNomeTagAnt <> objNodeAux2.selectSingleNode("NO_ATRB_MESG").Text And _
                                    strNomeTagAnt2 = objNodeAux2.selectSingleNode("NO_ATRB_MESG").Text And CLng(objNodeAux2.selectSingleNode("TP_OPER").Text) = lngTipoOperacaoAnt Then
                                    plngTipoOperacao = CLng(objNodeAux2.selectSingleNode("TP_OPER").Text)
                                    pstrCodigoMensagemSPB = objNodeAux2.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                                    pstrMensagemRetornoLegado = objNodeAux2.selectSingleNode("TP_MESG_RETN_INTE").Text
                                    Exit For
                                End If
                            Next

                        End If
                    End If
                Next

'                If blnSegundaTag = False Then
'                    plngTipoOperacao = CLng(objNode.selectSingleNode("TP_OPER").Text)
'                    pstrCodigoMensagemSPB = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
'                    pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
'                End If

                If plngTipoOperacao > 0 Then
                    Exit For
                End If
            End If
        Else
            
            plngTipoOperacao = CLng(objNode.selectSingleNode("TP_OPER").Text)
            pstrCodigoMensagemSPB = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
            pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
            
            Exit For
        End If
    Next

    If plngTipoOperacao = 0 Then
        'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
        lngCodigoErroNegocio = 3003
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoOperacaoPagDespesas Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o tipo de operação pela remessa
Public Function ObterTipoOperacaoNovoGenerico(ByRef pxmlDomRemessa As MSXML2.DOMDocument40, _
                                              ByRef plngTipoOperacao As Long, _
                                              ByRef pstrCodigoMensagemSPB As String, _
                                              ByRef pstrMensagemRetornoLegado As String) As Boolean

Dim strSQL                                  As String
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode

Dim strCampos                               As String
Dim arrCampos()                             As String
Dim strConteudos                            As String
Dim arrConteudos()                          As String

Dim intAtributos                            As Integer
Dim intTipoOperacaoAnt                      As Integer
Dim strCodigoMsgAnt                         As String
Dim strRetornoLegadoAnt                     As String
Dim strAtributoAnt                          As String

Dim strWhere                                As String

    On Error GoTo ErrorHandler

    strSQL = " Select   A.TP_OPER, " & vbCrLf & _
             "          A.CO_MESG_SPB_REGT_OPER, " & vbCrLf & _
             "          A.TP_MESG_RETN_INTE, " & vbCrLf & _
             "          B.NO_ATRB_MESG, " & vbCrLf & _
             "          B.DE_CNTD_ATRB " & vbCrLf & _
             " From     A8.TB_TIPO_OPER             A, " & vbCrLf & _
             "          A8.TB_TIPO_OPER_CNTD_ATRB   B  " & vbCrLf & _
             " Where    A.TP_OPER            =  B.TP_OPER(+) " & vbCrLf & _
             " And      A.TP_MESG_RECB_INTE  =  '" & CLng(pxmlDomRemessa.documentElement.selectSingleNode("TP_MESG").Text) & "'" & vbCrLf & _
             " And      A.DT_INIC_VIGE      <=  " & fgDataHoraServidor_To_Oracle & vbCrLf & _
             " And     (A.DT_FIM_VIGE       IS  NULL " & vbCrLf & _
             " Or       A.DT_FIM_VIGE       >=  " & fgDataHoraServidor_To_Oracle & ")" & vbCrLf & _
             " Order By A.TP_OPER"

    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlTipoOperacao.loadXML(fgQueryXMLLerTodos("TipoOperacao", strSQL, App.EXEName & "." & TypeName(Me)))

    If xmlTipoOperacao.xml = vbNullString Then
        'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
        lngCodigoErroNegocio = 3003
        GoTo ErrorHandler
    End If

    plngTipoOperacao = 0
    intTipoOperacaoAnt = 0
    strCampos = vbNullString
    
    For Each objNode In xmlTipoOperacao.selectNodes("//Grupo_TipoOperacao")
        
        If Val(objNode.selectSingleNode("TP_OPER").Text) = intTipoOperacaoAnt Or intTipoOperacaoAnt = 0 Then
            
            intTipoOperacaoAnt = Val(objNode.selectSingleNode("TP_OPER").Text)
            strCodigoMsgAnt = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
            strRetornoLegadoAnt = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
            
            strCampos = strCampos & "|" & objNode.selectSingleNode("NO_ATRB_MESG").Text
            strConteudos = strConteudos & "|" & objNode.selectSingleNode("DE_CNTD_ATRB").Text
            
        Else
        
            strWhere = vbNullString
            strAtributoAnt = vbNullString
            
            arrCampos = Split(Mid$(strCampos, 2), "|")
            arrConteudos = Split(Mid$(strConteudos, 2), "|")
            
            For intAtributos = LBound(arrCampos) To UBound(arrCampos)
                strWhere = strWhere & arrCampos(intAtributos) & "='" & arrConteudos(intAtributos) & "' and "
                
                If intAtributos > 0 And strAtributoAnt = arrCampos(intAtributos) Then
                    strWhere = Replace$(strWhere, " and ", " or  ")
                End If
                
                strAtributoAnt = arrCampos(intAtributos)
            Next
            
            If strWhere <> vbNullString Then
                strWhere = Left$(strWhere, Len(strWhere) - 5)
            End If
            'KIDA - 22/04/2009
            'RATS - 914
            'If Not pxmlDomRemessa.selectSingleNode("MESG[" & strWhere & "]") Is Nothing Then
            If Not pxmlDomRemessa.selectSingleNode("//*[" & strWhere & "]") Is Nothing Then
                plngTipoOperacao = intTipoOperacaoAnt
                pstrCodigoMensagemSPB = strCodigoMsgAnt
                pstrMensagemRetornoLegado = strRetornoLegadoAnt
                Exit Function
            
            Else
                intTipoOperacaoAnt = Val(objNode.selectSingleNode("TP_OPER").Text)
                strCodigoMsgAnt = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                strRetornoLegadoAnt = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
                
                strCampos = vbNullString
                strConteudos = vbNullString
                strCampos = strCampos & "|" & objNode.selectSingleNode("NO_ATRB_MESG").Text
                strConteudos = strConteudos & "|" & objNode.selectSingleNode("DE_CNTD_ATRB").Text
                
            End If
            
        End If
    
    Next
        
    strWhere = vbNullString
    arrCampos = Split(Mid$(strCampos, 2), "|")
    arrConteudos = Split(Mid$(strConteudos, 2), "|")
    
    For intAtributos = LBound(arrCampos) To UBound(arrCampos)
        strWhere = strWhere & "" & arrCampos(intAtributos) & "='" & arrConteudos(intAtributos) & "' and "
    Next
    
    If strWhere <> vbNullString Then
        strWhere = Left$(strWhere, Len(strWhere) - 5)
    End If
    
    'KIDA - 22/04/2009
    'RATS - 914
    'If Not pxmlDomRemessa.selectSingleNode("MESG[" & strWhere & "]") Is Nothing Then
    If Not pxmlDomRemessa.selectSingleNode("//*[" & strWhere & "]") Is Nothing Then
        plngTipoOperacao = intTipoOperacaoAnt
        pstrCodigoMensagemSPB = strCodigoMsgAnt
        pstrMensagemRetornoLegado = strRetornoLegadoAnt
    Else
        'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
        lngCodigoErroNegocio = 3003
        GoTo ErrorHandler
    End If
    
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoOperacaoNovoGenerico Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o tipo de operação pela remessa
Public Function ObterTipoOperacaoRegraFixa(ByVal pintLayOutEntrada As Integer, _
                                           ByRef pxmlDomRemessa As MSXML2.DOMDocument40, _
                                           ByRef plngTipoOperacao As Long, _
                                           ByRef pstrCodigoMensagemSPB As String, _
                                           ByRef pstrMensagemRetornoLegado As String) As Boolean

    On Error GoTo ErrorHandler

    Select Case pintLayOutEntrada
        Case enumTipoMensagemBUS.RegistroContratoSWAP_CETIP21
            
            plngTipoOperacao = enumTipoOperacaoLQS.RegistroContratoSWAPSemOpcaoBarreira
            pstrCodigoMensagemSPB = "CTP9001"
            pstrMensagemRetornoLegado = "95"
            
            If Not pxmlDomRemessa.selectSingleNode("//TP_CNTR_SWAP") Is Nothing Then
                If Trim$(pxmlDomRemessa.selectSingleNode("//TP_CNTR_SWAP").Text) <> vbNullString Then
                    plngTipoOperacao = enumTipoOperacaoLQS.RegistroContratoSWAPComOpcaoBarreira
                End If
            End If
            
    End Select
        
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoOperacaoNovoGenerico Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o Tipo de Operacao pela Remessa
Public Function ObterTipoOperacaoInteligente(ByRef pxmlDomRemessa As MSXML2.DOMDocument40, _
                                             ByRef plngTipoOperacao As Long, _
                                             ByRef pstrCodigoMensagemSPB As String, _
                                             ByRef pstrMensagemRetornoLegado As String) As Boolean

Dim strSQL                                  As String
Dim strXML                                  As String
Dim strXMLConteudo                          As String
Dim strXMLQuantidade                        As String
Dim strConteudoTag                          As String
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim xmlTipoConteudo                         As MSXML2.DOMDocument40
Dim xmlQuantidade                           As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode
Dim objNodeConteudo                         As MSXML2.IXMLDOMNode
Dim lngQuantidade                           As Long
Dim lngQuantidadeVerdadeira                 As Long

On Error GoTo ErrorHandler

    strSQL = " Select   A.TP_OPER, " & vbCrLf & _
             "          A.CO_MESG_SPB_REGT_OPER, " & vbCrLf & _
             "          A.TP_MESG_RETN_INTE " & vbCrLf & _
             " From     A8.TB_TIPO_OPER             A " & vbCrLf & _
             " Where    A.TP_MESG_RECB_INTE  =  '" & CLng(pxmlDomRemessa.documentElement.selectSingleNode("TP_MESG").Text) & "'" & vbCrLf & _
             " And      A.DT_INIC_VIGE      <=  " & fgDataHoraServidor_To_Oracle & vbCrLf & _
             " And     (A.DT_FIM_VIGE       IS  NULL " & vbCrLf & _
             " Or       A.DT_FIM_VIGE       >=  " & fgDataHoraServidor_To_Oracle & ")" & vbCrLf & _
             " Order By A.TP_OPER"

    strXML = fgQueryXMLLerTodos("TipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Trim(strXML) = vbNullString Then
        'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
        lngCodigoErroNegocio = 3003
        GoTo ErrorHandler
    End If

    plngTipoOperacao = 0

    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlTipoOperacao.loadXML strXML
    
    If xmlTipoOperacao.documentElement.selectNodes("//Grupo_TipoOperacao").Length = 1 Then
        plngTipoOperacao = CLng(xmlTipoOperacao.documentElement.selectSingleNode("//Grupo_TipoOperacao/TP_OPER").Text)
    Else
        For Each objNode In xmlTipoOperacao.documentElement.selectNodes("//Grupo_TipoOperacao")
            strSQL = " Select   B.NO_ATRB_MESG " & vbCrLf & _
                     " From     A8.TB_TIPO_OPER_CNTD_ATRB   B  " & vbCrLf & _
                     " Where    B.TP_OPER            =  " & CLng(objNode.selectSingleNode("TP_OPER").Text) & vbCrLf & _
                     " Group By B.NO_ATRB_MESG " & vbCrLf & _
                     " Order By B.NO_ATRB_MESG"

            strXMLQuantidade = fgQueryXMLLerTodos("Quantidade", strSQL, App.EXEName & "." & TypeName(Me))
            
            Set xmlQuantidade = CreateObject("MSXML2.DOMDocument.4.0")
            xmlQuantidade.loadXML strXMLQuantidade
            
            lngQuantidade = xmlQuantidade.documentElement.selectNodes("//Grupo_Quantidade").Length
        
            strSQL = " Select   B.NO_ATRB_MESG, " & vbCrLf & _
                     "          B.DE_CNTD_ATRB " & vbCrLf & _
                     " From     A8.TB_TIPO_OPER_CNTD_ATRB   B  " & vbCrLf & _
                     " Where    B.TP_OPER            =  " & CLng(objNode.selectSingleNode("TP_OPER").Text) & vbCrLf & _
                     " Order By B.NO_ATRB_MESG"

            strXMLConteudo = fgQueryXMLLerTodos("TipoConteudo", strSQL, App.EXEName & "." & TypeName(Me))
            
            Set xmlTipoConteudo = CreateObject("MSXML2.DOMDocument.4.0")
            xmlTipoConteudo.loadXML strXMLConteudo
            
            lngQuantidadeVerdadeira = 0
            
            For Each objNodeConteudo In xmlTipoConteudo.documentElement.selectNodes("//Grupo_TipoConteudo")
                If Not pxmlDomRemessa.documentElement.selectSingleNode(objNodeConteudo.selectSingleNode("NO_ATRB_MESG").Text) Is Nothing Then
                    strConteudoTag = pxmlDomRemessa.documentElement.selectSingleNode(objNodeConteudo.selectSingleNode("NO_ATRB_MESG").Text).Text
                    If strConteudoTag = objNodeConteudo.selectSingleNode("DE_CNTD_ATRB").Text Then
                        lngQuantidadeVerdadeira = lngQuantidadeVerdadeira + 1
                    End If
                End If
            Next
            
            If lngQuantidade = lngQuantidadeVerdadeira Then
                plngTipoOperacao = CLng(objNode.selectSingleNode("TP_OPER").Text)
                pstrCodigoMensagemSPB = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
                Exit For
            End If
        Next
    End If

    If plngTipoOperacao = 0 Then
        'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
        lngCodigoErroNegocio = 3003
        GoTo ErrorHandler
    End If
        
    Exit Function
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoOperacaoNovoGenerico Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o tipo de operação pela remessa
Public Function ObterTipoOperacaoCorretora(ByRef pxmlDomRemessa As MSXML2.DOMDocument40, _
                                           ByRef plngTipoOperacao As Long, _
                                           ByRef pstrCodigoMensagemSPB As String, _
                                           ByRef pstrMensagemRetornoLegado As String) As Boolean

Dim strSQL                                  As String
Dim strXML                                  As String
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode
Dim objNodeAux                              As MSXML2.IXMLDOMNode
Dim objNodeAux2                             As MSXML2.IXMLDOMNode
Dim blnSegundaTag                           As Boolean
Dim blnTerceiraTag                          As Boolean
Dim lngTipoOperacaoAnt                      As Long
Dim strNomeTagAnt                           As String
Dim lngTipoOperacaoAnt2                     As Long
Dim strNomeTagAnt2                          As String

On Error GoTo ErrorHandler

    strSQL = " Select   A.TP_OPER, " & vbCrLf & _
             "          A.CO_MESG_SPB_REGT_OPER, " & vbCrLf & _
             "          A.TP_MESG_RETN_INTE, " & vbCrLf & _
             "          B.NO_ATRB_MESG, " & vbCrLf & _
             "          B.DE_CNTD_ATRB " & vbCrLf & _
             " From     A8.TB_TIPO_OPER             A, " & vbCrLf & _
             "          A8.TB_TIPO_OPER_CNTD_ATRB   B  " & vbCrLf & _
             " Where    A.TP_OPER            =  B.TP_OPER(+) " & vbCrLf & _
             " And      A.TP_MESG_RECB_INTE  =  '" & CLng(pxmlDomRemessa.documentElement.selectSingleNode("TP_MESG").Text) & "'" & vbCrLf & _
             " And      A.DT_INIC_VIGE      <=  " & fgDataHoraServidor_To_Oracle & vbCrLf & _
             " And     (A.DT_FIM_VIGE       IS  NULL " & vbCrLf & _
             " Or       A.DT_FIM_VIGE       >=  " & fgDataHoraServidor_To_Oracle & ")" & vbCrLf & _
             " Order By A.TP_OPER"

    strXML = fgQueryXMLLerTodos("TipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Trim(strXML) = vbNullString Then
        'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
        lngCodigoErroNegocio = 3003
        GoTo ErrorHandler
    End If

    plngTipoOperacao = 0

    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlTipoOperacao.loadXML strXML
    blnSegundaTag = False
    For Each objNode In xmlTipoOperacao.documentElement.selectNodes("//Grupo_TipoOperacao")
        If objNode.selectSingleNode("DE_CNTD_ATRB").Text <> vbNullString Then
            If objNode.selectSingleNode("DE_CNTD_ATRB").Text = _
                pxmlDomRemessa.documentElement.selectSingleNode(objNode.selectSingleNode("NO_ATRB_MESG").Text).Text Then

                strNomeTagAnt = objNode.selectSingleNode("NO_ATRB_MESG").Text
                lngTipoOperacaoAnt = CLng(objNode.selectSingleNode("TP_OPER").Text)

                For Each objNodeAux In xmlTipoOperacao.documentElement.selectNodes("//Grupo_TipoOperacao")
                    If strNomeTagAnt <> objNodeAux.selectSingleNode("NO_ATRB_MESG").Text Then
                        blnSegundaTag = True

                        strNomeTagAnt2 = objNodeAux.selectSingleNode("NO_ATRB_MESG").Text
                        lngTipoOperacaoAnt2 = CLng(objNodeAux.selectSingleNode("TP_OPER").Text)

                        If objNodeAux.selectSingleNode("DE_CNTD_ATRB").Text = _
                            pxmlDomRemessa.documentElement.selectSingleNode(objNodeAux.selectSingleNode("NO_ATRB_MESG").Text).Text And _
                            lngTipoOperacaoAnt = CLng(objNodeAux.selectSingleNode("TP_OPER").Text) Then


                            For Each objNodeAux2 In xmlTipoOperacao.documentElement.selectNodes("//Grupo_TipoOperacao")
                                If strNomeTagAnt <> objNodeAux2.selectSingleNode("NO_ATRB_MESG").Text And _
                                    strNomeTagAnt2 <> objNodeAux2.selectSingleNode("NO_ATRB_MESG").Text Then

                                    blnTerceiraTag = True

                                    If objNodeAux2.selectSingleNode("DE_CNTD_ATRB").Text = _
                                        pxmlDomRemessa.documentElement.selectSingleNode(objNodeAux2.selectSingleNode("NO_ATRB_MESG").Text).Text And _
                                        lngTipoOperacaoAnt2 = CLng(objNodeAux2.selectSingleNode("TP_OPER").Text) Then

                                        plngTipoOperacao = CLng(objNodeAux2.selectSingleNode("TP_OPER").Text)
                                        pstrCodigoMensagemSPB = objNodeAux2.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                                        pstrMensagemRetornoLegado = objNodeAux2.selectSingleNode("TP_MESG_RETN_INTE").Text
                                        Exit For

                                    End If

                                End If
                            Next

                        End If
                    End If
                Next

'                If blnSegundaTag = False Then
'                    plngTipoOperacao = CLng(objNode.selectSingleNode("TP_OPER").Text)
'                    pstrCodigoMensagemSPB = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
'                    pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
'                End If

                If plngTipoOperacao > 0 Then
                    Exit For
                End If
            End If
        Else
            
            plngTipoOperacao = CLng(objNode.selectSingleNode("TP_OPER").Text)
            pstrCodigoMensagemSPB = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
            pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text
            
            Exit For
        End If
    Next

    If plngTipoOperacao = 0 Then
        'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
        lngCodigoErroNegocio = 3003
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoOperacaoCorretora Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o tipo de operação pela remessa
Public Function ObterTipoOperacaoMovInstrFinanceiro(ByRef pxmlDomRemessa As MSXML2.DOMDocument40, _
                                                    ByRef plngTipoOperacao As Long, _
                                                    ByRef pstrCodigoMensagemSPB As String, _
                                                    ByRef pstrMensagemRetornoLegado As String) As Boolean

Dim strSQL                                  As String
Dim strXML                                  As String
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode
Dim objNodeAux                              As MSXML2.IXMLDOMNode
Dim strNomeTagAnt                           As String
Dim blnSegundaTag                           As Boolean
Dim lngTipoOperacaoAnt                      As Long

On Error GoTo ErrorHandler

    strSQL = " Select   A.TP_OPER, " & vbCrLf & _
             "          A.CO_MESG_SPB_REGT_OPER, " & vbCrLf & _
             "          A.TP_MESG_RETN_INTE, " & vbCrLf & _
             "          B.NO_ATRB_MESG, " & vbCrLf & _
             "          B.DE_CNTD_ATRB " & vbCrLf & _
             " From     A8.TB_TIPO_OPER             A, " & vbCrLf & _
             "          A8.TB_TIPO_OPER_CNTD_ATRB   B  " & vbCrLf & _
             " Where    A.TP_OPER            =  B.TP_OPER(+) " & vbCrLf & _
             " And      A.TP_MESG_RECB_INTE  =  '" & CLng(pxmlDomRemessa.documentElement.selectSingleNode("TP_MESG").Text) & "'" & vbCrLf & _
             " And      A.DT_INIC_VIGE      <=  " & fgDataHoraServidor_To_Oracle & vbCrLf & _
             " And     (A.DT_FIM_VIGE       IS  NULL " & vbCrLf & _
             " Or       A.DT_FIM_VIGE       >=  " & fgDataHoraServidor_To_Oracle & ")" & vbCrLf & _
             " Order By A.TP_OPER "

    strXML = fgQueryXMLLerTodos("TipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Trim(strXML) = vbNullString Then
        'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
        lngCodigoErroNegocio = 3003
        GoTo ErrorHandler
    End If

    plngTipoOperacao = 0

    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlTipoOperacao.loadXML strXML
    blnSegundaTag = False

    For Each objNode In xmlTipoOperacao.documentElement.selectNodes("//Grupo_TipoOperacao")
        If objNode.selectSingleNode("DE_CNTD_ATRB").Text <> vbNullString Then
            If objNode.selectSingleNode("DE_CNTD_ATRB").Text = _
                pxmlDomRemessa.documentElement.selectSingleNode(objNode.selectSingleNode("NO_ATRB_MESG").Text).Text Then

                strNomeTagAnt = objNode.selectSingleNode("NO_ATRB_MESG").Text
                lngTipoOperacaoAnt = CLng(objNode.selectSingleNode("TP_OPER").Text)

                For Each objNodeAux In xmlTipoOperacao.documentElement.selectNodes("//Grupo_TipoOperacao")
                    If objNodeAux.selectSingleNode("DE_CNTD_ATRB").Text <> vbNullString Then
                        If strNomeTagAnt <> objNodeAux.selectSingleNode("NO_ATRB_MESG").Text Then
                            blnSegundaTag = True
                            If objNodeAux.selectSingleNode("DE_CNTD_ATRB").Text = _
                                pxmlDomRemessa.documentElement.selectSingleNode(objNodeAux.selectSingleNode("NO_ATRB_MESG").Text).Text And _
                                lngTipoOperacaoAnt = CLng(objNodeAux.selectSingleNode("TP_OPER").Text) Then
                                plngTipoOperacao = CLng(objNodeAux.selectSingleNode("TP_OPER").Text)
                                pstrCodigoMensagemSPB = objNodeAux.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                                pstrMensagemRetornoLegado = objNodeAux.selectSingleNode("TP_MESG_RETN_INTE").Text
                                Exit For
                            End If
                        End If
                    End If
                Next

                If blnSegundaTag = False Then

                    plngTipoOperacao = CLng(objNode.selectSingleNode("TP_OPER").Text)
                    pstrCodigoMensagemSPB = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                    pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text

                End If
                If plngTipoOperacao > 0 Then
                    Exit For
                End If
            End If
        Else

            If Trim(pxmlDomRemessa.selectSingleNode("//SG_SIST_CETIP").Text) <> "SCF" Then

                plngTipoOperacao = CLng(objNode.selectSingleNode("TP_OPER").Text)
                pstrCodigoMensagemSPB = objNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                pstrMensagemRetornoLegado = objNode.selectSingleNode("TP_MESG_RETN_INTE").Text

                Exit For
            End If
        End If
    Next

    If plngTipoOperacao = 0 Then
        
        flVerificarConciliacaoCETIP pxmlDomRemessa, plngTipoOperacao
        
        If plngTipoOperacao = 0 Then
            'Tipo de Operação para mensagem recebida do sistema BUS Inexistente
            lngCodigoErroNegocio = 3003
            GoTo ErrorHandler
        End If
    End If

    flVerificarConciliacaoCETIP pxmlDomRemessa, plngTipoOperacao, pstrCodigoMensagemSPB

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoOperacaoMovInstrFinanceiro Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o tipo de operação pela remessa
Public Function ObterTipoOperacaoEnvioTED(ByVal pintLayOutEntrada As Integer, _
                                          ByRef pxmlDomRemessa As MSXML2.DOMDocument40, _
                                          ByRef plngTipoOperacao As Long, _
                                          ByRef pstrCodigoMensagemSPB As String, _
                                          ByRef pstrMensagemRetornoLegado As String) As Boolean

    On Error GoTo ErrorHandler

    Select Case pintLayOutEntrada
        Case enumTipoMensagemBUS.RegistroContratoSWAP_CETIP21
            
            plngTipoOperacao = enumTipoOperacaoLQS.RegistroContratoSWAPSemOpcaoBarreira
            pstrCodigoMensagemSPB = "CTP9001"
            pstrMensagemRetornoLegado = "95"
            
            If Not pxmlDomRemessa.selectSingleNode("//TP_CNTR_SWAP") Is Nothing Then
                If Trim$(pxmlDomRemessa.selectSingleNode("//TP_CNTR_SWAP").Text) <> vbNullString Then
                    plngTipoOperacao = enumTipoOperacaoLQS.RegistroContratoSWAPComOpcaoBarreira
                End If
            End If
            
    End Select
        
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoOperacaoEnvioTED Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verifica se a operação é com ordem de pagamento, conciliação com CTP1002
Public Function flVerificarConciliacaoCETIP(ByRef pxmlDomRemessa As MSXML2.DOMDocument40, _
                                            ByRef plngTipoOperacao As Long, _
                                            Optional ByRef pstrCodigoMensagemSPB As String) As Boolean

Dim strSQL                                  As String
Dim strXML                                  As String
Dim strXMLMesg                              As String
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim xmlTipoMesgSPB                          As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    flVerificarConciliacaoCETIP = False
    
    If pxmlDomRemessa.documentElement.selectSingleNode("//SG_SIST_CETIP") Is Nothing Then
        Exit Function
    End If
    If pxmlDomRemessa.documentElement.selectSingleNode("//CO_OPER_CETIP") Is Nothing Then
        Exit Function
    End If
    If pxmlDomRemessa.documentElement.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
        Exit Function
    End If

    strSQL = " Select   TP_OPER " & vbCrLf & _
             " From     A8.TB_TIPO_OPER_CNCL_CETIP " & vbCrLf & _
             " Where    SG_SIST_CETIP       =     '" & pxmlDomRemessa.documentElement.selectSingleNode("//SG_SIST_CETIP").Text & "'" & vbCrLf & _
             " And      CO_OPER_CETIP       =      " & pxmlDomRemessa.documentElement.selectSingleNode("//CO_OPER_CETIP").Text & vbCrLf & _
             " And      IN_OPER_DEBT_CRED   =      " & pxmlDomRemessa.documentElement.selectSingleNode("//IN_OPER_DEBT_CRED").Text & vbCrLf & _
             " And      TP_MESG_RECB_INTE   =     '" & CLng(pxmlDomRemessa.documentElement.selectSingleNode("//TP_MESG").Text) & "'"

    strXML = fgQueryXMLLerTodos("TipoOperacaoConciliacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Trim(strXML) = vbNullString Then
        Exit Function
    End If

    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlTipoOperacao.loadXML strXML

    plngTipoOperacao = CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text)
    
    If plngTipoOperacao = enumTipoOperacaoLQS.MovimentacaoInstrumentoFinanceiroCTP4001 Then
    
        strSQL = " Select   CO_MESG_SPB_REGT_OPER " & vbCrLf & _
                 " From     A8.TB_TIPO_OPER " & vbCrLf & _
                 " Where    TP_OPER       =     '" & plngTipoOperacao & "'"
    
        strXMLMesg = fgQueryXMLLerTodos("TipoOperacaoConciliacao", strSQL, App.EXEName & "." & TypeName(Me))
    
        Set xmlTipoMesgSPB = CreateObject("MSXML2.DOMDocument.4.0")
        xmlTipoMesgSPB.loadXML strXMLMesg
        
        pstrCodigoMensagemSPB = xmlTipoMesgSPB.documentElement.selectSingleNode("//CO_MESG_SPB_REGT_OPER").Text
        
    End If
    
    Set xmlTipoOperacao = Nothing
    Set xmlTipoMesgSPB = Nothing
    
    flVerificarConciliacaoCETIP = True
    Exit Function

ErrorHandler:
    Set xmlTipoOperacao = Nothing
    Set xmlTipoMesgSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificarConciliacaoCETIP Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter o tipo de ooperação de retorno ao legado ( tipo mensagem de resposta - A7 )

Public Function ObterTipoOperacaoRetorno(ByVal plngTipoMensagem As Long) As Long

Dim strSQL                                  As String
Dim rsTipoOperacao                          As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " Select   TP_MESG_RETN_INTE " & vbCrLf & _
             " From     A8.TB_TIPO_OPER " & vbCrLf & _
             " Where    TP_MESG_RECB_INTE  = '" & plngTipoMensagem & "'"

    Set rsTipoOperacao = fgQuerySQL(strSQL)

    If rsTipoOperacao.EOF Then
        ObterTipoOperacaoRetorno = 0
    Else
        ObterTipoOperacaoRetorno = rsTipoOperacao!TP_MESG_RETN_INTE
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoOperacaoRetorno Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter tipo de operção para conciliação SELIC

Public Function ObterTiposOperacaoConciliacao( _
                    Optional plngLocalLiquidacao As Long, _
                    Optional pintRotinaAbertura As enumIndicadorSimNao = enumIndicadorSimNao.sim) As String

Dim strSQL                                  As String
Dim intTipoBackOffice                       As Integer

On Error GoTo ErrorHandler

    intTipoBackOffice = fgObterTipoBackOfficeUsuario

    strSQL = " Select    TP_OPER,                                         " & vbNewLine & _
             "           NO_TIPO_OPER,                                    " & vbNewLine & _
             "           CO_OPER_SELIC,                               " & vbNewLine & _
             "           CO_LOCA_LIQU                                     " & vbNewLine & _
             " From      A8.TB_TIPO_OPER                                  " & vbNewLine & _
             " Where     1=1                                              " & vbNewLine

    If plngLocalLiquidacao <> 0 Then
        strSQL = strSQL & " AND      CO_LOCA_LIQU       = " & plngLocalLiquidacao & " " & vbNewLine
    End If
    
    'Regra para restringir a Conciliação de determinados tipos de operação para o Tipo BackOffice respectivo
    '(Ex.: Despesa somente para o TipoBackOffice TESOURARIA)
    strSQL = strSQL & _
             " And      (TP_BKOF IS NULL OR TP_BKOF = " & intTipoBackOffice & ")" & vbNewLine
             
        Select Case pintRotinaAbertura
            Case enumIndicadorSimNao.sim
                strSQL = strSQL & _
                    " And       CO_OPER_SELIC  IS NOT NULL                   " & vbNewLine
            
            Case enumIndicadorSimNao.Nao
                strSQL = strSQL & _
                    " And       CO_OPER_SELIC  IS NULL                       " & vbNewLine
            
            Case enumIndicadorSimNao.Ambos
                'Não faz nada
                
        End Select
    
    strSQL = strSQL & _
             " Order By  NO_TIPO_OPER       ASC"

    ObterTiposOperacaoConciliacao = fgQueryXMLLerTodos("TipoOperacaoConciliacao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTiposOperacaoConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'recuperar um registro da tabela de Acao A8.TB_TIPO_OPER_ACAO_MESG_SPB

Public Function LerTipoOperacaoAcao(ByVal plngTipoOperacao As Long, _
                                    ByVal plngTipoAcao As Long) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandLerTipoOperacaoAcao

    strSQL = " SELECT   TP_OPER, " & vbCrLf & _
             "          TP_ACAO_OPER_ATIV_EXEC, " & vbCrLf & _
             "          CO_MESG_SPB_REGT_ACAO " & vbCrLf & _
             " FROM     A8.TB_TIPO_OPER_ACAO_MESG_SPB " & vbCrLf & _
             " WHERE    TP_OPER                 =   " & plngTipoOperacao & vbCrLf & _
             " AND      TP_ACAO_OPER_ATIV_EXEC  =   " & plngTipoAcao
    
    LerTipoOperacaoAcao = fgQueryXMLLer("Grupo_TipoOperacaoAcao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandLerTipoOperacaoAcao:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTipoOperacaoAcao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Bloquear a tabela de tipo de operação para inclusão de um operação

Public Sub BloquearTipoOperacaoParaInclusao()

Dim rsTipoOper                              As ADODB.Recordset
Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT   TP_OPER                         " & vbNewLine & _
             " FROM     A8.TB_TIPO_OPER_ACAO_MESG_SPB   " & vbNewLine & _
             " FOR UPDATE "

    Set rsTipoOper = fgQuerySQL(strSQL)

    Exit Sub

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "BloquearTipoOperacaoParaInclusao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Sub

'Obter Controle de previsão

Public Function ObterControlePrevisoes() As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT a.tp_oper,                                             " & vbCrLf & _
             "        b.no_tipo_oper,                                        " & vbCrLf & _
             "        a.co_situ_proc,                                        " & vbCrLf & _
             "        c.de_situ_proc,                                        " & vbCrLf & _
             "        a.no_proc_oper_ativ,                                   " & vbCrLf & _
             "        a.in_envi_prev_pj,                                     " & vbCrLf & _
             "        a.tp_liqu_oper_ativ,                                   " & vbCrLf & _
             "        d.no_tipo_liqu_oper_ativ,                              " & vbCrLf & _
             "        a.in_envi_prev_a6,                                     " & vbCrLf & _
             "        a.in_envi_relz_pj,                                     " & vbCrLf & _
             "        a.in_envi_relz_soli_a6,                                " & vbCrLf & _
             "        a.in_envi_relz_conf_a6,                                " & vbCrLf & _
             "        a.in_veri_regr_conf,                                   " & vbCrLf & _
             "        a.in_veri_regr_cncl,                                   " & vbCrLf & _
             "        a.in_veri_regr_libe,                                   " & vbCrLf & _
             "        a.in_envi_mesg_retn,                                   " & vbCrLf & _
             "        a.in_envi_mesg_spb,                                    " & vbCrLf & _
             "        a.in_disp_lanc_cnta_crrt,                              " & vbCrLf & _
             "        a.in_esto_pj_a6,                                       " & vbCrLf & _
             "        a.in_envi_aler                                         " & vbCrLf
    
    strSQL = strSQL & _
             " FROM   a8.tb_ctrl_proc_oper_ativ a,                           " & vbCrLf & _
             "        a8.tb_tipo_oper b,                                     " & vbCrLf & _
             "        a8.tb_situ_proc c,                                     " & vbCrLf & _
             "        a8.tb_tipo_liqu_oper_ativ d                            " & vbCrLf & _
             " WHERE  a.tp_oper                  = b.tp_oper (+)             " & vbCrLf & _
             " AND    a.co_situ_proc             = c.co_situ_proc (+)        " & vbCrLf & _
             " AND    DECODE(a.tp_liqu_oper_ativ, 0, NULL, a.tp_liqu_oper_ativ) = d.tp_liqu_oper_ativ (+)   " & vbCrLf
    
    strSQL = strSQL & _
             " ORDER BY a.tp_oper, a.no_proc_oper_ativ, a.co_situ_proc    "
    
    ObterControlePrevisoes = fgQueryXMLLerTodos("Grupo_ControlePrevisoes", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterControlePrevisoes Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar registros da tabela de controle dos processos de um operação (A8.TB_CTRL_PROC_OPER_ATIV)

Public Function LerProcessosOperacao() As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT DISTINCT NO_PROC_OPER_ATIV " & vbCrLf & _
             " FROM     A8.TB_CTRL_PROC_OPER_ATIV " & vbCrLf & _
             " ORDER BY NO_PROC_OPER_ATIV "
    
    LerProcessosOperacao = fgQueryXMLLerTodos("ProcessosOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerProcessosOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter os tipos de mensagens

Public Function ObterLayoutsOperacao(Optional ByVal pblnEntrada As Boolean = True) As String

Dim strSQL                                  As String
Dim strCampo                                As String

On Error GoTo ErrorHandler

    strCampo = IIf(pblnEntrada, "TP_MESG_RECB_INTE", "TP_MESG_RETN_INTE")

    strSQL = " SELECT DISTINCT " & strCampo & vbCrLf & _
             " FROM     A8.TB_TIPO_OPER " & vbCrLf & _
             " ORDER BY TO_NUMBER(" & strCampo & ")"
    
    ObterLayoutsOperacao = fgQueryXMLLerTodos("Grupo_LayoutsOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterLayoutsOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

