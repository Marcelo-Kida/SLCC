VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGradeHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : \r\nData Criação   : \r\nObjetivo       : \r\n\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F2945B10239"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
 
'' Objeto responsável pelo fornecimento de informações quanto às regras de horário
'' de funcionamento das câmaras e BACEN, bem como o controle de manutenção das
'' margens de segurança SLCC

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private strConstruct                        As String

Private objContext                       As COMSVCSLib.ObjectContext
Implements COMSVCSLib.ObjectControl
Implements COMSVCSLib.IObjectConstruct

Private Sub IObjectConstruct_Construct(ByVal pobjContruct As Object)
    
On Error Resume Next
    
    strConstruct = pobjContruct.ConstructString

End Sub

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

' Este método é responsável pela atualização geral da tabela de grade de horário e leitura da mesma
Public Function Executar(ByRef pxmlGradeHorario As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = pxmlGradeHorario.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Alterar"
            Executar = flAlterar(pxmlGradeHorario)
        Case "Excluir"
            Executar = flExcluir(pxmlGradeHorario)
        Case "Incluir"
            Executar = flIncluir(pxmlGradeHorario)
        Case "Ler"
            Executar = Ler(pxmlGradeHorario.selectSingleNode("//*/CO_GRAD_HORA").Text, _
                           pxmlGradeHorario.selectSingleNode("//*/SQ_ISPB").Text)
            
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
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura da tabela de grade de horario somente para obter as propriedades
Public Function ObterPropriedades() As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

        strSQL = " SELECT  CO_GRAD_HORA,             " & vbNewLine & _
                 "         SQ_ISPB,                  " & vbNewLine & _
                 "         QT_TEMP_MARG_SEGR,        " & vbNewLine & _
                 "         IN_SITU_GRAD_HORA,        " & vbNewLine & _
                 "         CO_USUA_ULTI_ATLZ,        " & vbNewLine & _
                 "         CO_ETCA_TRAB_ULTI_ATLZ,   " & vbNewLine & _
                 "         DH_ULTI_ATLZ              " & vbNewLine & _
                 " FROM    A8.TB_GRAD_HORA_INTE_MESG "

    ObterPropriedades = fgPropriedades("Grupo_GradeHorario", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura da tabela de grade de horario conforme filtro especificado
Public Function Ler(ByVal pstrCO_GRAD_HORA As String, _
                    ByVal plngSQ_ISPB As Long) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler
    
        strSQL = " SELECT  QT_TEMP_MARG_SEGR,        " & vbNewLine & _
                 "         IN_SITU_GRAD_HORA,        " & vbNewLine & _
                 "         CO_GRAD_HORA,             " & vbNewLine & _
                 "         SQ_ISPB,                  " & vbNewLine & _
                 "         CO_USUA_ULTI_ATLZ,        " & vbNewLine & _
                 "         CO_ETCA_TRAB_ULTI_ATLZ,   " & vbNewLine & _
                 "         DH_ULTI_ATLZ              " & vbNewLine & _
                 " FROM    A8.TB_GRAD_HORA_INTE_MESG " & vbNewLine & _
                 " WHERE   CO_GRAD_HORA  = '" & pstrCO_GRAD_HORA & "'" & vbNewLine & _
                 " AND     SQ_ISPB       =  " & plngSQ_ISPB

    Ler = fgQueryXMLLer("Grupo_GradeHorario", strSQL, App.EXEName & "." & TypeName(Me))
    
    'If Len(Ler) = 0 Then
       '4002 - Registro Inexistente ou excluído por outro usuário.
    '   lngCodigoErroNegocio = 4002
    '   GoTo ErrorHandler
    'End If

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela alteração das informações na tabela de grade de horário
Private Function flAlterar(ByVal pxmlGradeHorario As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    With pxmlGradeHorario.documentElement
        strSQL = " UPDATE  A8.TB_GRAD_HORA_INTE_MESG                                                         " & vbNewLine & _
                 " SET     QT_TEMP_MARG_SEGR         =  " & .selectSingleNode("QT_TEMP_MARG_SEGR").Text & ", " & vbNewLine & _
                 "         IN_SITU_GRAD_HORA         =  " & .selectSingleNode("IN_SITU_GRAD_HORA").Text & ", " & vbNewLine & _
                 "         CO_USUA_ULTI_ATLZ         = '" & fgUsuarioRede & "',                              " & vbNewLine & _
                 "         CO_ETCA_TRAB_ULTI_ATLZ    = '" & fgEstacaoTrabalhoUsuario & "',                   " & vbNewLine & _
                 "         DH_ULTI_ATLZ              = SYSDATE                                               " & vbNewLine & _
                 " WHERE   CO_GRAD_HORA              = '" & .selectSingleNode("CO_GRAD_HORA").Text & "'      " & vbNewLine & _
                 " AND     SQ_ISPB                   =  " & .selectSingleNode("SQ_ISPB").Text & "            " & vbNewLine & _
                 " AND     DH_ULTI_ATLZ              =  " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text)
    End With
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado ou excluído por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela exclusão das informações na tabela de grade de horário conforme filtro especificado
Private Function flExcluir(ByVal pxmlGradeHorario As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    With pxmlGradeHorario.documentElement
        strSQL = " DELETE                                                                    " & vbNewLine & _
                 " FROM    A8.TB_GRAD_HORA_INTE_MESG                                         " & vbNewLine & _
                 " WHERE   CO_GRAD_HORA = '" & .selectSingleNode("CO_GRAD_HORA").Text & "'   " & vbNewLine & _
                 " AND     SQ_ISPB      =  " & .selectSingleNode("SQ_ISPB").Text & ""
    End With
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado ou excluído por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela inclusão das informações na tabela de grade de horário
Private Function flIncluir(ByVal pxmlGradeHorario As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    With pxmlGradeHorario.documentElement
        strSQL = " INSERT INTO A8.TB_GRAD_HORA_INTE_MESG (               " & vbNewLine & _
                 "      CO_GRAD_HORA,                                    " & vbNewLine & _
                 "      SQ_ISPB,                                         " & vbNewLine & _
                 "      QT_TEMP_MARG_SEGR,                               " & vbNewLine & _
                 "      IN_SITU_GRAD_HORA,                               " & vbNewLine & _
                 "      CO_USUA_ULTI_ATLZ,                               " & vbNewLine & _
                 "      CO_ETCA_TRAB_ULTI_ATLZ,                          " & vbNewLine & _
                 "      DH_ULTI_ATLZ                                     " & vbNewLine & _
                 " ) VALUES (                                            " & vbNewLine & _
                 " '" & .selectSingleNode("CO_GRAD_HORA").Text & "',     " & vbNewLine & _
                 "  " & .selectSingleNode("SQ_ISPB").Text & ",           " & vbNewLine & _
                 "  " & .selectSingleNode("QT_TEMP_MARG_SEGR").Text & ", " & vbNewLine & _
                 "  " & .selectSingleNode("IN_SITU_GRAD_HORA").Text & ", " & vbNewLine & _
                 " '" & fgUsuarioRede & "',                              " & vbNewLine & _
                 " '" & fgEstacaoTrabalhoUsuario & "',                   " & vbNewLine & _
                 "      SYSDATE)"
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura das informações na tabela de grade de horário conforme filtro especificado
Public Function ObterDetalhesGradeHorario(ByVal pstrGradeHorario As String) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler
    
    If strConstruct = "OFF" Then
        ObterDetalhesGradeHorario = vbNullString
        Exit Function
    End If
    
    strSQL = " SELECT   A.CO_GRAD_HORA,                             " & vbNewLine & _
             "          A.SQ_ISPB,                                  " & vbNewLine & _
             "          A.IN_TIPO_GRAD,                             " & vbNewLine & _
             "          A.DT_EMIS_GRAD_BACEN,                       " & vbNewLine & _
             "          A.HO_ABER,                                  " & vbNewLine & _
             "          A.HO_ENCE,                                  " & vbNewLine & _
             "          B.NO_ISPB,                                  " & vbNewLine & _
             "          C.QT_TEMP_MARG_SEGR,                        " & vbNewLine & _
             "          C.IN_SITU_GRAD_HORA,                        " & vbNewLine & _
             "          C.DH_ULTI_ATLZ                              " & vbNewLine & _
             " FROM     A8.TB_GRADE_HORARIO       A,                " & vbNewLine & _
             "          A8.TB_INSTITUICAO_SPB     B,                " & vbNewLine & _
             "          A8.TB_GRAD_HORA_INTE_MESG C                 " & vbNewLine & _
             " WHERE    A.SQ_ISPB            =  B.SQ_ISPB                       " & vbNewLine & _
             " AND      A.CO_GRAD_HORA       =  C.CO_GRAD_HORA (+)         " & vbNewLine & _
             " AND      A.SQ_ISPB            =  C.SQ_ISPB      (+)         " & vbNewLine & _
             " AND      A.CO_GRAD_HORA       =  '" & pstrGradeHorario & "' " & vbNewLine & _
             " AND     (A.DT_INIC_VIGE_GRAD <=  " & fgDataHoraServidor_To_Oracle & vbNewLine & _
             " AND     (A.DT_FIM_VIGE_GRAD  IS  NULL " & vbNewLine & _
             " OR       A.DT_FIM_VIGE_GRAD  >=  " & fgDataHoraServidor_To_Oracle & "))" & vbNewLine & _
             " ORDER BY B.NO_ISPB ASC, IN_TIPO_GRAD ASC, DT_EMIS_GRAD_BACEN DESC"
    
    ObterDetalhesGradeHorario = fgQueryXMLLerTodos("GradeHorario", strSQL, App.EXEName & "." & TypeName(Me))
    
    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalhesGradeHorario Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura do codigo da mensagem SPB conforme filtro especificado
Public Function ObterMensagensSPB(ByVal pstrGradeHorario As String) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler
    
    strSQL = " SELECT   B.CO_MESG,                                  " & vbNewLine & _
             "          B.NO_MESG                                   " & vbNewLine & _
             " FROM     A8.TB_GRADE_MENSAGEM    A,                     " & vbNewLine & _
             "          A8.TB_MENSAGEM          B                      " & vbNewLine & _
             " WHERE    A.SQ_MESG = B.SQ_MESG                       " & vbNewLine & _
             " AND      A.CO_GRAD_HORA = '" & pstrGradeHorario & "' " & vbNewLine & _
             " ORDER BY B.CO_MESG ASC"
    
    ObterMensagensSPB = fgQueryXMLLerTodos("MensagemSPB", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagensSPB Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura de todos os registros das tabelas de Tipo TAG e Dominio
Public Function LerTodos() As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler
    
    strSQL = " SELECT   B.CO_DOMI,                      " & vbNewLine & _
             "          B.DE_DOMI                       " & vbNewLine & _
             " FROM     A8.TB_TIPO_TAG A,                  " & vbNewLine & _
             "          A8.TB_DOMINIO  B                   " & vbNewLine & _
             " WHERE    A.SQ_TIPO_TAG = B.SQ_TIPO_TAG   " & vbNewLine & _
             " AND      A.NO_TIPO_TAG = 'CodGrd'        " & vbNewLine & _
             " ORDER BY B.CO_DOMI ASC"
    
    LerTodos = fgQueryXMLLerTodos("GradeHorario", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela obtenção da grade de horário conforme o codigo da mensagem SPB especificada no filtro
Public Function ObterGradeHorario(ByVal pstrCodigoMensagem As String, _
                                  ByVal plngSequencialISPBDestino As Long) As String

Dim strSQL                                  As String
Dim strRetorno                              As String
Dim strRetornoLiberado                      As String
Dim objMensagemSPB                          As A8LQS.clsMensagemSPB
Dim xmlDomGradeHorario                      As MSXML2.DOMDocument40
Dim xmlDomRetorno                           As MSXML2.DOMDocument40
Dim blnGradeEventual                        As Boolean
Dim intMargemSeguranca                      As Integer
Dim dtaHorarioEncerramento                  As Date
Dim lngSequencialMensagem                   As Long

    On Error GoTo ErrorHandler
    
    strRetornoLiberado = "<Repeat_GradeHorario>" & _
                            "<Grupo_GradeHorario>" & _
                                "<HorarioAbertura>" & _
                                    fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000" & _
                                "</HorarioAbertura>" & _
                                "<HorarioEncerramento>" & _
                                    fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959" & _
                                "</HorarioEncerramento>" & _
                                "<MargemSeguranca>" & _
                                    "0" & _
                                "</MargemSeguranca>" & _
                            "</Grupo_GradeHorario>" & _
                         "</Repeat_GradeHorario>"
    
    If strConstruct = "OFF" Then
        ObterGradeHorario = strRetornoLiberado
        Exit Function
    End If
    
    Set objMensagemSPB = CreateObject("A8LQS.clsMensagemSPB")
    lngSequencialMensagem = objMensagemSPB.ObterSequenciaMensagem(pstrCodigoMensagem)
    Set objMensagemSPB = Nothing
    
    '--------------------------------------------------------------------------------------------------------
    '>>> Obtem Grade de Horário da Mensageria, considerando a Margem de Segurança configurada no SLCC
    '
    ' Obs.: Concatena a coluna << DT_EMIS_GRAD_BACEN >> nos horários, apenas para o retorno de uma
    '       data válida MAIS ATUAL, pois a função DATESERIAL do VB só funciona a partir do ano 1900
    '--------------------------------------------------------------------------------------------------------
    strSQL = " SELECT   TO_CHAR(A.DT_EMIS_GRAD_BACEN, 'YYYYMMDD')   DT_EMIS_GRAD_BACEN,     " & vbNewLine & _
             "          A.IN_TIPO_GRAD,                                                     " & vbNewLine & _
             "          TO_CHAR(A.DT_EMIS_GRAD_BACEN, 'YYYYMMDD') ||                        " & vbNewLine & _
             "          TO_CHAR(A.HO_ABER, 'HH24MISS')              HO_ABER,                " & vbNewLine & _
             "          TO_CHAR(A.DT_EMIS_GRAD_BACEN, 'YYYYMMDD') ||                        " & vbNewLine & _
             "          TO_CHAR(A.HO_ENCE, 'HH24MISS')              HO_ENCE,                " & vbNewLine & _
             "          NVL(C.QT_TEMP_MARG_SEGR, 0)                 QT_TEMP_MARG_SEGR,      " & vbNewLine & _
             "          NVL(C.IN_SITU_GRAD_HORA, 0)                 IN_SITU_GRAD_HORA       " & vbNewLine & _
             " FROM     A8.TB_GRADE_HORARIO         A,                                      " & vbNewLine & _
             "          A8.TB_GRADE_MENSAGEM        B,                                      " & vbNewLine & _
             "          A8.TB_GRAD_HORA_INTE_MESG   C                                       " & vbNewLine & _
             " WHERE    A.CO_GRAD_HORA              = B.CO_GRAD_HORA                        " & vbNewLine & _
             " AND      A.CO_GRAD_HORA              = C.CO_GRAD_HORA    (+)                 " & vbNewLine & _
             " AND      A.SQ_ISPB                   = C.SQ_ISPB         (+)                 " & vbNewLine & _
             " AND      B.SQ_MESG                   = " & lngSequencialMensagem & "         " & vbNewLine & _
             " AND      A.SQ_ISPB                   = " & plngSequencialISPBDestino & "     " & vbNewLine & _
             " AND     (A.DT_EMIS_GRAD_BACEN, A.IN_TIPO_GRAD)   IN                          " & vbNewLine & _
             "         (SELECT  MAX(A.DT_EMIS_GRAD_BACEN), A.IN_TIPO_GRAD                   " & vbNewLine & _
             "          FROM    A8.TB_GRADE_HORARIO        A,                               " & vbNewLine & _
             "                  A8.TB_GRADE_MENSAGEM       B                                " & vbNewLine & _
             "          WHERE   A.CO_GRAD_HORA          = B.CO_GRAD_HORA                    " & vbNewLine & _
             "          AND     B.SQ_MESG               = " & lngSequencialMensagem & "     " & vbNewLine & _
             "          AND     A.SQ_ISPB               = " & plngSequencialISPBDestino & " " & vbNewLine & _
             "          GROUP BY                                                            " & vbNewLine & _
             "                  A.IN_TIPO_GRAD)"
    
    strRetorno = fgQueryXMLLerTodos("GradeHorario", strSQL, App.EXEName & "." & TypeName(Me))
    
    If strRetorno = vbNullString Then
        ObterGradeHorario = strRetornoLiberado
        Exit Function
        '>>> Tratamento para grade de horário inexistente DESATIVADO.
        '    Quando não existir, retornar GRADE HORÁRIO CHEIA (00:00:00 a 23:59:59)
        '--------------------------------------------------------------------------------------
        'Grade de Horário Inexistente.
        'lngCodigoErroNegocio = 3019
        'GoTo ErrorHandler
        '--------------------------------------------------------------------------------------
    End If
    
    '>>> Monta XML de Retorno -----------------------------------------------------------------
    Set xmlDomRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlDomRetorno, "", "Repeat_GradeHorario", "")
    Call fgAppendNode(xmlDomRetorno, "Repeat_GradeHorario", "Grupo_GradeHorario", "")
    '------------------------------------------------------------------------------------------
    
    Set xmlDomGradeHorario = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlDomGradeHorario.loadXML(strRetorno)
    
    If Val(xmlDomGradeHorario.selectSingleNode("//IN_SITU_GRAD_HORA").Text) = enumIndicadorSimNao.nao Then
        ObterGradeHorario = strRetornoLiberado
        Exit Function
    End If
    
    With xmlDomGradeHorario
        'Verifica se o XML possui grade de horário eventual...
        If Not .selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario[IN_TIPO_GRAD='2']") Is Nothing Then
            '...se sim, verifica se a grade EVENTUAL é para a data atual
            If fgDtXML_To_Date(.selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario[IN_TIPO_GRAD='2']/DT_EMIS_GRAD_BACEN").Text) = _
               fgDataHoraServidor(enumFormatoDataHora.Data) Then
            
                Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                                 "HorarioAbertura", _
                                                 .selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario[IN_TIPO_GRAD='2']/HO_ABER").Text)
                
                intMargemSeguranca = Val(.selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario[IN_TIPO_GRAD='2']/QT_TEMP_MARG_SEGR").Text)
                
                'Verifica se a Margem de Segurança SLCC foi configurada...
                If intMargemSeguranca <> 0 Then
                    'se sim, deduz a margem de segurança do horário de encerramento
                    dtaHorarioEncerramento = fgDtHrStr_To_DateTime(.selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario[IN_TIPO_GRAD='2']/HO_ENCE").Text)
                    
                    dtaHorarioEncerramento = DateAdd("n", -intMargemSeguranca, dtaHorarioEncerramento)
                    
                    Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                                     "HorarioEncerramento", _
                                                     fgDtHr_To_Xml(dtaHorarioEncerramento))
                
                '...se não, retorna o encerramento da grade (mensageria)
                Else
                    Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                                     "HorarioEncerramento", _
                                                     .selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario[IN_TIPO_GRAD='2']/HO_ENCE").Text)
                End If
                
                blnGradeEventual = True
            End If
        End If
        
        '...se não, retorna a grade PADRÃO
        If Not blnGradeEventual Then
            Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                             "HorarioAbertura", _
                                             .selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario[IN_TIPO_GRAD='1']/HO_ABER").Text)
            
            intMargemSeguranca = Val(.selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario[IN_TIPO_GRAD='1']/QT_TEMP_MARG_SEGR").Text)
            
            'Verifica se a Margem de Segurança SLCC foi configurada...
            If intMargemSeguranca <> 0 Then
                'se sim, deduz a margem de segurança do horário de encerramento
                dtaHorarioEncerramento = fgDtHrStr_To_DateTime(.selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario[IN_TIPO_GRAD='1']/HO_ENCE").Text)
                
                dtaHorarioEncerramento = DateAdd("n", -intMargemSeguranca, dtaHorarioEncerramento)
                
                Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                                 "HorarioEncerramento", _
                                                 fgDtHr_To_Xml(dtaHorarioEncerramento))
            
            '...se não, retorna o encerramento da grade (mensageria)
            Else
                Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                                 "HorarioEncerramento", _
                                                 .selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario[IN_TIPO_GRAD='1']/HO_ENCE").Text)
            End If
        End If
    End With
    
    'Envia a margem de segurança utilizada, se não existir, o valor será ZERO
    Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                     "MargemSeguranca", _
                                     intMargemSeguranca)
    
    ObterGradeHorario = xmlDomRetorno.xml
    
    Set xmlDomGradeHorario = Nothing
    Set xmlDomRetorno = Nothing
    
    Exit Function
    
ErrorHandler:
    
    Set objMensagemSPB = Nothing
    Set xmlDomGradeHorario = Nothing
    Set xmlDomRetorno = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterGradeHorario Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela obtenção da grade de horário para as Camaras BMA e CETIP, conforme o codigo da mensagem SPB especificada no filtro
Public Function ObterGradeHorarioBMACetip(ByVal pstrCodigoMensagem As String) As String

Dim strSQL                                  As String
Dim strRetorno                              As String
Dim strRetornoLiberado                      As String
Dim objMensagemSPB                          As A8LQS.clsMensagemSPB
Dim xmlDomGradeHorario                      As MSXML2.DOMDocument40
Dim xmlDomRetorno                           As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode
Dim intMargemSeguranca                      As Integer
Dim dtaHorarioEncerramento                  As Date
Dim lngSequencialMensagem                   As Long

On Error GoTo ErrorHandler

    strRetornoLiberado = "<Repeat_GradeHorario>" & _
                            "<Grupo_GradeHorario>" & _
                                "<HorarioAbertura>" & _
                                    fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000" & _
                                "</HorarioAbertura>" & _
                                "<HorarioEncerramento>" & _
                                    fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959" & _
                                "</HorarioEncerramento>" & _
                                "<MargemSeguranca>" & _
                                    "0" & _
                                "</MargemSeguranca>" & _
                            "</Grupo_GradeHorario>" & _
                         "</Repeat_GradeHorario>"

    If strConstruct = "OFF" Then
        ObterGradeHorarioBMACetip = strRetornoLiberado
        Exit Function
    End If

    Set objMensagemSPB = CreateObject("A8LQS.clsMensagemSPB")
    lngSequencialMensagem = objMensagemSPB.ObterSequenciaMensagem(pstrCodigoMensagem)
    Set objMensagemSPB = Nothing

    '--------------------------------------------------------------------------------------------------------
    '>>> Obtem Grade de Horário da Mensageria, considerando a Margem de Segurança configurada no SLCC
    '
    ' Obs.: Concatena a coluna << DT_EMIS_GRAD_BACEN >> nos horários, apenas para o retorno de uma
    '       data válida MAIS ATUAL, pois a função DATESERIAL do VB só funciona a partir do ano 1900
    '--------------------------------------------------------------------------------------------------------
    strSQL = " SELECT   TO_CHAR(A.DT_EMIS_GRAD_BACEN, 'YYYYMMDD')   DT_EMIS_GRAD_BACEN,     " & vbNewLine & _
             "          A.IN_TIPO_GRAD,                                                     " & vbNewLine & _
             "          TO_CHAR(A.DT_EMIS_GRAD_BACEN, 'YYYYMMDD') ||                        " & vbNewLine & _
             "          TO_CHAR(A.HO_ABER, 'HH24MISS')              HO_ABER,                " & vbNewLine & _
             "          TO_CHAR(A.DT_EMIS_GRAD_BACEN, 'YYYYMMDD') ||                        " & vbNewLine & _
             "          TO_CHAR(A.HO_ENCE, 'HH24MISS')              HO_ENCE,                " & vbNewLine & _
             "          NVL(C.QT_TEMP_MARG_SEGR, 0)                 QT_TEMP_MARG_SEGR,      " & vbNewLine & _
             "          NVL(C.IN_SITU_GRAD_HORA, 0)                 IN_SITU_GRAD_HORA       " & vbNewLine & _
             " FROM     A8.TB_GRADE_HORARIO            A,                                   " & vbNewLine & _
             "          A8.TB_GRADE_MENSAGEM           B,                                   " & vbNewLine & _
             "          A8.TB_GRAD_HORA_INTE_MESG   C                                       " & vbNewLine & _
             " WHERE    A.CO_GRAD_HORA              = B.CO_GRAD_HORA                        " & vbNewLine & _
             " AND      A.CO_GRAD_HORA              = C.CO_GRAD_HORA    (+)                 " & vbNewLine & _
             " AND      A.SQ_ISPB                   = C.SQ_ISPB         (+)                 " & vbNewLine & _
             " AND      B.SQ_MESG                   = " & lngSequencialMensagem & "         " & vbNewLine & _
             " AND     (A.DT_EMIS_GRAD_BACEN, A.IN_TIPO_GRAD)   IN                          " & vbNewLine & _
             "         (SELECT  MAX(A.DT_EMIS_GRAD_BACEN), A.IN_TIPO_GRAD                   " & vbNewLine & _
             "          FROM    A8.TB_GRADE_HORARIO        A,                               " & vbNewLine & _
             "                  A8.TB_GRADE_MENSAGEM       B                                " & vbNewLine & _
             "          WHERE   A.CO_GRAD_HORA          = B.CO_GRAD_HORA                    " & vbNewLine & _
             "          AND     B.SQ_MESG               = " & lngSequencialMensagem & "     " & vbNewLine & _
             "          GROUP BY                                                            " & vbNewLine & _
             "                  A.IN_TIPO_GRAD)                                             " & vbNewLine

    strSQL = strSQL & _
             " GROUP BY  TO_CHAR(A.DT_EMIS_GRAD_BACEN, 'YYYYMMDD'),                                    " & vbNewLine & _
             "           A.IN_TIPO_GRAD,                                                               " & vbNewLine & _
             "           TO_CHAR(A.DT_EMIS_GRAD_BACEN, 'YYYYMMDD') || TO_CHAR(A.HO_ABER, 'HH24MISS'),  " & vbNewLine & _
             "           TO_CHAR(A.DT_EMIS_GRAD_BACEN, 'YYYYMMDD') || TO_CHAR(A.HO_ENCE, 'HH24MISS'),  " & vbNewLine & _
             "           NVL(C.QT_TEMP_MARG_SEGR, 0),                                                  " & vbNewLine & _
             "           NVL(C.IN_SITU_GRAD_HORA, 0)                                                   " & vbNewLine & _
             " ORDER BY  A.IN_TIPO_GRAD DESC,                                                          " & vbNewLine & _
             "           TO_CHAR(A.DT_EMIS_GRAD_BACEN, 'YYYYMMDD') || TO_CHAR(A.HO_ABER, 'HH24MISS'),  " & vbNewLine & _
             "           TO_CHAR(A.DT_EMIS_GRAD_BACEN, 'YYYYMMDD') || TO_CHAR(A.HO_ENCE, 'HH24MISS')   " & vbNewLine

    strRetorno = fgQueryXMLLerTodos("GradeHorario", strSQL, App.EXEName & "." & TypeName(Me))
    
    If strRetorno = vbNullString Then
        ObterGradeHorarioBMACetip = strRetornoLiberado
        Exit Function
        '>>> Tratamento para grade de horário inexistente DESATIVADO.
        '    Quando não existir, retornar GRADE HORÁRIO CHEIA (00:00:00 a 23:59:59)
        '--------------------------------------------------------------------------------------
        'Grade de Horário Inexistente.
        'lngCodigoErroNegocio = 3019
        'GoTo ErrorHandler
        '--------------------------------------------------------------------------------------
    End If

    '>>> Monta XML de Retorno -----------------------------------------------------------------
    Set xmlDomRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlDomRetorno, "", "Repeat_GradeHorario", "")
    Call fgAppendNode(xmlDomRetorno, "Repeat_GradeHorario", "Grupo_GradeHorario", "")
    '------------------------------------------------------------------------------------------

    Set xmlDomGradeHorario = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlDomGradeHorario.loadXML(strRetorno)

    If Val(xmlDomGradeHorario.selectSingleNode("//IN_SITU_GRAD_HORA").Text) = enumIndicadorSimNao.nao Then
        ObterGradeHorarioBMACetip = strRetornoLiberado
        Exit Function
    End If
    
    For Each objNode In xmlDomGradeHorario.documentElement.selectNodes("//Repeat_GradeHorario/*")
        With objNode
            'Verifica se o XML possui grade de horário eventual...
            If .selectSingleNode("IN_TIPO_GRAD").Text = 2 Then
                '...se sim, verifica se a grade EVENTUAL é para a data atual
                If fgDtXML_To_Date(.selectSingleNode("DT_EMIS_GRAD_BACEN").Text) = _
                   fgDataHoraServidor(enumFormatoDataHora.Data) Then

                    Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                                     "HorarioAbertura", _
                                                     .selectSingleNode("HO_ABER").Text)

                    intMargemSeguranca = Val(.selectSingleNode("QT_TEMP_MARG_SEGR").Text)

                    'Verifica se a Margem de Segurança SLCC foi configurada...
                    If intMargemSeguranca <> 0 Then
                        'se sim, deduz a margem de segurança do horário de encerramento
                        dtaHorarioEncerramento = fgDtHrStr_To_DateTime(.selectSingleNode("HO_ENCE").Text)

                        dtaHorarioEncerramento = DateAdd("n", -intMargemSeguranca, dtaHorarioEncerramento)

                        Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                                         "HorarioEncerramento", _
                                                         fgDtHr_To_Xml(dtaHorarioEncerramento))

                    '...se não, retorna o encerramento da grade (mensageria)
                    Else
                        Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                                         "HorarioEncerramento", _
                                                         .selectSingleNode("HO_ENCE").Text)
                    End If
                    'Considerar a grade eventual
                    Exit For
                End If
            Else
                '...se não, retorna a grade PADRÃO
                Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                                 "HorarioAbertura", _
                                                 .selectSingleNode("HO_ABER").Text)

                intMargemSeguranca = Val(.selectSingleNode("QT_TEMP_MARG_SEGR").Text)

                'Verifica se a Margem de Segurança SLCC foi configurada...
                If intMargemSeguranca <> 0 Then
                    'se sim, deduz a margem de segurança do horário de encerramento
                    dtaHorarioEncerramento = fgDtHrStr_To_DateTime(.selectSingleNode("HO_ENCE").Text)

                    dtaHorarioEncerramento = DateAdd("n", -intMargemSeguranca, dtaHorarioEncerramento)

                    Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                                     "HorarioEncerramento", _
                                                     fgDtHr_To_Xml(dtaHorarioEncerramento))

                '...se não, retorna o encerramento da grade (mensageria)
                Else
                    Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                                     "HorarioEncerramento", _
                                                     .selectSingleNode("HO_ENCE").Text)
                End If
                Exit For
            End If
        End With
    Next

    'Envia a margem de segurança utilizada, se não existir, o valor será ZERO
    Call fgAppendNode(xmlDomRetorno, "Grupo_GradeHorario", _
                                     "MargemSeguranca", _
                                     intMargemSeguranca)

    ObterGradeHorarioBMACetip = xmlDomRetorno.xml

    Set xmlDomGradeHorario = Nothing
    Set xmlDomRetorno = Nothing
    
    Exit Function

ErrorHandler:
    Set objMensagemSPB = Nothing
    Set xmlDomGradeHorario = Nothing
    Set xmlDomRetorno = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterGradeHorarioBMACetip Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Este método é responsável pela verificação da grade de horário de envio da mensagem SPB conforme o codigo da mensagem SPB especificada no filtro
Public Function VerificarGradeHorarioEnvioMensagem(ByVal pstrCodigoMensagem As String, _
                                                   ByVal plngSequencialISPBDestino As Long, _
                                          Optional ByVal plngLocalLiquidacao As Long = 13) As Boolean

Dim xmlDomRetorno                           As MSXML2.DOMDocument40
Dim strHorarioAbertura                      As String
Dim strHorarioEncerramento                  As String
Dim strDataHoraAtual                        As String

    On Error GoTo ErrorHandler
    
    If strConstruct = "OFF" Then
        VerificarGradeHorarioEnvioMensagem = True
        Exit Function
    End If

    Set xmlDomRetorno = CreateObject("MSXML2.DOMDocument.4.0")

    If plngLocalLiquidacao = enumLocalLiquidacao.CETIP Or _
       plngLocalLiquidacao = enumLocalLiquidacao.BMD Or _
       plngLocalLiquidacao = enumLocalLiquidacao.BMC Then
        Call xmlDomRetorno.loadXML(ObterGradeHorarioBMACetip(pstrCodigoMensagem))
    
    ElseIf plngLocalLiquidacao = enumLocalLiquidacao.BMA Then
        VerificarGradeHorarioEnvioMensagem = True
        Exit Function
    
    Else
        Call xmlDomRetorno.loadXML(ObterGradeHorario(pstrCodigoMensagem, plngSequencialISPBDestino))
    
    End If

    'Captura data e hora atual
    strDataHoraAtual = fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora))

    'Captura a grade de horário
    strHorarioAbertura = Mid(xmlDomRetorno.selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario/HorarioAbertura").Text, 9, 6)
    strHorarioEncerramento = Mid(xmlDomRetorno.selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario/HorarioEncerramento").Text, 9, 6)

    If TimeSerial(Mid(strDataHoraAtual, 9, 2), Mid(strDataHoraAtual, 11, 2), Mid(strDataHoraAtual, 13, 2)) < _
       TimeSerial(Mid(strHorarioAbertura, 1, 2), Mid(strHorarioAbertura, 3, 2), Mid(strHorarioAbertura, 5, 2)) Or _
       TimeSerial(Mid(strDataHoraAtual, 9, 2), Mid(strDataHoraAtual, 11, 2), Mid(strDataHoraAtual, 13, 2)) > _
       TimeSerial(Mid(strHorarioEncerramento, 1, 2), Mid(strHorarioEncerramento, 3, 2), Mid(strHorarioEncerramento, 5, 2)) Then

        VerificarGradeHorarioEnvioMensagem = False
    Else
        VerificarGradeHorarioEnvioMensagem = True
    End If

    Set xmlDomRetorno = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set xmlDomRetorno = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "VerificarGradeHorarioEnvioMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

