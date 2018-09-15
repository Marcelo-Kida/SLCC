VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVeiculoLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA\r\nPacote         :\r\nClasse         : clsVeiculoLegal\r\nData Criação   : 22/07/2003\r\nObjetivo       :\r\n\r\nAnalista       : Carlos Fortes\r\n\r\nProgramador    : Cassiano Nicolosi\r\nData           : 22/07/2003\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F15C88500D3"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"

'Objeto responsável pela manutenção do cadastro de veiculo legal

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

'Para Atualização de registros por Remessa. Não leva em consideração a data de atualização
Private Function flAtualizarPorLote(ByRef pxmlVeiculoLegal As MSXML2.DOMDocument40)
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    With pxmlVeiculoLegal.documentElement

        strSQL = " UPDATE A8.TB_VEIC_LEGA                                                             " & vbCrLf & _
                 " SET    CO_EMPR           = " & .selectSingleNode("CO_EMPR").Text & " ,             " & vbCrLf & _
                 "        TP_BKOF           = " & .selectSingleNode("TP_BKOF").Text & " ,             " & vbCrLf & _
                 "        NO_VEIC_LEGA      = '" & .selectSingleNode("NO_VEIC_LEGA").Text & "' ,      " & vbCrLf & _
                 "        NO_REDU_VEIC_LEGA = '" & .selectSingleNode("NO_REDU_VEIC_LEGA").Text & "' , " & vbCrLf & _
                 "        CO_CNPJ_VEIC_LEGA =  " & IIf(.selectSingleNode("CO_CNPJ_VEIC_LEGA").Text <> vbNullString, .selectSingleNode("CO_CNPJ_VEIC_LEGA").Text, " NULL ") & "  , " & vbCrLf & _
                 "        DT_INIC_VIGE      =  " & fgDtXML_To_Oracle(.selectSingleNode("DT_INIC_VIGE").Text) & " , " & vbCrLf
                 
        If .selectSingleNode("DT_FIM_VIGE").Text = vbNullString Or _
           .selectSingleNode("DT_FIM_VIGE").Text = 0 Then
            strSQL = strSQL & _
                "        DT_FIM_VIGE       =  " & " NULL , " & vbCrLf
        Else
            strSQL = strSQL & _
                "        DT_FIM_VIGE       =  " & fgDtXML_To_Oracle(.selectSingleNode("DT_FIM_VIGE").Text) & " , " & vbCrLf
        End If
        
        strSQL = strSQL & _
                 "        CO_USUA_ULTI_ATLZ = '" & fgObterUsuarioRede & "' ,                          " & vbCrLf & _
                 "        CO_ETCA_TRAB_ULTI_ATLZ = '" & fgObterEstacaoTrabalhoUsuario & "' ,               " & vbCrLf & _
                 "        DH_ULTI_ATLZ      =  SYSDATE                                                " & vbCrLf

        strSQL = strSQL & _
                 " WHERE  CO_VEIC_LEGA = '" & .selectSingleNode("CO_VEIC_LEGA").Text & "' " & vbCrLf & _
                 " AND    SG_SIST      = '" & .selectSingleNode("SG_SIST_ORIG").Text & "' "

    End With
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado ou excluído por outro usuário.
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "AtualizarLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Private Function flSQLCodigoGrupoPadrao(ByVal lngTipoBackOffice As Long) As String
    
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " ( SELECT CO_GRUP_VEIC_LEGA FROM A8.TB_GRUP_VEIC_LEGA " & vbCrLf & _
             "   WHERE  TP_BKOF = " & lngTipoBackOffice & vbCrLf & _
             "     AND  NO_GRUP_VEIC_LEGA = '" & strGRUPO_VEIC_LEGA_PADRAO & "' )"
             
    flSQLCodigoGrupoPadrao = strSQL

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flSQLCodigoGrupoPadrao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Para Inclusão de registros por Remessa. Não leva em consideração a data de atualização
Private Function flIncluirPorLote(ByRef pxmlVeiculoLegal As MSXML2.DOMDocument40) As Boolean

Dim strSQLCodigoGrupoPadrao                 As String
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A8.TB_VEIC_LEGA (        " & vbCrLf & _
             "             CO_VEIC_LEGA,            " & vbCrLf & _
             "             SG_SIST,                 " & vbCrLf & _
             "             CO_EMPR,                 " & vbCrLf & _
             "             CO_GRUP_VEIC_LEGA,       " & vbCrLf & _
             "             TP_BKOF,                 " & vbCrLf & _
             "             NO_VEIC_LEGA,            " & vbCrLf & _
             "             NO_REDU_VEIC_LEGA,       " & vbCrLf & _
             "             CO_CNPJ_VEIC_LEGA,       " & vbCrLf & _
             "             DT_INIC_VIGE,            " & vbCrLf & _
             "             DT_FIM_VIGE,             " & vbCrLf & _
             "             CO_USUA_ULTI_ATLZ,       " & vbCrLf & _
             "             CO_ETCA_TRAB_ULTI_ATLZ,  " & vbCrLf & _
             "             DH_ULTI_ATLZ             " & vbCrLf & _
             " ) VALUES (                           " & vbCrLf
             
    With pxmlVeiculoLegal.documentElement

        strSQL = strSQL & _
             " '" & .selectSingleNode("CO_VEIC_LEGA").Text & "' ,                   " & vbCrLf & _
             " '" & .selectSingleNode("SG_SIST_ORIG").Text & "' ,                   " & vbCrLf & _
             "  " & .selectSingleNode("CO_EMPR").Text & "  ,                        " & vbCrLf & _
             "  " & flSQLCodigoGrupoPadrao(.selectSingleNode("TP_BKOF").Text) & " , " & vbCrLf & _
             "  " & .selectSingleNode("TP_BKOF").Text & " , " & vbCrLf & _
             " '" & .selectSingleNode("NO_VEIC_LEGA").Text & "' ,                   " & vbCrLf & _
             " '" & .selectSingleNode("NO_REDU_VEIC_LEGA").Text & "' ,              " & vbCrLf & _
             "  " & .selectSingleNode("CO_CNPJ_VEIC_LEGA").Text & " ,               " & vbCrLf & _
             "  " & fgDtXML_To_Oracle(.selectSingleNode("DT_INIC_VIGE").Text) & " , " & vbCrLf
             
        If .selectSingleNode("DT_FIM_VIGE").Text = vbNullString Or _
            .selectSingleNode("DT_FIM_VIGE").Text = 0 Then
           strSQL = strSQL & _
               " NULL , " & vbCrLf
        Else
           strSQL = strSQL & _
               " " & fgDtXML_To_Oracle(.selectSingleNode("DT_FIM_VIGE").Text) & " , " & vbCrLf
        End If
             
        strSQL = strSQL & _
             " '" & fgObterUsuarioRede & "' ,                                        " & vbCrLf & _
             " '" & fgObterEstacaoTrabalhoUsuario & "' ,                             " & vbCrLf & _
             " SYSDATE )" & vbCrLf
    End With

    fgExecuteSQL (strSQL)

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
 

'Processamento de remessas de cadastro de veiculo legal do legado
 
Public Function ProcessarRemessa(ByRef xmlDomRemessa As MSXML2.DOMDocument40) As String

Dim strDHAtualizacao                        As String
Dim blnExiste                               As Boolean
Dim strQuery                                As String
Dim rsVeiculoLegal                          As ADODB.Recordset

On Error GoTo ErrorHandler

    'Procura o registro para atualizar ou incluir
    strQuery = " SELECT CO_VEIC_LEGA, SG_SIST  " & vbCrLf & _
               " FROM A8.TB_VEIC_LEGA          " & vbCrLf & _
               " WHERE CO_VEIC_LEGA = '" & xmlDomRemessa.documentElement.selectSingleNode("CO_VEIC_LEGA").Text & "' " & vbCrLf & _
               " AND   SG_SIST      = '" & xmlDomRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text & "' "

    Set rsVeiculoLegal = fgQuerySQL(strQuery)

    If rsVeiculoLegal.EOF Then
        flIncluirPorLote xmlDomRemessa
    Else
        flAtualizarPorLote xmlDomRemessa
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set xmlDomRemessa = Nothing
    Set rsVeiculoLegal = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ProcessarLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Alterar o código do grupo de um veiculo legal

Public Function AlterarCodigoGrupo(ByRef pxmlVeiculoLegal As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler
            
    strSQL = " UPDATE A8.TB_VEIC_LEGA           " & _
            "  SET    CO_GRUP_VEIC_LEGA      =  " & pxmlVeiculoLegal.documentElement.selectSingleNode("CO_GRUP_VEIC_LEGA").Text & "," & _
            "         CO_USUA_ULTI_ATLZ      = '" & fgObterUsuarioRede & "'," & _
            "         CO_ETCA_TRAB_ULTI_ATLZ = '" & fgObterEstacaoTrabalhoUsuario & "'," & _
            "         DH_ULTI_ATLZ           =  SYSDATE " & _
            " WHERE   CO_VEIC_LEGA           = '" & pxmlVeiculoLegal.documentElement.selectSingleNode("CO_VEIC_LEGA").Text & "' " & _
            " AND     SG_SIST = '" & pxmlVeiculoLegal.documentElement.selectSingleNode("SG_SIST").Text & "'" & _
            " AND     DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pxmlVeiculoLegal.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado ou excluído por outro usuário.
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarCodigoGrupo Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Executar a operação : LerTodos

Public Function Executar(ByRef pxmlPropriedades As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlPropriedades.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pxmlPropriedades)

        Case "AlterarEmpresaVeiculoLegal"
            Executar = AlterarEmpresaVeiculoLegal(pxmlPropriedades)

        Case "LerTodos"
            Executar = LerTodos(, Val(pxmlPropriedades.selectSingleNode("//*/CO_GRUP_VEIC_LEGA").Text))
        
        Case "ObterDetalheVeiculoLegal"
            Call pxmlPropriedades.loadXML(pxmlPropriedades.selectSingleNode("Repeat_Filtros/Grupo_Filtros/Repeat_Filtros").xml)
            Executar = ObterDetalheVeiculoLegal(pxmlPropriedades.xml)

        Case Else
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

'Recuperar registros da tabela A8.TB_VEIC_LEGA

Public Function LerTodos(Optional ByVal pstrCodigoVeiculoLegal As String, _
                         Optional ByVal plngCodigoGrupoVeiculoLegal As Long, _
                         Optional ByVal pblnIndicadorVigencia As String = "N", _
                         Optional ByVal penumTipoBackOffice As Integer = 0, _
                         Optional ByVal pstrSegregaDados As String = "S", _
                         Optional ByVal plngCodigoEmpresa As Long, _
                         Optional ByVal pstrSiglaSistema As String = vbNullString, _
                         Optional ByVal plngContaCustodia As Long = 0) As String
                         
Dim strSQL                                  As String
Dim strCondicao                             As String
    
On Error GoTo ErrorHandler

    strSQL = " SELECT    A.CO_VEIC_LEGA,                            " & vbNewLine & _
             "           A.SG_SIST,                                 " & vbNewLine & _
             "           A.CO_EMPR,                                 " & vbNewLine & _
             "           A.TP_BKOF,                                 " & vbNewLine & _
             "           A.CO_GRUP_VEIC_LEGA,                       " & vbNewLine & _
             "           C.NO_GRUP_VEIC_LEGA,                       " & vbNewLine & _
             "           A.NO_VEIC_LEGA,                            " & vbNewLine & _
             "           A.NO_REDU_VEIC_LEGA,                       " & vbNewLine & _
             "           A.CO_CNPJ_VEIC_LEGA,                       " & vbNewLine & _
             "           A.DT_INIC_VIGE,                            " & vbNewLine & _
             "           A.DT_FIM_VIGE,                             " & vbNewLine & _
             "           A.CO_USUA_ULTI_ATLZ,                       " & vbNewLine & _
             "           A.CO_ETCA_TRAB_ULTI_ATLZ,                  " & vbNewLine & _
             "           A.ID_PART_CAMR_CETIP,                      " & vbNewLine & _
             "           A.CO_CNTA_CUTD_PADR_SELIC,                 " & vbNewLine & _
             "           A.TP_TITL_BMA,                             " & vbNewLine & _
             "           A.DH_ULTI_ATLZ                             " & vbNewLine & _
             " FROM      A8.TB_VEIC_LEGA        A,                  " & vbNewLine & _
             "           A7.TB_SIST             B,                  " & vbNewLine & _
             "           A8.TB_GRUP_VEIC_LEGA   C                   " & vbNewLine & _
             " WHERE     A.SG_SIST = B.SG_SIST                      " & vbNewLine & _
             " AND       A.CO_EMPR = B.CO_EMPR                      " & vbNewLine & _
             " AND       A.CO_GRUP_VEIC_LEGA = C.CO_GRUP_VEIC_LEGA  " & vbNewLine

    If pstrCodigoVeiculoLegal <> vbNullString Then
        strCondicao = " AND A.CO_VEIC_LEGA = '" & pstrCodigoVeiculoLegal & "' "
    End If
    
    If pstrSiglaSistema <> vbNullString Then
        strCondicao = strCondicao & " AND A.SG_SIST = '" & pstrSiglaSistema & "' "
    End If
    
    If plngCodigoEmpresa <> 0 Then
        strCondicao = strCondicao & " AND A.CO_EMPR = " & plngCodigoEmpresa
    End If
    
    If plngCodigoGrupoVeiculoLegal <> 0 Then
        strCondicao = strCondicao & " AND A.CO_GRUP_VEIC_LEGA = " & plngCodigoGrupoVeiculoLegal
    End If
    
    If penumTipoBackOffice <> 0 Then
        strCondicao = strCondicao & " AND A.TP_BKOF   = " & penumTipoBackOffice
    End If

    If plngContaCustodia <> 0 Then
        strCondicao = strCondicao & " AND A.CO_CNTA_CUTD_PADR_SELIC   = " & plngContaCustodia
    End If

    If UCase(pblnIndicadorVigencia) = "S" Then
        strCondicao = strCondicao & " AND   A.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                                    " AND  (A.DT_FIM_VIGE    IS  NULL       " & _
                                    " OR    A.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
    End If
    
    If pstrSegregaDados = "S" And penumTipoBackOffice = 0 Then
        strCondicao = strCondicao & fgSegregaDados("A8.TB_VEIC_LEGA", _
                                                   False, _
                                                   "A", _
                                                   "A", _
                                                   False, _
                                                   False, _
                                                   True, _
                                                   False, _
                                                   True, _
                                                   False)
    End If
    
    strSQL = strSQL & strCondicao
    strSQL = strSQL & " Order By  A.CO_GRUP_VEIC_LEGA ASC, A.NO_VEIC_LEGA ASC"
    
    LerTodos = fgQueryXMLLerTodos("VeiculoLegal", strSQL, App.EXEName & "." & TypeName(Me))

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

'Recuperar um registro da tabela A8.TB_VEIC_LEGA

Public Function Ler(ByVal pstrCodigoVeiculoLegal As String, _
                    ByVal pstrSiglaSistema As String, _
           Optional ByVal pblnValidaRemessa As Boolean = False) As String
                         
Dim strSQL                                   As String
    
    On Error GoTo ErrorHandler

    strSQL = " Select    A.CO_VEIC_LEGA,                           " & vbNewLine & _
             "           A.SG_SIST,                                " & vbNewLine & _
             "           A.CO_EMPR,                                " & vbNewLine & _
             "           A.TP_BKOF,                                " & vbNewLine & _
             "           A.CO_GRUP_VEIC_LEGA,                      " & vbNewLine & _
             "           A.NO_VEIC_LEGA,                           " & vbNewLine & _
             "           A.NO_REDU_VEIC_LEGA,                      " & vbNewLine & _
             "           A.CO_CNPJ_VEIC_LEGA,                      " & vbNewLine & _
             "           A.DT_INIC_VIGE,                           " & vbNewLine & _
             "           A.DT_FIM_VIGE,                            " & vbNewLine & _
             "           A.CO_USUA_ULTI_ATLZ,                      " & vbNewLine & _
             "           A.CO_ETCA_TRAB_ULTI_ATLZ,                 " & vbNewLine & _
             "           A.DH_ULTI_ATLZ,                           " & vbNewLine & _
             "           A.ID_PART_CAMR_CETIP,                     " & vbNewLine & _
             "           A.CO_CNTA_CUTD_PADR_SELIC,                " & vbNewLine & _
             "           A.TP_TITL_BMA,                            " & vbNewLine & _
             "           A.CO_TITL_BMA,                            " & vbNewLine & _
             "           B.NO_EMPR,                                " & vbNewLine & _
             "           C.NO_SIST,                                " & vbNewLine & _
             "           D.NO_GRUP_VEIC_LEGA                       " & vbNewLine

    strSQL = strSQL & _
             " From      A8.TB_VEIC_LEGA  A,                       " & vbNewLine & _
             "           A8.TB_EMPRESA_HO    B,                    " & vbNewLine & _
             "           A7.TB_SIST       C,                       " & vbNewLine & _
             "           A8.TB_GRUP_VEIC_LEGA D                    " & vbNewLine & _
             " Where     D.CO_GRUP_VEIC_LEGA = A.CO_GRUP_VEIC_LEGA " & vbNewLine & _
             " AND       B.CO_EMPR = A.CO_EMPR                     " & vbNewLine & _
             " AND       C.SG_SIST = A.SG_SIST                     " & vbNewLine & _
             " AND       C.CO_EMPR = A.CO_EMPR                     " & vbNewLine & _
             " AND       A.CO_VEIC_LEGA = '" & pstrCodigoVeiculoLegal & "' "
            
    If pstrSiglaSistema <> "" Then
       strSQL = strSQL & " AND A.SG_SIST = '" & pstrSiglaSistema & "'"
    End If
    
    Ler = fgQueryXMLLer("VeiculoLegal", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Len(Ler) = 0 Then
        If Not pblnValidaRemessa Then
            '4002 - Registro Inexistente ou excluído por outro usuário.
            lngCodigoErroNegocio = 4002
            GoTo ErrorHandler
        End If
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Recuperar dados de um veiculo legal

Public Function ObterDetalheVeiculoLegal(ByVal pstrFiltro As String) As String


Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strSQL                                  As String
Dim strWhereFiltro                          As String

Dim strValores                              As String
Dim intTipoBackOffice                       As Integer

    On Error GoTo ErrorHandler
    
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    xmlDomFiltro.loadXML pstrFiltro
    
    'Filtro por Empresa
    strValores = vbNullString
    For Each objDomNode In xmlDomFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_EMPR  IN (" & strValores & ")" & vbNewLine
    End If
    
    'Filtro por Sistema
    strValores = vbNullString
    For Each objDomNode In xmlDomFiltro.selectNodes("Repeat_Filtros/Grupo_Sistema/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.SG_SIST  IN (" & strValores & ")" & vbNewLine
    End If
    
    'Filtro por Grupo de Veículo Legal
    strValores = vbNullString
    For Each objDomNode In xmlDomFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_GRUP_VEIC_LEGA  IN (" & strValores & ")" & vbNewLine
    End If
     
    'Filtro por Veículo Legal
    strValores = vbNullString
    For Each objDomNode In xmlDomFiltro.selectNodes("Repeat_Filtros/Grupo_VeiculoLegal/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_VEIC_LEGA  IN (" & strValores & ")" & vbNewLine
    End If
     
    'Filtro por Código CNPJ
    strValores = vbNullString
    For Each objDomNode In xmlDomFiltro.selectNodes("Repeat_Filtros/Grupo_CodigoCNPJ/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_CNPJ_VEIC_LEGA  IN (" & strValores & ")" & vbNewLine
    End If
     
    'Filtro por Identificador de participante CETIP
    strValores = vbNullString
    For Each objDomNode In xmlDomFiltro.selectNodes("Repeat_Filtros/Grupo_IdParticipacaoCETIP/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.ID_PART_CAMR_CETIP  IN (" & strValores & ")" & vbNewLine
    End If

    'Filtro por Conta Padrão SELIC
    strValores = vbNullString
    For Each objDomNode In xmlDomFiltro.selectNodes("Repeat_Filtros/Grupo_ContaPadraoSELIC/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_CNTA_CUTD_PADR_SELIC  IN (" & strValores & ")" & vbNewLine
    End If
     
    'Filtro por Tipo de Título BMA
    strValores = vbNullString
    For Each objDomNode In xmlDomFiltro.selectNodes("Repeat_Filtros/Grupo_TipoTituloBMA/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_TITL_BMA  IN ('" & strValores & "')" & vbNewLine
    End If
    
    intTipoBackOffice = 99
    For Each objDomNode In xmlDomFiltro.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
        intTipoBackOffice = Val(objDomNode.Text)
    Next
    
    Select Case intTipoBackOffice
        Case 99
            strWhereFiltro = strWhereFiltro & _
                             fgSegregaDados("A8.TB_VEIC_LEGA", False, "A", "A", False, False, _
                             True, False, False, False, , True)
        Case Is > 0
            strWhereFiltro = strWhereFiltro & vbNewLine & _
                             " AND A.TP_BKOF = " & intTipoBackOffice & vbNewLine
    End Select
    
    Set xmlDomFiltro = Nothing
    
    strSQL = " SELECT A.CO_VEIC_LEGA,                          " & vbNewLine & _
             "        A.SG_SIST,                               " & vbNewLine & _
             "        A.CO_EMPR,                               " & vbNewLine & _
             "        A.CO_GRUP_VEIC_LEGA,                     " & vbNewLine & _
             "        A.TP_BKOF,                               " & vbNewLine & _
             "        A.NO_VEIC_LEGA,                          " & vbNewLine & _
             "        A.NO_REDU_VEIC_LEGA,                     " & vbNewLine & _
             "        A.CO_CNPJ_VEIC_LEGA,                     " & vbNewLine & _
             "        A.DT_INIC_VIGE,                          " & vbNewLine & _
             "        A.DT_FIM_VIGE,                           " & vbNewLine & _
             "        A.CO_USUA_ULTI_ATLZ,                     " & vbNewLine & _
             "        A.CO_ETCA_TRAB_ULTI_ATLZ,                " & vbNewLine & _
             "        A.DH_ULTI_ATLZ,                          " & vbNewLine & _
             "        A.ID_PART_CAMR_CETIP,                    " & vbNewLine & _
             "        A.CO_CNTA_CUTD_PADR_SELIC,               " & vbNewLine & _
             "        A.TP_TITL_BMA,                           " & vbNewLine & _
             "        A.CO_TITL_BMA,                           " & vbNewLine & _
             "        B.NO_EMPR,                               " & vbNewLine & _
             "        C.NO_GRUP_VEIC_LEGA,                     " & vbNewLine & _
             "        D.DE_BKOF                                " & vbNewLine
    
    strSQL = strSQL & _
             " FROM   A8.TB_VEIC_LEGA        A,                " & vbNewLine & _
             "        A8.TB_EMPRESA_HO       B,                " & vbNewLine & _
             "        A8.TB_GRUP_VEIC_LEGA   C,                " & vbNewLine & _
             "        A8.TB_TIPO_BKOF        D                 " & vbNewLine
             
    strSQL = strSQL & _
             " WHERE  A.CO_EMPR = B.CO_EMPR                     " & vbNewLine & _
             " AND    A.CO_GRUP_VEIC_LEGA = C.CO_GRUP_VEIC_LEGA " & vbNewLine & _
             " AND    A.TP_BKOF = D.TP_BKOF                     " & vbNewLine
             
    strSQL = strSQL & strWhereFiltro
    
    ObterDetalheVeiculoLegal = fgQueryXMLLerTodos("Repeat_VeiculoLegal", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set xmlDomFiltro = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar os veiculos legais de um um grupo de veiculo Legal

Public Function ObterRelacaoVeicLegalPorGrupoVeicLegal(ByVal pstrXMLFiltros As String, _
                                              Optional ByVal penumTipoBackOffice As enumTipoBackOffice) As String

Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")

    strSQL = " SELECT    A.CO_GRUP_VEIC_LEGA,                           " & vbNewLine & _
             "           B.NO_GRUP_VEIC_LEGA,                           " & vbNewLine & _
             "           A.CO_VEIC_LEGA,                                " & vbNewLine & _
             "           A.SG_SIST,                                     " & vbNewLine & _
             "           A.NO_VEIC_LEGA                                 " & vbNewLine & _
             " FROM      A8.TB_VEIC_LEGA        A,                      " & vbNewLine & _
             "           A8.TB_GRUP_VEIC_LEGA   B                       " & vbNewLine & _
             " WHERE     B.CO_GRUP_VEIC_LEGA    = A.CO_GRUP_VEIC_LEGA   "
             
    If penumTipoBackOffice <> 0 And _
       penumTipoBackOffice = enumTipoBackOffice.Todos Then
       strSQL = strSQL & vbNewLine & _
                " AND A.TP_BKOF = " & penumTipoBackOffice
    End If

    If pstrXMLFiltros <> "" Then
       Call xmlDomFiltros.loadXML(pstrXMLFiltros)
        
       If Not xmlDomFiltros.selectSingleNode("//GrupoVeiculoLegal") Is Nothing Then
           strSQL = strSQL & vbNewLine & _
                    " AND A.CO_GRUP_VEIC_LEGA = " & xmlDomFiltros.selectSingleNode("//GrupoVeiculoLegal").Text
       End If
    
       If Not xmlDomFiltros.selectSingleNode("//VeiculoLegal") Is Nothing Then
          strSQL = strSQL & vbNewLine & _
                   " AND A.CO_VEIC_LEGA = '" & xmlDomFiltros.selectSingleNode("//VeiculoLegal").Text & "' "
       End If
    
       If Not xmlDomFiltros.selectSingleNode("//BancoLiquidante") Is Nothing Then
          strSQL = strSQL & vbNewLine & _
                   " AND A.CO_EMPR = " & xmlDomFiltros.selectSingleNode("//BancoLiquidante").Text
       End If
    
    End If
    
    strSQL = strSQL & fgSegregaDados("A8.TB_VEIC_LEGA", _
                                     False, _
                                     "A", _
                                     "A", _
                                     False, _
                                     False, _
                                     penumTipoBackOffice = 0, _
                                     False, _
                                     True, _
                                     False)
    
    strSQL = strSQL & " ORDER BY A.CO_GRUP_VEIC_LEGA, B.NO_GRUP_VEIC_LEGA, A.CO_VEIC_LEGA, A.SG_SIST, A.NO_VEIC_LEGA"
    
    ObterRelacaoVeicLegalPorGrupoVeicLegal = fgQueryXMLLerTodos("RelacaoVeicLegal", strSQL, App.EXEName & "." & TypeName(Me))

    Set xmlDomFiltros = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set xmlDomFiltros = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterRelacaoVeicLegalPorGrupoVeicLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o codigo de grupo de veiculo legal de um veiculo legal

Public Function ObterGrupoVeiculoLegal(ByVal pstrCodigoVeiculoLegal As String, _
                              Optional ByVal pstrSiglaSistema As String = vbNullString) As Long

Dim strSQL                                  As String
Dim rsGrupoVeicLegal                        As ADODB.Recordset

    On Error GoTo ErrorHandler

    Set rsGrupoVeicLegal = CreateObject("ADODB.Recordset")

    strSQL = " SELECT    CO_GRUP_VEIC_LEGA   " & vbNewLine & _
             " FROM      A8.TB_VEIC_LEGA     " & vbNewLine & _
             " WHERE     CO_VEIC_LEGA   =   '" & pstrCodigoVeiculoLegal & "'"

    If pstrSiglaSistema <> vbNullString Then
        strSQL = strSQL & _
             " AND       SG_SIST        =   '" & pstrSiglaSistema & "'"
    End If
    
    Set rsGrupoVeicLegal = fgQuerySQL(strSQL)

    If rsGrupoVeicLegal.EOF Then
        '4022 - Veículo Legal inválido.
        lngCodigoErroNegocio = 4022
        GoTo ErrorHandler
    End If

    ObterGrupoVeiculoLegal = rsGrupoVeicLegal!CO_GRUP_VEIC_LEGA

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterGrupoVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter o tipo de BackOffice de um veiculo legal

Public Function ObterTipoBackOffice(ByVal pstrCodigoVeiculoLegal As String, _
                           Optional ByVal pstrSiglaSistema As String = vbNullString) As Long

Dim strSQL                                  As String
Dim rsTipoBackOffice                        As ADODB.Recordset

    On Error GoTo ErrorHandler

    Set rsTipoBackOffice = CreateObject("ADODB.Recordset")

    strSQL = " SELECT    TP_BKOF             " & vbNewLine & _
             " FROM      A8.TB_VEIC_LEGA     " & vbNewLine & _
             " WHERE     CO_VEIC_LEGA   =   '" & pstrCodigoVeiculoLegal & "'"

    If pstrSiglaSistema <> vbNullString Then
        strSQL = strSQL & _
             " AND       SG_SIST        =   '" & pstrSiglaSistema & "'"
    End If
    
    Set rsTipoBackOffice = fgQuerySQL(strSQL)

    If rsTipoBackOffice.EOF Then
        '4022 - Veículo Legal inválido.
        lngCodigoErroNegocio = 4022
        GoTo ErrorHandler
    End If

    ObterTipoBackOffice = rsTipoBackOffice!TP_BKOF

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoBackOffice Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter o veiculo legal de acordo com o CNPJ / Conta Custodia

Public Function ObterCodigoVeiculoLegal(ByVal pstrCodigoMensagem As String, _
                                        ByVal pvntCNPJ_IDENT_PART As Variant, _
                                        ByVal plngCodigoEmpresa As Long, _
                                        ByRef pstrCodigoVeiculoLegal As String, _
                                        ByRef pstrSiglaSistema As String, _
                                        ByRef plngTipoBackoffice As Long, _
                               Optional ByRef pstrNomeVeicLega As String, _
                               Optional ByRef pvntIDENT_PART_CAMR As Variant = 0, _
                               Optional ByVal pblnCliente1 As Boolean = False) As Boolean

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

Dim vntIdentdCtrapartCamr                   As Variant
Dim strIdentdCtrapartCamr                   As String
Dim xmlIdentdCtrapartCamr                   As MSXML2.DOMDocument40

Dim lngLocalLiquidacao                      As Long

    On Error GoTo ErrorHandler
        
    pstrCodigoVeiculoLegal = vbNullString
    pstrSiglaSistema = vbNullString
    pstrCodigoMensagem = Trim(pstrCodigoMensagem)
        
    lngLocalLiquidacao = enumLocalLiquidacao.CETIP
    
    strSQL = " SELECT   CO_VEIC_LEGA,       " & vbNewLine & _
             "          SG_SIST,            " & vbNewLine & _
             "          TP_BKOF,            " & vbNewLine & _
             "          NO_VEIC_LEGA,       " & vbNewLine & _
             "          ID_PART_CAMR_CETIP  " & vbNewLine & _
             " FROM     A8.TB_VEIC_LEGA     " & vbNewLine & _
             " WHERE    CO_EMPR =           " & plngCodigoEmpresa & _
             " AND      DT_FIM_VIGE IS NULL " & vbNewLine

    Select Case pstrCodigoMensagem
        Case "CTP1002"
            strSQL = strSQL & "    AND   ID_PART_CAMR_CETIP = " & pvntCNPJ_IDENT_PART

        Case "BMA0015"
            strSQL = strSQL & "    AND   SUBSTR(LPAD(CO_CNTA_CUTD_PADR_SELIC,9,0),1,4) =  '" & pvntCNPJ_IDENT_PART & "'"
            
            lngLocalLiquidacao = enumLocalLiquidacao.BMA

        Case "BMC0112"
            strSQL = strSQL & "    AND   CO_CNPJ_VEIC_LEGA = " & Trim(pvntCNPJ_IDENT_PART)
            strSQL = strSQL & "    AND   SG_SIST           = 'E2'"

            lngLocalLiquidacao = enumLocalLiquidacao.BMC
        
        Case "SEL1611"
            If pblnCliente1 Then
                strSQL = strSQL & "    AND   SUBSTR(LPAD(CO_CNTA_CUTD_PADR_SELIC,9,0),1,8)  =  '" & pvntCNPJ_IDENT_PART & "'"
            Else
                strSQL = strSQL & "    AND   CO_CNTA_CUTD_PADR_SELIC                        =  '" & Val(pvntCNPJ_IDENT_PART) & "'"
            End If

            lngLocalLiquidacao = enumLocalLiquidacao.SELIC
        
        Case "LTR0001", "LTR0005R2", "LTR0007"
            pvntIDENT_PART_CAMR = Val(pvntIDENT_PART_CAMR)
            strSQL = strSQL & "    AND   ID_PART_CAMR_CETIP = " & pvntIDENT_PART_CAMR

        Case "CCR0006", "CCR0007"
            strSQL = strSQL & "    AND   TP_BKOF = " & enumTipoBackOffice.Comex

        Case "BMC0005", "BMC0015"
            strSQL = strSQL & "    AND   CO_CNPJ_VEIC_LEGA = " & Trim(pvntCNPJ_IDENT_PART)
            strSQL = strSQL & "    AND   TP_BKOF           = " & plngTipoBackoffice

        Case Else
            strSQL = strSQL & "    AND   CO_CNPJ_VEIC_LEGA = " & Trim(pvntCNPJ_IDENT_PART)

    End Select

    strSQL = strSQL & " ORDER BY SG_SIST DESC"

    Set objRS = fgQuerySQL(strSQL)

    If Not objRS.EOF Then
        If pstrCodigoMensagem = "SEL1611" Then
            If objRS.RecordCount > 1 Then
                pstrCodigoVeiculoLegal = vbNullString
                pstrNomeVeicLega = vbNullString
                pstrSiglaSistema = vbNullString
                'Obter o tipo de backoffice qdo acha n veiculo legal
                flObterTipoBackOfficeSEL1611 objRS, plngTipoBackoffice
            Else
                pstrCodigoVeiculoLegal = objRS.Fields("CO_VEIC_LEGA")
                pstrNomeVeicLega = objRS.Fields("NO_VEIC_LEGA")
                pstrSiglaSistema = objRS.Fields("SG_SIST")
                plngTipoBackoffice = objRS.Fields("TP_BKOF")
                
                If pvntIDENT_PART_CAMR <> 0 Then
                    pvntIDENT_PART_CAMR = IIf(IsNull(objRS.Fields("ID_PART_CAMR_CETIP")), 0, objRS.Fields("ID_PART_CAMR_CETIP"))
                End If
            End If
        Else
            pstrCodigoVeiculoLegal = objRS.Fields("CO_VEIC_LEGA")
            pstrNomeVeicLega = objRS.Fields("NO_VEIC_LEGA")
            pstrSiglaSistema = objRS.Fields("SG_SIST")
            plngTipoBackoffice = objRS.Fields("TP_BKOF")
        End If

        ObterCodigoVeiculoLegal = True
    
    Else
        pstrCodigoVeiculoLegal = vbNullString
        pstrSiglaSistema = vbNullString
        plngTipoBackoffice = 0
        pstrNomeVeicLega = vbNullString
        ObterCodigoVeiculoLegal = False
    
    End If

    If Not ObterCodigoVeiculoLegal Then
        Select Case pstrCodigoMensagem
            Case "LTR0001", "LTR0005R2", "LTR0007"
                strIdentdCtrapartCamr = LerTodosIDPartcamara(pvntIDENT_PART_CAMR, _
                                                             "", _
                                                             "", _
                                                             lngLocalLiquidacao)
            Case Else
                strIdentdCtrapartCamr = LerTodosIDPartcamara(pvntCNPJ_IDENT_PART, _
                                                             "", _
                                                             "", _
                                                             lngLocalLiquidacao)
        End Select
        
        If strIdentdCtrapartCamr = vbNullString Then
            ObterCodigoVeiculoLegal = False
        
            Select Case pstrCodigoMensagem
                Case "LTR0001", "LTR0005R2", "LTR0007"
                    
                    pvntCNPJ_IDENT_PART = Val(pvntCNPJ_IDENT_PART)
                    
                    strSQL = " SELECT   CO_VEIC_LEGA,       " & vbNewLine & _
                             "          SG_SIST,            " & vbNewLine & _
                             "          TP_BKOF,            " & vbNewLine & _
                             "          NO_VEIC_LEGA,       " & vbNewLine & _
                             "          ID_PART_CAMR_CETIP  " & vbNewLine & _
                             " FROM     A8.TB_VEIC_LEGA     " & vbNewLine & _
                             " WHERE    CO_EMPR           = " & plngCodigoEmpresa & vbNewLine & _
                             " AND      CO_CNPJ_VEIC_LEGA = " & pvntCNPJ_IDENT_PART
        
                    Set objRS = fgQuerySQL(strSQL)
                
                    If Not objRS.EOF Then
                        pstrCodigoVeiculoLegal = objRS.Fields("CO_VEIC_LEGA")
                        pstrNomeVeicLega = objRS.Fields("NO_VEIC_LEGA")
                        pstrSiglaSistema = objRS.Fields("SG_SIST")
                        plngTipoBackoffice = objRS.Fields("TP_BKOF")
                        
                        ObterCodigoVeiculoLegal = True
                    End If
                
            End Select
            
        Else
            Set xmlIdentdCtrapartCamr = CreateObject("MSXML2.DOMDocument.4.0")
            
            xmlIdentdCtrapartCamr.loadXML strIdentdCtrapartCamr
            
            pstrCodigoVeiculoLegal = xmlIdentdCtrapartCamr.selectSingleNode("//CO_VEIC_LEGA").Text
            pstrSiglaSistema = xmlIdentdCtrapartCamr.selectSingleNode("//SG_SIST").Text
            
            strIdentdCtrapartCamr = Ler(pstrCodigoVeiculoLegal, pstrSiglaSistema)

            xmlIdentdCtrapartCamr.loadXML strIdentdCtrapartCamr
            
            plngTipoBackoffice = xmlIdentdCtrapartCamr.selectSingleNode("//TP_BKOF").Text
            pstrNomeVeicLega = xmlIdentdCtrapartCamr.selectSingleNode("//NO_VEIC_LEGA").Text
            
            Set xmlIdentdCtrapartCamr = Nothing
            
            ObterCodigoVeiculoLegal = True
        End If
        
    End If
    
    objRS.Close

    Exit Function

ErrorHandler:
    Set xmlIdentdCtrapartCamr = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterCodigoVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o tipo de backoffice para a mensagem SEL1611 por veiculo legal, caso tenha mais de um tipo de backoffice para o mesmo veiculo legal

Private Function flObterTipoBackOfficeSEL1611(ByRef prsVeiculoLegal As ADODB.Recordset, _
                                              ByRef plngTipoBackoffice As Long) As Boolean

Dim blnDiferente                            As Boolean

On Error GoTo ErrorHandler

    prsVeiculoLegal.MoveFirst
    plngTipoBackoffice = prsVeiculoLegal.Fields("TP_BKOF")
    blnDiferente = False

    Do While Not prsVeiculoLegal.EOF

        If plngTipoBackoffice <> prsVeiculoLegal.Fields("TP_BKOF") Then
            plngTipoBackoffice = 0
            blnDiferente = True
            Exit Do
        End If

        prsVeiculoLegal.MoveNext
    Loop

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterTipoBackOfficeSEL1611 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Obter Identificadores Participante Camara

Public Function LerTodosIDPartcamara(ByVal vntIdentificador As Variant, _
                            Optional ByVal pstrCodigoVeiculoLegal As String = vbNullString, _
                            Optional ByVal pstrSiglaSistema As String = vbNullString, _
                            Optional ByVal plngLocalLiquidacao As Long = 0) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler
    
    strSQL = " SELECT   CO_VEIC_LEGA,                   " & vbNewLine & _
             "          SG_SIST,                        " & vbNewLine & _
             "          CO_LOCA_LIQU,                   " & vbNewLine & _
             "          TP_IDEF_PART_CAMR,              " & vbNewLine & _
             "          ID_PART_CAMR                    " & vbNewLine & _
             " FROM     A8.TB_IDEF_PART_CAMR_VEIC_LEGA  " & vbNewLine & _
             " WHERE    ID_PART_CAMR =                 '" & vntIdentificador & "'" & vbNewLine
        
    If plngLocalLiquidacao <> 0 Then
        strSQL = strSQL & " AND CO_LOCA_LIQU = " & plngLocalLiquidacao
    End If

    If pstrCodigoVeiculoLegal <> vbNullString Then
        strSQL = strSQL & " AND CO_LOCA_LIQU  = '" & pstrCodigoVeiculoLegal & "'"
    End If

    If pstrSiglaSistema <> vbNullString Then
        strSQL = strSQL & " AND SG_SIST  = '" & pstrCodigoVeiculoLegal & "'"
    End If

    LerTodosIDPartcamara = fgQueryXMLLerTodos("IDPartcamaraVeiculoLegal", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosIDPartcamara Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Gerenciar as operações de Inclusão,Alteração e Exclusão
Public Function Salvar(ByVal pxmlVeiculoLegal As MSXML2.DOMDocument40) As Boolean
On Error GoTo ErrorHandler
    
    'flRegrasNegocio pxmlVeiculoLegal
            
    If pxmlVeiculoLegal.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pxmlVeiculoLegal
    ElseIf pxmlVeiculoLegal.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        flAlterar pxmlVeiculoLegal
    ElseIf pxmlVeiculoLegal.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        flExcluir pxmlVeiculoLegal
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

'Incluir um registro na tabela de tipo de operação

Private Function flIncluir(ByRef pxmlVeiculoLegal As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A8.TB_VEIC_LEGA (" & _
                "CO_VEIC_LEGA, " & _
                "SG_SIST, " & _
                "CO_EMPR, " & _
                "CO_GRUP_VEIC_LEGA, " & _
                "TP_BKOF, " & _
                "NO_VEIC_LEGA, " & _
                "NO_REDU_VEIC_LEGA, " & _
                "CO_CNPJ_VEIC_LEGA, " & _
                "DT_INIC_VIGE, " & _
                "DT_FIM_VIGE, " & _
                "CO_USUA_ULTI_ATLZ, " & _
                "CO_ETCA_TRAB_ULTI_ATLZ, " & _
                "DH_ULTI_ATLZ, " & _
                "ID_PART_CAMR_CETIP, " & _
                "CO_CNTA_CUTD_PADR_SELIC, " & _
                "TP_TITL_BMA, " & _
                "CO_TITL_BMA "
                
    With pxmlVeiculoLegal.documentElement
        strSQL = strSQL & " ) VALUES ( " & _
               "'" & .selectSingleNode("CO_VEIC_LEGA").Text & "', " & _
               "'" & .selectSingleNode("SG_SIST").Text & "', " & _
               .selectSingleNode("CO_EMPR").Text & "," & _
               .selectSingleNode("CO_GRUP_VEIC_LEGA").Text & "," & _
               .selectSingleNode("TP_BKOF").Text & "," & _
               "'" & .selectSingleNode("NO_VEIC_LEGA").Text & "', " & _
               "'" & .selectSingleNode("NO_REDU_VEIC_LEGA").Text & "', " & _
               IIf(Trim(.selectSingleNode("CO_CNPJ_VEIC_LEGA").Text) = "", " NULL ", .selectSingleNode("CO_CNPJ_VEIC_LEGA").Text) & ", " & _
               fgDtXML_To_Oracle(.selectSingleNode("DT_INIC_VIGE").Text) & ", " & _
               fgDtXML_To_Oracle(.selectSingleNode("DT_FIM_VIGE").Text) & ", " & _
               "'" & fgObterUsuarioRede & "'," & _
               "'" & fgObterEstacaoTrabalhoUsuario & "'," & _
               "SYSDATE," & _
               IIf(Trim(.selectSingleNode("ID_PART_CAMR_CETIP").Text) = 0, " NULL ", .selectSingleNode("ID_PART_CAMR_CETIP").Text) & ", " & _
               IIf(Trim(.selectSingleNode("CO_CNTA_CUTD_PADR_SELIC").Text) = 0, " NULL ", .selectSingleNode("CO_CNTA_CUTD_PADR_SELIC").Text) & ", " & _
               "'" & .selectSingleNode("TP_TITL_BMA").Text & "', " & _
               IIf(Trim(.selectSingleNode("CO_TITL_BMA").Text) = 0, " NULL ", .selectSingleNode("CO_TITL_BMA").Text) & ")"
    End With
    
    fgExecuteSQL strSQL
     
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Alterar um registro na tabela de tipo de operação
Private Function flAlterar(ByRef pxmlVeiculoLegal As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
On Error GoTo ErrorHandler

    With pxmlVeiculoLegal.documentElement
        strSQL = " UPDATE A8.TB_VEIC_LEGA " & _
                "  SET      CO_EMPR                 =" & .selectSingleNode("CO_EMPR").Text & ", " & _
                "           CO_GRUP_VEIC_LEGA       =" & .selectSingleNode("CO_GRUP_VEIC_LEGA").Text & ", " & _
                "           TP_BKOF                 =" & .selectSingleNode("TP_BKOF").Text & ", " & _
                "           NO_VEIC_LEGA            ='" & .selectSingleNode("NO_VEIC_LEGA").Text & "', " & _
                "           NO_REDU_VEIC_LEGA       ='" & .selectSingleNode("NO_REDU_VEIC_LEGA").Text & "', " & _
                "           CO_CNPJ_VEIC_LEGA       =" & IIf(Trim(.selectSingleNode("CO_CNPJ_VEIC_LEGA").Text) = "", " NULL ", .selectSingleNode("CO_CNPJ_VEIC_LEGA").Text) & ", " & _
                "           DT_INIC_VIGE            =" & fgDtXML_To_Oracle(.selectSingleNode("DT_INIC_VIGE").Text) & ", " & _
                "           DT_FIM_VIGE             =" & fgDtXML_To_Oracle(.selectSingleNode("DT_FIM_VIGE").Text) & ", " & _
                "           CO_USUA_ULTI_ATLZ       ='" & fgObterUsuarioRede & "', " & _
                "           CO_ETCA_TRAB_ULTI_ATLZ  ='" & fgObterEstacaoTrabalhoUsuario & "', " & _
                "           DH_ULTI_ATLZ            =SYSDATE, " & _
                "           ID_PART_CAMR_CETIP      =" & IIf(Val(.selectSingleNode("ID_PART_CAMR_CETIP").Text) = 0, "NULL", .selectSingleNode("ID_PART_CAMR_CETIP").Text) & ", " & _
                "           CO_CNTA_CUTD_PADR_SELIC =" & IIf(Val(.selectSingleNode("CO_CNTA_CUTD_PADR_SELIC").Text) = 0, "NULL", .selectSingleNode("CO_CNTA_CUTD_PADR_SELIC").Text) & ", " & _
                "           TP_TITL_BMA             ='" & .selectSingleNode("TP_TITL_BMA").Text & "', " & _
                "           CO_TITL_BMA             =" & IIf(Val(.selectSingleNode("CO_TITL_BMA").Text) = 0, "NULL", .selectSingleNode("CO_TITL_BMA").Text) & " " & _
                "  WHERE    CO_VEIC_LEGA            ='" & .selectSingleNode("CO_VEIC_LEGA").Text & "'" & _
                "  AND      SG_SIST                 ='" & .selectSingleNode("SG_SIST").Text & "'" & _
                "  AND      DH_ULTI_ATLZ            = " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text)
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

'Excluir um registro na tabela de tipo de operação

Private Function flExcluir(ByRef pxmlVeiculoLegal As MSXML2.DOMDocument40) As String

Dim strSQL              As String
Dim strSQL1             As String
Dim objRS               As ADODB.Recordset

On Error GoTo ErrorHandler

    With pxmlVeiculoLegal.documentElement
        strSQL1 = " SELECT ( " & _
                  " SELECT COUNT(CO_VEIC_LEGA) FROM A8.TB_OPER_ATIV " & _
                  " WHERE  CO_VEIC_LEGA  = '" & .selectSingleNode("CO_VEIC_LEGA").Text & "'" & _
                  " ) + ( " & _
                  " SELECT COUNT(CO_VEIC_LEGA) FROM A8HIST.TB_OPER_ATIV " & _
                  " WHERE  CO_VEIC_LEGA  = '" & .selectSingleNode("CO_VEIC_LEGA").Text & "'" & _
                  " ) AS QTDE FROM DUAL "

        Set objRS = fgQuerySQL(strSQL1)
    
        If objRS!QTDE > 0 Then
            '4277 - Existem operações vinculadas a esse Veículo Legal. Operação Cancelada!
            lngCodigoErroNegocio = 4277
            GoTo ErrorHandler
        Else
            strSQL = " DELETE FROM A8.TB_VEIC_LEGA " & _
                     " WHERE    CO_VEIC_LEGA  ='" & .selectSingleNode("CO_VEIC_LEGA").Text & "'" & _
                     " AND      SG_SIST       ='" & .selectSingleNode("SG_SIST").Text & "'"
        
            If fgExecuteSQL(strSQL) = 0 Then
                '4002-Registro inexistente ou excluído por outro usuário
                lngCodigoErroNegocio = 4002
                GoTo ErrorHandler
            End If
        End If
    End With

Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter os nomes da colunas da tabela de tipo de operação

Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT * FROM A8.TB_VEIC_LEGA "
    
    ObterPropriedades = fgPropriedades("Grupo_VeiculoLegal", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Alterar empresas de diversos veículos legais ao mesmo tempo
Public Function AlterarEmpresaVeiculoLegal(ByRef pxmlProcessamento As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim objXMLNode                              As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler

    strSQL = " UPDATE A8.TB_VEIC_LEGA " & _
             " SET    CO_EMPR      =  " & pxmlProcessamento.selectSingleNode("//CO_EMPR").Text & " " & vbNewLine & _
             " WHERE  "
    
    For Each objXMLNode In pxmlProcessamento.selectNodes("//Repeat_Processamento/*")
        strSQL = strSQL & _
             " OR    (CO_VEIC_LEGA = '" & objXMLNode.selectSingleNode("CO_VEIC_LEGA").Text & "' AND " & vbNewLine & _
             "        SG_SIST      = '" & objXMLNode.selectSingleNode("SG_SIST").Text & "')         " & vbNewLine
    Next
    
    strSQL = Replace$(strSQL, "WHERE   OR", "WHERE   ")
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarEmpresaVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
