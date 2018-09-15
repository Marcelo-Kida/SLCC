VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegraTransporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF7585E0346"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Objeto responsável pela persistência de regras de transporte e controle dos dados pertinentes a elas.
Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

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

'Validar mensagens consistindo dados como empresa, tipo de mensagem, sistema e regra de transporte.
Public Function AutenticarMensagem(ByRef xmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim strDetlaheOcorrenciaErro               As String

On Error GoTo ErrorHandler
    
    If Not flExisteTipoMensagem(xmlMensagem) Then
        strDetlaheOcorrenciaErro = "Tipo de Mensagem não cadastrado." & vbCrLf & _
                                   "Código do Tipo de Mensagem: " & xmlMensagem.documentElement.selectSingleNode("TP_MESG").Text
        
        xmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").appendChild fgCreateCDATASection(strDetlaheOcorrenciaErro)
        AutenticarMensagem = False
        Exit Function
    End If
    
    If Not IsNumeric(xmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text) Then
        strDetlaheOcorrenciaErro = "Código de Empresa inválido." & vbCrLf & _
                                   "Código da Empresa : " & xmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text
        
        xmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").appendChild fgCreateCDATASection(strDetlaheOcorrenciaErro)
        AutenticarMensagem = False
        Exit Function
    End If
    
    If Not flValidaEmpresa(xmlMensagem) Then
        strDetlaheOcorrenciaErro = "Código de Empresa inválido." & vbCrLf & _
                                   "Código da Empresa : " & xmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text
        
        xmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").appendChild fgCreateCDATASection(strDetlaheOcorrenciaErro)
        AutenticarMensagem = False
        Exit Function
    End If
    
    If Not flValidaSistemaOrigem(xmlMensagem) Then
        strDetlaheOcorrenciaErro = "Sistema Origem não cadastrado." & vbCrLf & _
                                   "Código da Empresa : " & xmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text & vbCrLf & _
                                   "Sigla Sistema     : " & xmlMensagem.documentElement.selectSingleNode("SG_SIST_ORIG").Text
        
        xmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").appendChild fgCreateCDATASection(strDetlaheOcorrenciaErro)
        AutenticarMensagem = False
        Exit Function
    End If
            
    If Not flValidaSistemaDestino(xmlMensagem) Then
        strDetlaheOcorrenciaErro = "Sistema não cadastrado." & vbCrLf & _
                                   "Código da Empresa : " & xmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text & vbCrLf & _
                                   "Sigla Sistema     : " & xmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text
        
        xmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").appendChild fgCreateCDATASection(strDetlaheOcorrenciaErro)
        AutenticarMensagem = False
        Exit Function
    End If
            
    If Not flExisteRegraTransporte(xmlMensagem) Then
        strDetlaheOcorrenciaErro = "Regra de Transporte não cadastrada." & vbCrLf & _
                                   "Código do Tipo de Mensagem  : " & xmlMensagem.documentElement.selectSingleNode("TP_MESG").Text & vbCrLf & _
                                   "Sigla Sistema               : " & xmlMensagem.documentElement.selectSingleNode("SG_SIST_ORIG").Text & vbCrLf & _
                                   "Código da Empresa           : " & xmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text

        xmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").appendChild fgCreateCDATASection(strDetlaheOcorrenciaErro)
        AutenticarMensagem = False
        Exit Function
    End If
    
'    If Not flExisteRegraSistemaDestino(CLng(xmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text), _
'                                       xmlMensagem.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
'                                       xmlMensagem.documentElement.selectSingleNode("TP_MESG").Text, _
'                                       xmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text) Then
'        strDetlaheOcorrenciaErro = "Regra de Transformação para o Sistema Destino não cadastrada." & vbCrLf & _
'                                   "Código do Tipo de Mensagem  : " & xmlMensagem.documentElement.selectSingleNode("TP_MESG").Text & vbCrLf & _
'                                   "Sigla Sistema Origem        : " & xmlMensagem.documentElement.selectSingleNode("SG_SIST_ORIG").Text & vbCrLf & _
'                                   "Sigla Sistema Destino       : " & xmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text & vbCrLf & _
'                                   "Código da Empresa           : " & xmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text
'
'        xmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").appendChild fgCreateCDATASection(strDetlaheOcorrenciaErro)
'        AutenticarMensagem = False
'        Exit Function
'    End If
    
    AutenticarMensagem = True

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AutenticarMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verificar a existência de regra de transporte para um determinado tipo de mensagem.
Public Function ExisteRegraTransporte(ByVal pstrTipoMensagem As String, _
                                      ByVal plngFormatoMensagemSaida As Long) As Boolean

Dim Rs                                      As Recordset
Dim strSQL                                  As String

On Error GoTo ErrorHandler

        strSQL = " SELECT  TP_MESG " & _
                 "   FROM  A7.TB_REGR_TRAP_MESG " & _
                 "  WHERE  TP_MESG      = '" & IIf(IsNumeric(pstrTipoMensagem), Val(pstrTipoMensagem), Trim(pstrTipoMensagem)) & "'" & _
                 "    AND  TP_FORM_MESG_SAID = " & plngFormatoMensagemSaida & _
                 "    AND  DT_FIM_VIGE_REGR_TRAP  IS NULL "

    Set Rs = fgQuerySQL(strSQL)
    
    ExisteRegraTransporte = Not (Rs.EOF)
    
    Rs.Close
    
    Set Rs = Nothing
    
    Exit Function
ErrorHandler:
    
    Set Rs = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExisteRegraTransporte Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verificar a existência de um determinado tipo de mensagem.
Private Function flExisteTipoMensagem(ByRef xmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim objTipoMensagem                               As A7Server.clsTipoMensagem

On Error GoTo ErrorHandler
    
    Set objTipoMensagem = CreateObject("A7Server.clsTipoMensagem")
    flExisteTipoMensagem = objTipoMensagem.ExisteTipoMensagem(xmlMensagem.selectSingleNode("//Grupo_Mensagem/TP_MESG").Text, _
                                                              CLng("0" & xmlMensagem.selectSingleNode("//Grupo_Mensagem/TP_FORM_MESG_SAID").Text))
    Set objTipoMensagem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objTipoMensagem = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExisteTipoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Função privada que verifica a existência de uma regra de transporte e retorna os dados relativos a ela.
Private Function flExisteRegraTransporte(ByRef xmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    With xmlMensagem.documentElement
    
        strSQL = " SELECT  A.TP_MESG, " & _
                 "         A.TP_FORM_MESG_SAID, " & _
                 "         A.DH_INIC_VIGE_REGR_TRAP " & _
                 "   FROM  A7.TB_REGR_TRAP_MESG A, A7.TB_REGR_SIST_DEST B " & _
                 "  WHERE  RTRIM(LTRIM(A.TP_MESG))  = '" & IIf(IsNumeric(.selectSingleNode("TP_MESG").Text), Val(.selectSingleNode("TP_MESG").Text), Trim(.selectSingleNode("TP_MESG").Text)) & "'" & _
                 "    AND  A.SG_SIST_ORIG = '" & .selectSingleNode("SG_SIST_ORIG").Text & "'" & _
                 "    AND  A.CO_EMPR_ORIG =  " & CLng(.selectSingleNode("CO_EMPR_ORIG").Text)
                 
        If Trim(.selectSingleNode("SG_SIST_DEST").Text) <> "" Then
            strSQL = strSQL & "    AND  B.SG_SIST_DEST = '" & Trim(.selectSingleNode("SG_SIST_DEST").Text) & "' "
        End If

        strSQL = strSQL & "    AND  A.TP_MESG = B.TP_MESG " & _
                 "    AND  A.SG_SIST_ORIG = B.SG_SIST_ORIG " & _
                 "    AND  A.CO_EMPR_ORIG = B.CO_EMPR_ORIG " & _
                 "    AND  A.TP_FORM_MESG_SAID = B.TP_FORM_MESG_SAID " & _
                 "    AND  A.DH_INIC_VIGE_REGR_TRAP = B.DH_INIC_VIGE_REGR_TRAP " & _
                 "    AND  A.DH_INIC_VIGE_REGR_TRAP  <= " & fgDataHoraServidor_To_Oracle & _
                 "    AND  (A.DT_FIM_VIGE_REGR_TRAP    IS NULL " & _
                 "     OR   A.DT_FIM_VIGE_REGR_TRAP >= " & fgDataHoraServidor_To_Oracle & ")"
        
        Set objRS = fgQuerySQL(strSQL)
        
        If objRS.EOF Then
            flExisteRegraTransporte = False
        Else
            flExisteRegraTransporte = True
            .selectSingleNode("TP_FORM_MESG_SAID").Text = objRS.fields("TP_FORM_MESG_SAID")
            .selectSingleNode("DH_INIC_VIGE_REGR_TRAP").Text = fgDtHr_To_Xml(objRS.fields("DH_INIC_VIGE_REGR_TRAP"))
        End If
    
    End With
    
    objRS.Close
     
    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExisteRegraTransporte Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verifica se já existe mensagem transitada para uma determinada regra de transporte.
Private Function flExisteMensagem(ByRef poObjRegra As MSXML2.DOMDocument40) As Boolean

Dim objMensagem                             As A7Server.clsMensagem

On Error GoTo ErrorHandler


    Set objMensagem = CreateObject("A7Server.clsMensagem")
    
    With poObjRegra.documentElement
        flExisteMensagem = objMensagem.ExisteMensagem(.selectSingleNode("//TP_MESG").Text, _
                                                      .selectSingleNode("//SG_SIST_ORIG").Text, _
                                                      .selectSingleNode("//CO_EMPR_ORIG").Text, _
                                                      .selectSingleNode("//DH_INIC_VIGE_REGR_TRAP").Text)
    End With
    
    Set objMensagem = Nothing
  
    Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExisteMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter todas as regras de transporte cadasreadas de acordo com o filtro informado.
Public Function LerTodos(Optional ByVal plngCodigoEmpresa As Long, _
                         Optional ByVal pstrSistema As String, _
                         Optional ByVal plngRegraAtiva As Long) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = "SELECT RT.TP_MESG,       " & vbCrLf & _
                   " RT.TP_FORM_MESG_SAID,  " & _
                   " RT.CO_TEXT_XML,   " & _
                   " TM.NO_TIPO_MESG,  " & vbCrLf & _
                   " TM.TP_NATZ_MESG,  " & vbCrLf & _
                   " RT.SG_SIST_ORIG,  " & vbCrLf & _
                   " SIST.NO_SIST,     " & vbCrLf & _
                   " RT.CO_EMPR_ORIG,  " & vbCrLf & _
                   " RT.DH_INIC_VIGE_REGR_TRAP,   " & vbCrLf & _
                   " RT.DT_FIM_VIGE_REGR_TRAP,    " & vbCrLf & _
                   " RT.TP_FORM_MESG_ENTR,        " & vbCrLf & _
                   " RT.DH_ULTI_ATLZ              " & vbCrLf & _
             "FROM A7.TB_TIPO_MESG TM,            " & vbCrLf & _
                  "A7.TB_SIST    SIST,            " & vbCrLf & _
                  "A7.TB_REGR_TRAP_MESG RT        " & vbCrLf & _
            "WHERE RT.TP_MESG = TM.TP_MESG        " & _
              "AND RT.TP_FORM_MESG_SAID = TM.TP_FORM_MESG_SAID " & _
              "AND RT.SG_SIST_ORIG = SIST.SG_SIST " & _
              "AND RT.CO_EMPR_ORIG = SIST.CO_EMPR " & _
              "AND RT.SG_SIST_ORIG = SIST.SG_SIST "

    If plngCodigoEmpresa <> 0 Then
        strSQL = strSQL & " AND RT.CO_EMPR_ORIG = " & plngCodigoEmpresa
    End If
                  
    If Trim$(pstrSistema) <> vbNullString Then
        strSQL = strSQL & " AND RT.SG_SIST_ORIG = '" & pstrSistema & "'"
    End If
    
    If plngRegraAtiva = 1 Then
        strSQL = strSQL & " AND RT.DT_FIM_VIGE_REGR_TRAP IS NULL "
    End If
    
    strSQL = strSQL & "ORDER BY RT.DH_INIC_VIGE_REGR_TRAP DESC"
    
    LerTodos = fgQueryXMLLerTodos("Regra", strSQL, App.EXEName & "." & TypeName(Me))
                   
    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter os atributos de uma determinada regra de transporte.
Public Function Ler(ByVal plngCodigoEmpresa As Long, _
                    ByVal pstrTipoMensagem As String, _
                    ByVal plngTipoFormatoMensagemSaida As Long, _
                    ByVal pstrSistemaOrigem As String, _
                    ByVal pstrDataHoraInicioVigencia As String) As String

Dim strSQL                                  As String
Dim objLer                                  As MSXML2.DOMDocument40
Dim strLer                                  As String

On Error GoTo ErrorHandler

    Set objLer = CreateObject("MSXML2.DOMDocument.4.0")

    Call fgAppendNode(objLer, "", "Grupo_RegraTransporteMensagem", "")

    'Informações da Regra de Tranformação
    strSQL = " SELECT  " & vbNewLine & _
             " A.TP_MESG, " & vbNewLine & _
             " A.TP_FORM_MESG_SAID, " & vbNewLine & _
             " A.SG_SIST_ORIG, " & vbNewLine & _
             " A.CO_EMPR_ORIG, " & vbNewLine & _
             " A.DH_INIC_VIGE_REGR_TRAP, " & vbNewLine & _
             " A.IN_EXIS_REGR_TRNF, " & vbNewLine & _
             " A.TP_FORM_MESG_ENTR, " & vbNewLine & _
             " A.CO_TEXT_XML, " & vbNewLine & _
             " A.TP_CTER_DELI, " & vbNewLine & _
             " A.DT_FIM_VIGE_REGR_TRAP, " & vbNewLine & _
             " A.CO_USUA_ULTI_ATLZ, " & vbNewLine & _
             " A.CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
             " A.DH_ULTI_ATLZ, " & vbNewLine & _
             " '' AS TX_REGR_TRNF_MESG, " & vbNewLine & _
             " '' AS IN_CONS " & vbNewLine & _
             "  FROM  A7.TB_REGR_TRAP_MESG A " & vbNewLine & _
             "  WHERE A.CO_EMPR_ORIG                        = " & plngCodigoEmpresa & vbNewLine & _
             "    AND RTRIM(LTRIM(UPPER(A.TP_MESG)))        = '" & Trim(UCase(pstrTipoMensagem)) & "'" & vbNewLine & _
             "    AND A.TP_FORM_MESG_SAID                   =  " & plngTipoFormatoMensagemSaida & vbNewLine & _
             "    AND RTRIM(LTRIM(UPPER(A.SG_SIST_ORIG)))   = '" & Trim(UCase(pstrSistemaOrigem)) & "'" & vbNewLine & _
             "    AND A.DH_INIC_VIGE_REGR_TRAP              =  " & fgDtHrXML_To_Oracle(pstrDataHoraInicioVigencia)

    strLer = fgQueryXMLLer("Grupo_RegraTransporte", strSQL, App.EXEName & "." & TypeName(Me))
        
    If Len(strLer) = 0 Then
        '4002-Registro inexistente ou excluído por outro usuário
        lngCodigoErroNegocio = 4002
        GoTo ErrorHandler
    End If
    
    If strLer <> vbNullString Then
        Call fgAppendXML(objLer, "Grupo_RegraTransporteMensagem", strLer)
        objLer.documentElement.selectSingleNode("Grupo_RegraTransporte/TX_REGR_TRNF_MESG").appendChild fgCreateCDATASection(fgSelectVarchar4000(objLer.documentElement.selectSingleNode("Grupo_RegraTransporte/CO_TEXT_XML").Text))
                
        If flRegraParaConsulta(plngCodigoEmpresa, _
                               pstrTipoMensagem, _
                               plngTipoFormatoMensagemSaida, _
                               pstrSistemaOrigem, _
                               pstrDataHoraInicioVigencia) Then
            objLer.documentElement.selectSingleNode("Grupo_RegraTransporte/IN_CONS").Text = "S"
        Else
            objLer.documentElement.selectSingleNode("Grupo_RegraTransporte/IN_CONS").Text = "N"
        End If
    End If
    
    'Informação dos Sistemas de Destino
    strSQL = " SELECT * " & vbNewLine & _
             "   FROM A7.TB_REGR_SIST_DEST " & vbNewLine & _
             "  WHERE CO_EMPR_ORIG                      = " & plngCodigoEmpresa & vbNewLine & _
             "    AND LTRIM(RTRIM(UPPER(TP_MESG)))      = '" & Trim(UCase(pstrTipoMensagem)) & "'" & vbNewLine & _
             "    AND TP_FORM_MESG_SAID                 = " & plngTipoFormatoMensagemSaida & vbNewLine & _
             "    AND LTRIM(RTRIM(UPPER(SG_SIST_ORIG))) = '" & Trim(UCase(pstrSistemaOrigem)) & "'" & vbNewLine & _
             "    AND DH_INIC_VIGE_REGR_TRAP            = " & fgDtHrXML_To_Oracle(pstrDataHoraInicioVigencia)
    
    strLer = fgQueryXMLLerTodos("SistemaDestino", strSQL, App.EXEName & "." & TypeName(Me))
    
    If strLer <> vbNullString Then
        Call fgAppendXML(objLer, "Grupo_RegraTransporteMensagem", strLer)
    Else
        Call fgAppendNode(objLer, "Grupo_RegraTransporteMensagem", "Repeat_SistemaDestino", "")
    End If
    
    Ler = objLer.xml
    
    Set objLer = Nothing

    Exit Function
ErrorHandler:
    Set objLer = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a estrutura da tabela TB_REGR_TRAP_MESG.
Public Function ObterPropriedades() As String

Dim strSQL                                  As String
Dim objPropriedade                          As MSXML2.DOMDocument40
Dim strPropriedade                           As String

On Error GoTo ErrorHandler

    Set objPropriedade = CreateObject("MSXML2.DOMDocument.4.0")
       
    Call fgAppendNode(objPropriedade, "", "Grupo_RegraTransporteMensagem", "")

    'Informações da Regra de Tranformação
    strSQL = " SELECT  " & _
             " TP_MESG, " & _
             " TP_FORM_MESG_SAID, " & _
             " SG_SIST_ORIG, " & _
             " CO_EMPR_ORIG, " & _
             " DH_INIC_VIGE_REGR_TRAP, " & _
             " IN_EXIS_REGR_TRNF, " & _
             " TP_FORM_MESG_ENTR, " & _
             " CO_TEXT_XML, " & _
             " TP_CTER_DELI, " & _
             " DT_FIM_VIGE_REGR_TRAP, " & _
             " CO_USUA_ULTI_ATLZ, " & _
             " CO_ETCA_TRAB_ULTI_ATLZ, " & _
             " DH_ULTI_ATLZ, " & _
             " '' AS TX_REGR_TRNF_MESG, " & _
             " '' AS IN_REGR_ATIV " & _
             "  FROM  A7.TB_REGR_TRAP_MESG "

    strPropriedade = fgPropriedades("Grupo_RegraTransporte", strSQL, "A7Server.clsRegraTransporte")

    Call fgAppendXML(objPropriedade, "Grupo_RegraTransporteMensagem", strPropriedade)
    
    'Informação dos Sistemas de Destino
    strSQL = " SELECT * " & _
             "   FROM A7.TB_REGR_SIST_DEST"
    
    strPropriedade = fgPropriedades("Grupo_SistemaDestino", strSQL, "A7Server.clsRegraTransporte")
    
    Call fgAppendNode(objPropriedade, "Grupo_RegraTransporteMensagem", "Repeat_SistemaDestino", "")
    
    Call fgAppendXML(objPropriedade, "Repeat_SistemaDestino", strPropriedade)
    
    ObterPropriedades = objPropriedade.xml
    
    Set objPropriedade = Nothing

    Exit Function
ErrorHandler:

    Set objPropriedade = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Controlar o roteamento e execução dos métodos privados, tais como:
' -  Incluir
' -  Alterar
' -  Excluir
' -  LerTodos
' -  Ler
Public Function Executar(ByRef pobjRegra As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pobjRegra.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar"
            Executar = Salvar(pobjRegra)
        
        Case "Ler"
            Executar = Ler(CLng(pobjRegra.selectSingleNode("//CO_EMPR_ORIG").Text), _
                           pobjRegra.selectSingleNode("//TP_MESG").Text, _
                           CLng(pobjRegra.selectSingleNode("//TP_FORM_MESG_SAID").Text), _
                           pobjRegra.selectSingleNode("//SG_SIST_ORIG").Text, _
                           pobjRegra.selectSingleNode("//DH_INIC_VIGE_REGR_TRAP").Text)
        
        Case "LerTodos"
            Executar = LerTodos(CLng(pobjRegra.selectSingleNode("//CO_EMPR_ORIG").Text), _
                                     pobjRegra.selectSingleNode("//SG_SIST_ORIG").Text, _
                                     pobjRegra.selectSingleNode("//IN_REGR_ATIV").Text)
        
        Case "Excluir"
            Executar = Excluir(pobjRegra)
            
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

'Promover a exclusão de uma regra de transporte cadastrada.
Public Function Excluir(ByRef pobjRegraTraducao As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    If flExisteMensagem(pobjRegraTraducao) Then
        'Erro: Regra esta sendo utilizazda
        lngCodigoErroNegocio = 1006
        GoTo ErrorHandler
    End If
    
    With pobjRegraTraducao.documentElement
        strSQL = "DELETE FROM A7.TB_REGR_SIST_DEST " & _
                 " WHERE TP_MESG                = '" & Trim(.selectSingleNode("//TP_MESG").Text) & "'" & _
                 "   AND TP_FORM_MESG_SAID      =  " & .selectSingleNode("//TP_FORM_MESG_SAID").Text & _
                 "   AND SG_SIST_ORIG           = '" & .selectSingleNode("//SG_SIST_ORIG").Text & "'" & _
                 "   AND CO_EMPR_ORIG           =  " & .selectSingleNode("//CO_EMPR_ORIG").Text & _
                 "   AND DH_INIC_VIGE_REGR_TRAP =  " & fgDtHrXML_To_Oracle(.selectSingleNode("//DH_INIC_VIGE_REGR_TRAP").Text)

        Call fgExecuteSQL(strSQL)

        strSQL = "DELETE FROM A7.TB_REGR_TRAP_MESG " & _
                 " WHERE TP_MESG                = '" & Trim(.selectSingleNode("//TP_MESG").Text) & "'" & _
                 "   AND TP_FORM_MESG_SAID      =  " & .selectSingleNode("//TP_FORM_MESG_SAID").Text & _
                 "   AND SG_SIST_ORIG           = '" & .selectSingleNode("//SG_SIST_ORIG").Text & "'" & _
                 "   AND CO_EMPR_ORIG           =  " & .selectSingleNode("//CO_EMPR_ORIG").Text & _
                 "   AND DH_INIC_VIGE_REGR_TRAP =  " & fgDtHrXML_To_Oracle(.selectSingleNode("//DH_INIC_VIGE_REGR_TRAP").Text) & _
                 "   AND DH_ULTI_ATLZ           =  " & fgDtHrXML_To_Oracle(.selectSingleNode("//DH_ULTI_ATLZ").Text)
        
        If fgExecuteSQL(strSQL) = 0 Then
           '4001 - Registro foi alterado ou excluído por outro usuário.
           lngCodigoErroNegocio = 4001
           GoTo ErrorHandler
        End If

    End With
    
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Excluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Rotear o comando salvar para as rotinas privadas flIncluir ou flAlterar.
Public Function Salvar(ByRef pobjRegraTraducao As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    If pobjRegraTraducao.selectSingleNode("//Grupo_RegraTransporte/@Operacao").Text = "Incluir" Then
        
        If Not flValidaSistema(pobjRegraTraducao) Then
            'Sistema de origem ou destino não está vigente
            lngCodigoErroNegocio = 1016
            GoTo ErrorHandler
        End If
        
        flIncluir pobjRegraTraducao
        flIncluirSistemasDestino pobjRegraTraducao
    
    ElseIf pobjRegraTraducao.selectSingleNode("//Grupo_RegraTransporte/@Operacao").Text = "Alterar" Then
        
'        If flExisteMensagem(pobjRegraTraducao) Then
'            If pobjRegraTraducao.selectSingleNode("//Grupo_RegraTransporte/DT_FIM_VIGE_REGR_TRAP").Text = "" Or _
'               pobjRegraTraducao.selectSingleNode("//Grupo_RegraTransporte/DT_FIM_VIGE_REGR_TRAP").Text = gstrDataVazia Then
'                lngCodigoErroNegocio = 1006
'                GoTo ErrorHandler
'            End If
'        End If
'
'        If Not flValidaSistema(pobjRegraTraducao) Then
'            'Sistema de origem ou destino não está vigente
'            lngCodigoErroNegocio = 1016
'            GoTo ErrorHandler
'        End If
            
        flAlterar pobjRegraTraducao
        flIncluirSistemasDestino pobjRegraTraducao
        
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

'Promover a inclusão de uma nova regra de transporte.
Private Function flIncluir(ByRef pobjRegraTraducao As MSXML2.DOMDocument40) As Boolean

Dim lngCodCntdAtrbXml                        As Long
Dim strSQL                                   As String

On Error GoTo ErrorHandler
            
    'Incluir Dados
    With pobjRegraTraducao.documentElement
        
        lngCodCntdAtrbXml = fgInsertVarchar4000(.selectSingleNode("//TX_REGR_TRNF_MESG").Text)
    
        strSQL = "INSERT INTO A7.TB_REGR_TRAP_MESG (" & _
                " TP_MESG, " & _
                " TP_FORM_MESG_SAID, " & _
                " SG_SIST_ORIG, " & _
                " CO_EMPR_ORIG, " & _
                " DH_INIC_VIGE_REGR_TRAP, " & _
                " IN_EXIS_REGR_TRNF, " & _
                " TP_FORM_MESG_ENTR, " & _
                " CO_TEXT_XML, " & _
                " TP_CTER_DELI, " & _
                " DT_FIM_VIGE_REGR_TRAP, " & _
                " CO_USUA_ULTI_ATLZ, " & _
                " CO_ETCA_TRAB_ULTI_ATLZ, " & _
                " DH_ULTI_ATLZ " & _
                " ) VALUES ("
        
        strSQL = strSQL & "'" & Trim(.selectSingleNode("//TP_MESG").Text) & "'," & _
                          .selectSingleNode("//TP_FORM_MESG_SAID").Text & "," & _
                          "'" & .selectSingleNode("//SG_SIST_ORIG").Text & "'," & _
                          .selectSingleNode("//CO_EMPR_ORIG").Text & "," & _
                          fgDtHrXML_To_Oracle(.selectSingleNode("//DH_INIC_VIGE_REGR_TRAP").Text) & "," & _
                          "'" & .selectSingleNode("//IN_EXIS_REGR_TRNF").Text & "'," & _
                          .selectSingleNode("//TP_FORM_MESG_ENTR").Text & "," & _
                          lngCodCntdAtrbXml & ", " & _
                          IIf(.selectSingleNode("//TP_CTER_DELI").Text = vbNullString, "NULL", "'" & .selectSingleNode("//TP_CTER_DELI").Text & "'") & "," & _
                          fgDtHrXML_To_Oracle(.selectSingleNode("//DT_FIM_VIGE_REGR_TRAP").Text) & "," & _
                          "'" & fgObterUsuarioRede & "'," & _
                          "'" & fgObterEstacaoTrabalhoUsuario & "'," & _
                          "SYSDATE)"
        
        fgExecuteSQL strSQL
            
    End With
      
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Promover a alteração de uma regra de transporte cadastrada.
Private Function flAlterar(ByRef pobjRegraTraducao As MSXML2.DOMDocument40) As Boolean

Dim lngCodCntdAtrbXml                        As Long
Dim strSQL                                   As String

On Error GoTo ErrorHandler
            
    'Incluir Dados
    With pobjRegraTraducao.documentElement
    
        lngCodCntdAtrbXml = CLng(.selectSingleNode("//CO_TEXT_XML").Text)
            
        strSQL = " DELETE FROM A7.TB_TEXT_XML " & _
                 "  WHERE CO_TEXT_XML = " & lngCodCntdAtrbXml
            
        fgExecuteSQL strSQL
            
        lngCodCntdAtrbXml = fgInsertVarchar4000(.selectSingleNode("//TX_REGR_TRNF_MESG").Text)
        
        strSQL = "UPDATE A7.TB_REGR_TRAP_MESG " & _
                 "   SET IN_EXIS_REGR_TRNF        = '" & .selectSingleNode("//IN_EXIS_REGR_TRNF").Text & "'," & vbCrLf & _
                        "TP_FORM_MESG_ENTR        = " & .selectSingleNode("//TP_FORM_MESG_ENTR").Text & "," & vbCrLf & _
                        "TP_CTER_DELI             = '" & .selectSingleNode("//TP_CTER_DELI").Text & "'," & vbCrLf & _
                        "CO_TEXT_XML              = " & lngCodCntdAtrbXml & "," & vbCrLf & _
                        "DT_FIM_VIGE_REGR_TRAP    = " & fgDtXML_To_Oracle(.selectSingleNode("//DT_FIM_VIGE_REGR_TRAP").Text) & "," & vbCrLf & _
                        "CO_USUA_ULTI_ATLZ        = '" & fgObterUsuarioRede & "'," & vbCrLf & _
                        "CO_ETCA_TRAB_ULTI_ATLZ   = '" & fgObterEstacaoTrabalhoUsuario & "'," & vbCrLf & _
                        "DH_ULTI_ATLZ             = SYSDATE" & vbCrLf & _
                 " WHERE TP_MESG                = '" & Trim(.selectSingleNode("//TP_MESG").Text) & "'" & vbCrLf & _
                 "   AND TP_FORM_MESG_SAID      = " & .selectSingleNode("//TP_FORM_MESG_SAID").Text & vbCrLf & _
                 "   AND SG_SIST_ORIG           = '" & .selectSingleNode("//SG_SIST_ORIG").Text & "'" & vbCrLf & _
                 "   AND CO_EMPR_ORIG           = " & .selectSingleNode("//CO_EMPR_ORIG").Text & vbCrLf & _
                 "   AND DH_INIC_VIGE_REGR_TRAP = " & fgDtHrXML_To_Oracle(.selectSingleNode("//DH_INIC_VIGE_REGR_TRAP").Text) & _
                 "   AND DH_ULTI_ATLZ           = " & fgDtHrXML_To_Oracle(.selectSingleNode("//DH_ULTI_ATLZ").Text)
        
        If fgExecuteSQL(strSQL) = 0 Then
           '4001 - Registro foi alterado ou excluído por outro usuário.
           lngCodigoErroNegocio = 4001
           GoTo ErrorHandler
        End If
        
    End With
      
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Incluir sistemas destinos para uma determinada regra de transporte.
Private Sub flIncluirSistemasDestino(ByRef pobjRegraTraducao As MSXML2.DOMDocument40)

Dim strSQL                                   As String

Dim objGrupoSistemaDestino                  As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    'Antes de salvar apagar todos os atributos deeste evento
    
    With pobjRegraTraducao
        strSQL = "DELETE FROM A7.TB_REGR_SIST_DEST " & _
                 " WHERE TP_MESG                = '" & Trim(.selectSingleNode("//Grupo_RegraTransporte/TP_MESG").Text) & "'" & vbCrLf & _
                 "   AND TP_FORM_MESG_SAID      =  " & .selectSingleNode("//Grupo_RegraTransporte/TP_FORM_MESG_SAID").Text & vbCrLf & _
                 "   AND SG_SIST_ORIG           = '" & .selectSingleNode("//Grupo_RegraTransporte/SG_SIST_ORIG").Text & "'" & vbCrLf & _
                 "   AND CO_EMPR_ORIG           =  " & .selectSingleNode("//Grupo_RegraTransporte/CO_EMPR_ORIG").Text & vbCrLf & _
                 "   AND DH_INIC_VIGE_REGR_TRAP =  " & fgDtHrXML_To_Oracle(.selectSingleNode("//Grupo_RegraTransporte/DH_INIC_VIGE_REGR_TRAP").Text) '& ")"
    End With
    
    fgExecuteSQL strSQL
    
    If pobjRegraTraducao.selectNodes("//Grupo_SistemaDestino").length = 0 Then
        Exit Sub
    End If
    
    'Inclui Sistemas Destino
    strSQL = "BEGIN" & vbCrLf
    For Each objGrupoSistemaDestino In pobjRegraTraducao.selectNodes("//Repeat_SistemaDestino/*")
        
        With objGrupoSistemaDestino
        
            'Verificar se já existe uma regra cadastrada para
            'o sistema de origem X  Sist destino
            If flExisteRegraSistemaDestino(CLng(.selectSingleNode("./CO_EMPR_ORIG").Text), _
                                           .selectSingleNode("./SG_SIST_ORIG").Text, _
                                           .selectSingleNode("./TP_MESG").Text, _
                                           .selectSingleNode("./SG_SIST_DEST").Text) Then
                lngCodigoErroNegocio = 1014
                GoTo ErrorHandler
            End If
            
            'Incluir XML do Formato
            strSQL = strSQL & " INSERT INTO A7.TB_REGR_SIST_DEST ( " & _
                              " TP_MESG,      " & _
                              " TP_FORM_MESG_SAID, " & _
                              " SG_SIST_ORIG, " & _
                              " CO_EMPR_ORIG, " & _
                              " DH_INIC_VIGE_REGR_TRAP, " & _
                              " SG_SIST_DEST, " & _
                              " CO_EMPR_DEST  " & _
                              " ) VALUES ( "
            strSQL = strSQL & "'" & Trim(.selectSingleNode("./TP_MESG").Text) & "'," & _
                              .selectSingleNode("./TP_FORM_MESG_SAID").Text & "," & _
                              "'" & .selectSingleNode("./SG_SIST_ORIG").Text & "'," & _
                              .selectSingleNode("./CO_EMPR_ORIG").Text & "," & _
                              fgDtHrXML_To_Oracle(.selectSingleNode("./DH_INIC_VIGE_REGR_TRAP").Text) & "," & _
                              "'" & .selectSingleNode("./SG_SIST_DEST").Text & "'," & _
                              .selectSingleNode("./CO_EMPR_DEST").Text & "); "
        End With
    
        If Len(strSQL) > 60000 Then
            strSQL = strSQL & "END;"
            fgExecuteSQL strSQL
            strSQL = "BEGIN" & vbCrLf
        End If
            
    Next
      
    If Len(strSQL) > Len("BEGIN" & vbCrLf) Then
        'Executa o Ultimo bloco
        strSQL = strSQL & "END;"
        fgExecuteSQL strSQL
    End If
      
    Exit Sub

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluirAtributos Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Sub

'Verifica se existe a associação entre uma regra de transporte e um sistema destino.
Private Function flExisteRegraSistemaDestino(plngCodigoEmpresa As Long, _
                                             pstrSistemaOrigem As String, _
                                             pstrTipoMensagem As String, _
                                             pstrSistemaDestino As String) As Boolean
'Verifica se já existe uma regra vigente para a relação Origem X Mesg X Destino
On Error GoTo ErrorHandler
Dim strSQL                                  As String
Dim Rs                                      As Recordset

    strSQL = "SELECT max(DH_INIC_VIGE_REGR_TRAP) as DH_INIC_VIGE_REGR_TRAP " & _
             "  FROM A7.TB_REGR_TRAP_MESG " & _
             " WHERE TP_MESG = '" & Trim(pstrTipoMensagem) & "'" & _
             "   AND CO_EMPR_ORIG = " & plngCodigoEmpresa & _
             "   AND SG_SIST_ORIG = '" & pstrSistemaOrigem & "'"

    Set Rs = fgQuerySQL(strSQL)
    
    If Not Rs.EOF Then

        strSQL = "SELECT 'X' " & _
                 "  FROM A7.TB_REGR_SIST_DEST " & _
                 " WHERE CO_EMPR_ORIG = " & plngCodigoEmpresa & _
                 "   AND TP_MESG = '" & Trim(pstrTipoMensagem) & "'" & _
                 "   AND SG_SIST_ORIG = '" & pstrSistemaOrigem & "'" & _
                 "   AND SG_SIST_DEST = '" & pstrSistemaDestino & "'" & _
                 "   AND DH_INIC_VIGE_REGR_TRAP = " & fgDtHrXML_To_Oracle(fgDtHr_To_Xml(Rs!DH_INIC_VIGE_REGR_TRAP))
        
        Set Rs = fgQuerySQL(strSQL)
            
        If Not Rs.EOF Then
            'Já existe ragra para este sistema destino
            'Nçao poermitir nova inclusão
            flExisteRegraSistemaDestino = True
        End If
    
    End If

    Set Rs = Nothing
    
    Exit Function
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExisteRegraSistemaDestino Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter as mensagens que ainda não contém regra de transporte cadastrada para uma empresa/sistema.
Public Function ObterEventosDisponiveis(ByVal plngCodigoEmpresa As Long, _
                                        ByVal pstrSistema As String) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT A.TP_MESG, " & vbCrLf & _
                     "A.NO_TIPO_MESG, " & vbCrLf & _
                     "A.TP_FORM_MESG_SAID, " & vbCrLf & _
                     "A.TP_NATZ_MESG  " & vbCrLf & _
             "   FROM A7.TB_TIPO_MESG A  " & vbCrLf & _
             "  WHERE A.DT_INIC_VIGE_MESG <= " & fgDataHoraServidor_To_Oracle & vbCrLf & _
             "    AND (A.DT_FIM_VIGE_MESG   IS  NULL" & vbCrLf & _
             "     OR  A.DT_FIM_VIGE_MESG  >= " & fgDataHoraServidor_To_Oracle & ")" & vbCrLf & _
             "    AND (A.TP_MESG, A.TP_FORM_MESG_SAID) NOT IN (SELECT TP_MESG, TP_FORM_MESG_SAID " & vbCrLf & _
                                     "    FROM A7.TB_REGR_TRAP_MESG " & vbCrLf & _
                                     "   WHERE CO_EMPR_ORIG = " & plngCodigoEmpresa & vbCrLf & _
                                     "     AND SG_SIST_ORIG = '" & pstrSistema & "'" & vbCrLf & _
                                     "     AND DT_FIM_VIGE_REGR_TRAP IS NULL )"
    
    ObterEventosDisponiveis = fgQueryXMLLerTodos("Evento", strSQL, App.EXEName & "." & TypeName(Me))
                   
    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterEventosDisponiveis Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
                             
'Obter os detalhes de uma regra de transporte cadastrada.
Public Function ObterInformacoesRegra(ByRef xmlMensagem As MSXML2.DOMDocument40) As String

Dim objRegra                                As MSXML2.DOMDocument40
Dim strSQL                                  As String
Dim Rs                                      As Recordset

On Error GoTo ErrorHandler

    With xmlMensagem.documentElement
        strSQL = " SELECT   TM.TP_MESG, " & vbCrLf & _
                 "          TM.TP_FORM_MESG_SAID AS TP_FORM_SAID_MESG," & vbCrLf & _
                 "          TM.CO_TEXT_XML       AS CO_TEXT_XML_MESG ," & vbCrLf & _
                 "          TM.TP_NATZ_MESG,     " & vbCrLf & _
                 "          TM.NO_TITU_MESG,     " & vbCrLf & _
                 "          RT.IN_EXIS_REGR_TRNF," & vbCrLf & _
                 "          RT.TP_FORM_MESG_ENTR AS TP_FORM_ENTR_MESG, " & vbCrLf & _
                 "          TM.TP_CTER_DELI      AS TP_CTER_DELI_SAID, " & vbCrLf & _
                 "          RT.TP_CTER_DELI      AS TP_CTER_DELI_ENTR, " & vbCrLf & _
                 "          RT.CO_TEXT_XML       AS CO_TEXT_XML_REGR   " & vbCrLf & _
                 "   FROM   A7.TB_TIPO_MESG TM, " & vbCrLf & _
                 "          A7.TB_REGR_TRAP_MESG RT " & vbCrLf & _
                 "  WHERE   TM.TP_MESG = RT.TP_MESG " & vbCrLf & _
                 "    AND   RTRIM(LTRIM(RT.TP_MESG)) = '" & IIf(IsNumeric(.selectSingleNode("//TP_MESG").Text), Val(.selectSingleNode("//TP_MESG").Text), Trim(.selectSingleNode("//TP_MESG").Text)) & "'" & vbCrLf & _
                 "    AND   RT.CO_EMPR_ORIG =  " & .selectSingleNode("CO_EMPR_ORIG").Text & vbCrLf & _
                 "    AND   RT.SG_SIST_ORIG = '" & .selectSingleNode("SG_SIST_ORIG").Text & "'" & vbCrLf & _
                 "    AND   RT.DH_INIC_VIGE_REGR_TRAP = " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_INIC_VIGE_REGR_TRAP").Text) & vbCrLf & _
                 "    AND   RT.TP_FORM_MESG_SAID = " & .selectSingleNode("TP_FORM_MESG_SAID").Text & vbCrLf & _
                 "    AND   RT.TP_FORM_MESG_SAID = TM.TP_FORM_MESG_SAID"
    End With
    
    Set Rs = fgQuerySQL(strSQL)

    If Not Rs.EOF Then
        
        Set objRegra = CreateObject("MSXML2.DOMDocument.4.0")
                
        fgAppendNode objRegra, vbNullString, "RegraTransporte", vbNullString
        fgAppendNode objRegra, "RegraTransporte", "TP_MESG", Trim(Rs!TP_MESG)
        fgAppendNode objRegra, "RegraTransporte", "TP_NATZ_MESG", Rs!TP_NATZ_MESG
        fgAppendNode objRegra, "RegraTransporte", "TP_FORM_SAID_MESG", Rs!TP_FORM_SAID_MESG
        fgAppendNode objRegra, "RegraTransporte", "TP_FORM_ENTR_MESG", Rs!TP_FORM_ENTR_MESG
        fgAppendNode objRegra, "RegraTransporte", "IN_EXIS_REGR_TRNF", Rs!IN_EXIS_REGR_TRNF
        fgAppendNode objRegra, "RegraTransporte", "TP_CTER_DELI_ENTR", IIf(IsNull(Rs!TP_CTER_DELI_ENTR), "", Rs!TP_CTER_DELI_ENTR)
        fgAppendNode objRegra, "RegraTransporte", "TP_CTER_DELI_SAID", IIf(IsNull(Rs!TP_CTER_DELI_SAID), "", Rs!TP_CTER_DELI_SAID)
        fgAppendNode objRegra, "RegraTransporte", "TX_REGR_TRNF_MESG", vbNullString
        fgAppendNode objRegra, "RegraTransporte", "NO_TITU_MESG", IIf(IsNull(Rs!NO_TITU_MESG), "", Rs!NO_TITU_MESG)
        
        objRegra.documentElement.selectSingleNode("TX_REGR_TRNF_MESG").appendChild fgCreateCDATASection(fgSelectVarchar4000(Rs!CO_TEXT_XML_REGR))
        
        fgAppendNode objRegra, "RegraTransporte", "TX_VALID_SAID_MESG", vbNullString
                
        If Not IsNull(Rs!CO_TEXT_XML_MESG) Then
            objRegra.documentElement.selectSingleNode("TX_VALID_SAID_MESG").appendChild fgCreateCDATASection(fgSelectVarchar4000(Rs!CO_TEXT_XML_MESG))
        End If
        
        ObterInformacoesRegra = objRegra.xml
        
        Set objRegra = Nothing
        
    End If
    
    Rs.Close
    
    Set Rs = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objRegra = Nothing
    Set Rs = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterInformacoesRegra Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Valida os sistemas destino de uma regra de transporte na sua inclusão.
Private Function flValidaSistema(ByRef xmlRegra As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim Rs                                      As ADODB.Recordset
Dim xmlNode                                 As IXMLDOMNode

On Error GoTo ErrorHandler

    For Each xmlNode In xmlRegra.selectNodes("//Grupo_RegraTransporte|//Grupo_SistemaDestino")
    
        strSQL = " SELECT   SG_SIST,    " & vbCrLf & _
                "           CO_EMPR     " & vbCrLf & _
                " FROM      A7.TB_SIST  " & vbCrLf & _
                " WHERE     CO_EMPR = " & CLng(xmlNode.selectSingleNode("CO_EMPR_ORIG").Text) & vbCrLf
        
        If xmlNode.selectSingleNode("SG_SIST_DEST") Is Nothing Then
            strSQL = strSQL & " AND       SG_SIST = '" & xmlNode.selectSingleNode("SG_SIST_ORIG").Text & "' " & vbCrLf
        Else
            strSQL = strSQL & " AND       SG_SIST = '" & xmlNode.selectSingleNode("SG_SIST_DEST").Text & "' " & vbCrLf
        End If
                
        strSQL = strSQL & " AND   DT_INIC_VIGE_SIST <= " & fgDataHoraServidor_To_Oracle & _
                          " AND  (DT_FIM_VIGE_SIST    IS  NULL       " & _
                          " OR    DT_FIM_VIGE_SIST  >= " & fgDataHoraServidor_To_Oracle & ")"
                
        Set Rs = fgQuerySQL(strSQL)
        
        If Rs.EOF Then
            flValidaSistema = False
            Exit For
        Else
            flValidaSistema = True
        End If

    Next
    
    Rs.Close
    Set Rs = Nothing

    Exit Function
    
ErrorHandler:
    Set Rs = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flValidaSistema Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Validar a empresa de uma regra de transporte.
Private Function flValidaEmpresa(ByRef xmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim objEmpresa                      As A6A7A8.clsEmpresa
Dim strEmpresa                      As String

On Error GoTo ErrorHandler
    
    Set objEmpresa = CreateObject("A6A7A8.clsEmpresa")
    
    strEmpresa = objEmpresa.Ler(CLng(xmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text))
    
    If Trim(strEmpresa) = vbNullString Then
        flValidaEmpresa = False
    Else
        flValidaEmpresa = True
    End If
    
    Set objEmpresa = Nothing

    Exit Function
    
ErrorHandler:
    
    Set objEmpresa = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flValidaEmpresa Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Validar o sistema de origem de uma regra de transporte.
Private Function flValidaSistemaOrigem(ByRef xmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim Rs                                      As ADODB.Recordset

On Error GoTo ErrorHandler
    
    If InStr(1, xmlMensagem.documentElement.selectSingleNode("SG_SIST_ORIG").Text, "'") > 0 Then
        flValidaSistemaOrigem = False
        Exit Function
    End If

    strSQL = " SELECT   SG_SIST,    " & vbCrLf & _
            "           CO_EMPR     " & vbCrLf & _
            " FROM      A7.TB_SIST  " & vbCrLf & _
            " WHERE     CO_EMPR = " & CLng(xmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text) & vbCrLf & _
            " AND       SG_SIST = '" & xmlMensagem.documentElement.selectSingleNode("SG_SIST_ORIG").Text & "' "
    
    strSQL = strSQL & " AND   DT_INIC_VIGE_SIST <= " & fgDataHoraServidor_To_Oracle & _
                      " AND  (DT_FIM_VIGE_SIST    IS  NULL       " & _
                      " OR    DT_FIM_VIGE_SIST  >= " & fgDataHoraServidor_To_Oracle & ")"
            
    
    Set Rs = fgQuerySQL(strSQL)
    
    If Rs.EOF Then
        flValidaSistemaOrigem = False
    Else
        flValidaSistemaOrigem = True
    End If

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flValidaSistemaOrigem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Validar o sistema destino de uma regra de transporte.
Private Function flValidaSistemaDestino(ByRef xmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim Rs                                      As ADODB.Recordset

On Error GoTo ErrorHandler

    If InStr(1, xmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text, "'") > 0 Then
        flValidaSistemaDestino = False
        Exit Function
    End If
    
    If Trim(xmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text) <> vbNullString Then
    
        strSQL = " SELECT   SG_SIST,    " & vbCrLf & _
                "           CO_EMPR     " & vbCrLf & _
                " FROM      A7.TB_SIST  " & vbCrLf & _
                " WHERE     CO_EMPR = " & CLng(xmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text) & vbCrLf & _
                " AND       SG_SIST = '" & xmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text & "' "
        
        strSQL = strSQL & " AND   DT_INIC_VIGE_SIST <= " & fgDataHoraServidor_To_Oracle & _
                          " AND  (DT_FIM_VIGE_SIST    IS  NULL       " & _
                          " OR    DT_FIM_VIGE_SIST  >= " & fgDataHoraServidor_To_Oracle & ")"
                
        
        Set Rs = fgQuerySQL(strSQL)
        
        If Rs.EOF Then
            flValidaSistemaDestino = False
        Else
            flValidaSistemaDestino = True
        End If
    
    Else
        flValidaSistemaDestino = True
    End If
    
    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flValidaSistemaDestino Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verificar a vigência de uma regra de transporte.
Private Function flRegraParaConsulta(ByVal plngCodigoEmpresa As Long, _
                                     ByVal pstrTipoMensagem As String, _
                                     ByVal plngTipoFormatoMensagemSaida As Long, _
                                     ByVal pstrSistemaOrigem As String, _
                                     ByVal pstrDataInicioVigencia As String) As Boolean

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset
Dim strLer                                  As String

On Error GoTo ErrorHandler

    'Informações da Regra de Tranformação
    strSQL = " SELECT   " & vbNewLine & _
             " TP_MESG, " & vbNewLine & _
             " TP_FORM_MESG_SAID, " & vbNewLine & _
             " DH_INIC_VIGE_REGR_TRAP " & vbNewLine & _
             "  FROM  A7.TB_REGR_TRAP_MESG " & vbNewLine & _
             "  WHERE CO_EMPR_ORIG                      = " & plngCodigoEmpresa & vbNewLine & _
             "    AND LTRIM(RTRIM(UPPER(TP_MESG)))      = '" & Trim(UCase(pstrTipoMensagem)) & "'" & vbNewLine & _
             "    AND TP_FORM_MESG_SAID                 =  " & plngTipoFormatoMensagemSaida & vbNewLine & _
             "    AND LTRIM(RTRIM(UPPER(SG_SIST_ORIG))) = '" & Trim(UCase(pstrSistemaOrigem)) & "'" & vbNewLine & _
             " ORDER BY DH_INIC_VIGE_REGR_TRAP DESC "

    Set objRS = fgQuerySQL(strSQL)
    
    If fgDtHr_To_Xml(objRS!DH_INIC_VIGE_REGR_TRAP) = pstrDataInicioVigencia Then
        flRegraParaConsulta = False
    Else
        flRegraParaConsulta = True
    End If
    

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flRegraParaConsulta Function", lngCodigoErroNegocio, intNumeroSequencialErro)
   
End Function

