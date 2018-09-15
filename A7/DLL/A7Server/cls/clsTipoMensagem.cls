VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF757DE03A9"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Objeto responsável pelo fornecimento de informações de tipos de mensagem, bem como a manipulação dos dados cadastrados
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

'Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a estrutura da tabela:
' - TB_TIPO_MESG
Public Function ObterPropriedades() As String

Dim strSQL                                  As String
Dim objPropriedade                          As MSXML2.DOMDocument40
Dim strPropriedade                          As String

On Error GoTo ErrorHandler

    Set objPropriedade = CreateObject("MSXML2.DOMDocument.4.0")
       
    Call fgAppendNode(objPropriedade, "", "Tipo_Mensagem", "")

    strSQL = " SELECT  " & _
             " TP_MESG, " & _
             " TP_FORM_MESG_SAID, " & _
             " NO_TIPO_MESG, " & _
             " TP_CTER_DELI, " & _
             " TP_NATZ_MESG, " & _
             " TP_CTER_DELI, " & _
             " CO_PRIO_FILA_SAID_MESG,  " & _
             " CO_TEXT_XML,             " & _
             " NO_TITU_MESG,            " & _
             " DT_INIC_VIGE_MESG,       " & _
             " DT_FIM_VIGE_MESG,        " & _
             " CO_USUA_ULTI_ATLZ,       " & _
             " CO_ETCA_TRAB_ULTI_ATLZ,  " & _
             " DH_ULTI_ATLZ,            " & _
             " '' AS TX_VALID_SAID_MESG " & _
             " FROM A7.TB_TIPO_MESG  "

    strPropriedade = fgPropriedades("Grupo_TipoMensagem", strSQL, "A7Server.clsTipoMensagem")

    Call fgAppendXML(objPropriedade, "Tipo_Mensagem", strPropriedade)

    strSQL = " SELECT    * " & _
            " FROM      A7.TB_TIPO_MESG_ATRB "
    
    strPropriedade = fgPropriedades("Grupo_TipoMensagemAtributo", strSQL, "A7Server.clsTipoMensagem")
    
    Call fgAppendNode(objPropriedade, "Tipo_Mensagem", "Repeat_TipoMensagemAtributo", "")
    
    Call fgAppendXML(objPropriedade, "Repeat_TipoMensagemAtributo", strPropriedade)
    
    ObterPropriedades = objPropriedade.xml
    
    Set objPropriedade = Nothing

    Exit Function
ErrorHandler:

    Set objPropriedade = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Consultar todos os atributos de um tipo de mensagem específico (Tipo de mensagem + atributos da mensagem)
Public Function Ler(ByVal pstrTipoMensagem As String, _
                    ByVal plngTipoFormatoMensagemSaida As Long) As String

Dim strSQL                                  As String
Dim xmlLer                                  As MSXML2.DOMDocument40
Dim strLer                                  As String

On Error GoTo ErrorHandler

    Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlLer, "", "Tipo_Mensagem", "")
       
    strSQL = " SELECT  " & _
             " A.TP_MESG, " & _
             " A.TP_FORM_MESG_SAID, " & _
             " A.NO_TIPO_MESG, " & _
             " A.TP_CTER_DELI, " & _
             " A.TP_NATZ_MESG, " & _
             " A.CO_PRIO_FILA_SAID_MESG, " & _
             " A.CO_TEXT_XML, " & _
             " A.NO_TITU_MESG, " & _
             " A.DT_INIC_VIGE_MESG, " & _
             " A.DT_FIM_VIGE_MESG, " & _
             " A.CO_USUA_ULTI_ATLZ, " & _
             " A.CO_ETCA_TRAB_ULTI_ATLZ, " & _
             " A.DH_ULTI_ATLZ, " & _
             " '' AS TX_VALID_SAID_MESG, " & _
             " A.NO_TITU_MESG " & _
             " FROM  A7.TB_TIPO_MESG A " & _
             " WHERE A.TP_MESG = '" & Trim(pstrTipoMensagem) & "'" & _
             "   AND A.TP_FORM_MESG_SAID = " & plngTipoFormatoMensagemSaida
    
    strLer = fgQueryXMLLer("Grupo_TipoMensagem", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Len(strLer) = 0 Then
        '4002-Registro inexistente ou excluído por outro usuário
        lngCodigoErroNegocio = 4002
        GoTo ErrorHandler
    End If
    
    Call fgAppendXML(xmlLer, "Tipo_Mensagem", strLer)
    
    strSQL = " SELECT    B.TP_MESG,          " & _
             "           B.TP_FORM_MESG_SAID, " & _
             "           A.NO_ATRB_MESG,     " & _
             "           A.TP_DADO_ATRB_MESG," & _
             "           A.QT_CTER_ATRB,     " & _
             "           A.QT_CASA_DECI_ATRB," & _
             "           A.NO_TRAP_ATRB,     " & _
             "           A.IN_ATRB_PRMT_VALO_NEGT,     " & _
             "           B.NO_ATRB_MESG,     " & _
             "           B.TP_FORM_MESG,     " & _
             "           B.NU_ORDE_AGRU_ATRB," & _
             "           B.IN_OBRI_ATRB,     " & _
             "           B.QT_REPE,          " & _
             "           B.NU_NIVE_MESG_ATRB, " & _
             "           DECODE(B.TP_FORM_MESG, 2, 5, B.TP_FORM_MESG) AUXORDEM " & _
             " FROM      A7.TB_ATRB_MESG   A, " & _
             "           A7.TB_TIPO_MESG_ATRB    B " & _
             " WHERE     A.NO_ATRB_MESG = B.NO_ATRB_MESG   " & _
             "   AND     RTRIM(LTRIM(TP_MESG)) = '" & Trim(pstrTipoMensagem) & "'" & _
             "   AND     B.TP_FORM_MESG_SAID = " & plngTipoFormatoMensagemSaida & _
             " ORDER BY  15, NU_ORDE_AGRU_ATRB"
             
    
    strLer = fgQueryXMLLerTodos("TipoMensagemAtributo", strSQL, App.EXEName & "." & TypeName(Me))
    
    If strLer <> vbNullString Then
        Call fgAppendXML(xmlLer, "Grupo_TipoMensagem", strLer)
    End If
    
    Ler = xmlLer.xml
    
    Set xmlLer = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlLer = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Consultar todos os tipos de mensagem cadastrados
Public Function LerTodos(Optional ByVal pstrIndicadorVigencia As String = "N") As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT     A.TP_MESG,                 " & vbNewLine & _
             "            A.TP_FORM_MESG_SAID,       " & vbNewLine & _
             "            A.DT_INIC_VIGE_MESG,       " & vbNewLine & _
             "            A.NO_TIPO_MESG,            " & vbNewLine & _
             "            A.TP_NATZ_MESG,            " & vbNewLine & _
             "            A.TP_CTER_DELI,            " & vbNewLine & _
             "            A.CO_PRIO_FILA_SAID_MESG,  " & vbNewLine & _
             "            A.DT_FIM_VIGE_MESG,        " & vbNewLine & _
             "            A.CO_USUA_ULTI_ATLZ,       " & vbNewLine & _
             "            A.CO_ETCA_TRAB_ULTI_ATLZ,  " & vbNewLine & _
             "            A.DH_ULTI_ATLZ,            " & vbNewLine & _
             "            A.CO_TEXT_XML,             " & vbNewLine & _
             "            A.NO_TITU_MESG             " & vbNewLine & _
             " FROM      A7.TB_TIPO_MESG A " & vbNewLine
             
    strSQL = strSQL & " ORDER BY  TP_MESG, TP_FORM_MESG_SAID ASC "
    
    LerTodos = fgQueryXMLLerTodos("TipoMensagem", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Controlar o roteamento e execução dos métodos privados, tais como:
' -  Incluir
' -  Alterar
' -  Excluir
' -  LerTodos
' -  Ler
Public Function Executar(ByRef pxmlTipoMensagem As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlTipoMensagem.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar"
            Executar = Salvar(pxmlTipoMensagem)
        Case "Ler"
            Executar = Ler(pxmlTipoMensagem.selectSingleNode("//TP_MESG").Text, _
                           CLng(pxmlTipoMensagem.selectSingleNode("//TP_FORM_MESG_SAID").Text))
        Case "LerTodos"
            Executar = LerTodos()
        Case "Excluir"
            Executar = Excluir(pxmlTipoMensagem)
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

'Roteia comando de Salvar para as funções privadas Incluir ou Alterar
Public Function Salvar(ByRef pxmlTipoMensagem As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    flRegrasNegocio pxmlTipoMensagem
            
    If pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/@Operacao").Text = "Incluir" Then
        flIncluir pxmlTipoMensagem
        flIncluirAtributos pxmlTipoMensagem
    ElseIf pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/@Operacao").Text = "Alterar" Then
        'If flExisteRegraTraducao(pxmlTipoMensagem) Then
'            If MsgBox("Existe Regra cadastrada para essa mensagem. Deseja alterar somente o nome? ", vbYesNo, "Tipo de Mensagem") = vbNo Then
         '       lngCodigoErroNegocio = 1005
          '      GoTo ErrorHandler
'            Else
'                Call flAlterarNomeTitulo(pxmlTipoMensagem)
'            End If
        'Else
            flAlterar pxmlTipoMensagem
            flIncluirAtributos pxmlTipoMensagem
        'End If
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

'Gera XSDs das mensagens quando ocorre a limpeza da base
'Este método é chamado por um executável que é utilizado no Santander
Public Sub GerarXSD(pxmlTipoMensagem As DOMDocument40)
Dim lngCodCntdAtrbXml                        As Long
Dim strSQL                                   As String

On Error GoTo ErrorHandler
    
    With pxmlTipoMensagem
            
        lngCodCntdAtrbXml = CLng(.selectSingleNode("//Grupo_TipoMensagem/CO_TEXT_XML").Text)
            
        strSQL = " DELETE FROM A7.TB_TEXT_XML " & _
                 "  WHERE CO_TEXT_XML = " & lngCodCntdAtrbXml
            
        fgExecuteSQL strSQL
            
        lngCodCntdAtrbXml = fgInsertVarchar4000(.selectSingleNode("//Grupo_TipoMensagem/TX_VALID_SAID_MESG").Text)
        
        strSQL = "UPDATE A7.TB_TIPO_MESG " & _
                 "SET NO_TIPO_MESG           = '" & fgLimpaCaracterEspecial(.selectSingleNode("//Grupo_TipoMensagem/NO_TIPO_MESG").Text) & "'," & vbCrLf & _
                     "TP_NATZ_MESG           = " & .selectSingleNode("//Grupo_TipoMensagem/TP_NATZ_MESG").Text & "," & vbCrLf & _
                     "TP_CTER_DELI           = '" & .selectSingleNode("//Grupo_TipoMensagem/TP_CTER_DELI").Text & "'," & vbCrLf & _
                     "CO_PRIO_FILA_SAID_MESG = " & .selectSingleNode("//Grupo_TipoMensagem/CO_PRIO_FILA_SAID_MESG").Text & "," & vbCrLf & _
                     "CO_TEXT_XML            = " & lngCodCntdAtrbXml & "," & vbCrLf & _
                     "DT_INIC_VIGE_MESG      = " & fgDtXML_To_Oracle(.selectSingleNode("//Grupo_TipoMensagem/DT_INIC_VIGE_MESG").Text) & "," & vbCrLf & _
                     "DT_FIM_VIGE_MESG       = " & fgDtXML_To_Oracle(.selectSingleNode("//Grupo_TipoMensagem/DT_FIM_VIGE_MESG").Text) & "," & vbCrLf & _
                     "CO_USUA_ULTI_ATLZ      = 'SISTEMA'," & vbCrLf & _
                     "CO_ETCA_TRAB_ULTI_ATLZ = 'SERVIDOR'," & vbCrLf & _
                     "DH_ULTI_ATLZ           = SYSDATE" & vbCrLf & _
                 " WHERE TP_MESG             = '" & .selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text & "'" & _
                 "   AND TP_FORM_MESG_SAID   = " & .selectSingleNode("//Grupo_TipoMensagem/TP_FORM_MESG_SAID").Text & _
                 "   AND DH_ULTI_ATLZ        = " & fgDtHrXML_To_Oracle(.selectSingleNode("//Grupo_TipoMensagem/DH_ULTI_ATLZ").Text)
    
        If fgExecuteSQL(strSQL) = 0 Then
           '4001 - Registro foi alterado ou excluído por outro usuário.
           lngCodigoErroNegocio = 4001
           GoTo ErrorHandler
        End If
    
    End With
      
    
    Exit Sub

ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerarXSD", lngCodigoErroNegocio, intNumeroSequencialErro)
End Sub

'Promover a inclusão de um novo tipo de mensagem
Private Function flIncluir(ByRef pxmlTipoMensagem As MSXML2.DOMDocument40) As Boolean


Dim lngCodCntdAtrbXml                        As Long
Dim strSQL                                   As String

On Error GoTo ErrorHandler

    With pxmlTipoMensagem
        
        lngCodCntdAtrbXml = fgInsertVarchar4000(.selectSingleNode("//Grupo_TipoMensagem/TX_VALID_SAID_MESG").Text)
        
        strSQL = "INSERT INTO A7.TB_TIPO_MESG ( " & _
                "TP_MESG,           " & _
                "TP_FORM_MESG_SAID,           " & _
                "NO_TIPO_MESG,      " & _
                "TP_NATZ_MESG,      " & _
                "TP_CTER_DELI,      " & _
                "CO_PRIO_FILA_SAID_MESG, " & _
                "CO_TEXT_XML,  " & _
                "NO_TITU_MESG, " & _
                "DT_INIC_VIGE_MESG, " & _
                "DT_FIM_VIGE_MESG,  " & _
                "CO_USUA_ULTI_ATLZ, " & _
                "CO_ETCA_TRAB_ULTI_ATLZ, " & _
                "DH_ULTI_ATLZ       " & _
                " ) VALUES ( "
        
        strSQL = strSQL & "'" & .selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text & "'," & _
                .selectSingleNode("//Grupo_TipoMensagem/TP_FORM_MESG_SAID").Text & "," & _
                "'" & .selectSingleNode("//Grupo_TipoMensagem/NO_TIPO_MESG").Text & "'," & _
                .selectSingleNode("//Grupo_TipoMensagem/TP_NATZ_MESG").Text & "," & _
                IIf(.selectSingleNode("//Grupo_TipoMensagem/TP_CTER_DELI").Text = vbNullString, "NULL", "'" & .selectSingleNode("//Grupo_TipoMensagem/TP_CTER_DELI").Text & "'") & "," & _
                "'" & .selectSingleNode("//Grupo_TipoMensagem/CO_PRIO_FILA_SAID_MESG").Text & "'," & _
                lngCodCntdAtrbXml & "," & _
                "'" & .selectSingleNode("//Grupo_TipoMensagem/NO_TITU_MESG").Text & "'," & _
                fgDtXML_To_Oracle(.selectSingleNode("//Grupo_TipoMensagem/DT_INIC_VIGE_MESG").Text) & "," & _
                fgDtXML_To_Oracle(.selectSingleNode("//Grupo_TipoMensagem/DT_FIM_VIGE_MESG").Text) & "," & _
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

'Promover a alteração de um tipo de mensagem existente.
Private Function flAlterar(ByRef pxmlTipoMensagem As MSXML2.DOMDocument40) As Boolean

Dim lngCodCntdAtrbXml                        As Long
Dim strSQL                                   As String

On Error GoTo ErrorHandler
    
    With pxmlTipoMensagem
            
        lngCodCntdAtrbXml = CLng(.selectSingleNode("//Grupo_TipoMensagem/CO_TEXT_XML").Text)
            
        strSQL = " DELETE FROM A7.TB_TEXT_XML " & _
                 "  WHERE CO_TEXT_XML = " & lngCodCntdAtrbXml
            
        fgExecuteSQL strSQL
            
        lngCodCntdAtrbXml = fgInsertVarchar4000(.selectSingleNode("//Grupo_TipoMensagem/TX_VALID_SAID_MESG").Text)
        
        strSQL = "UPDATE A7.TB_TIPO_MESG " & _
                 "SET NO_TIPO_MESG           = '" & fgLimpaCaracterEspecial(.selectSingleNode("//Grupo_TipoMensagem/NO_TIPO_MESG").Text) & "'," & vbCrLf & _
                     "TP_NATZ_MESG           = " & .selectSingleNode("//Grupo_TipoMensagem/TP_NATZ_MESG").Text & "," & vbCrLf & _
                     "TP_CTER_DELI           = '" & .selectSingleNode("//Grupo_TipoMensagem/TP_CTER_DELI").Text & "'," & vbCrLf & _
                     "NO_TITU_MESG           = '" & .selectSingleNode("//Grupo_TipoMensagem/NO_TITU_MESG").Text & "'," & vbCrLf & _
                     "CO_PRIO_FILA_SAID_MESG = " & .selectSingleNode("//Grupo_TipoMensagem/CO_PRIO_FILA_SAID_MESG").Text & "," & vbCrLf & _
                     "CO_TEXT_XML            = " & lngCodCntdAtrbXml & "," & vbCrLf & _
                     "DT_INIC_VIGE_MESG      = " & fgDtXML_To_Oracle(.selectSingleNode("//Grupo_TipoMensagem/DT_INIC_VIGE_MESG").Text) & "," & vbCrLf & _
                     "DT_FIM_VIGE_MESG       = " & fgDtXML_To_Oracle(.selectSingleNode("//Grupo_TipoMensagem/DT_FIM_VIGE_MESG").Text) & "," & vbCrLf & _
                     "CO_USUA_ULTI_ATLZ      = '" & fgObterUsuarioRede & "'," & vbCrLf & _
                     "CO_ETCA_TRAB_ULTI_ATLZ = '" & fgObterEstacaoTrabalhoUsuario & "'," & vbCrLf & _
                     "DH_ULTI_ATLZ           = SYSDATE" & vbCrLf & _
                 " WHERE TP_MESG             = '" & .selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text & "'" & _
                 "   AND TP_FORM_MESG_SAID   = " & .selectSingleNode("//Grupo_TipoMensagem/TP_FORM_MESG_SAID").Text & _
                 "   AND DH_ULTI_ATLZ        = " & fgDtHrXML_To_Oracle(.selectSingleNode("//Grupo_TipoMensagem/DH_ULTI_ATLZ").Text)
    
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

'Incluir os atributos para o tipo de mensagem alterado ou incluido.
Private Sub flIncluirAtributos(ByRef pxmlTipoMensagem As MSXML2.DOMDocument40)

Dim strSQL                                   As String

Dim objGrupoAtributo                        As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    'Antes de salvar apagar todos os atributos deeste evento
    
    strSQL = " DELETE FROM A7.TB_TIPO_MESG_ATRB " & _
             " WHERE TP_MESG = '" & pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text & "'" & _
             "   AND TP_FORM_MESG_SAID = " & pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/TP_FORM_MESG_SAID").Text
    
    fgExecuteSQL strSQL
    
    If pxmlTipoMensagem.selectNodes("//Grupo_TipoMensagemAtributo").length = 0 Then
        Exit Sub
    End If
    
    'Inclui novos atributos do evento
            
    strSQL = "BEGIN" & vbCrLf
    For Each objGrupoAtributo In pxmlTipoMensagem.selectNodes("//Repeat_TipoMensagemAtributo/*")
        
        strSQL = strSQL & _
                "INSERT INTO A7.TB_TIPO_MESG_ATRB(TP_MESG, " & _
                                                  "TP_FORM_MESG_SAID, " & _
                                                  "NO_ATRB_MESG, " & _
                                                  "TP_FORM_MESG, " & _
                                                  "IN_OBRI_ATRB, " & _
                                                  "NU_ORDE_AGRU_ATRB, " & vbCrLf & _
                                                  "QT_REPE, " & vbCrLf & _
                                                  "NU_NIVE_MESG_ATRB) " & vbCrLf & _
                "VALUES( '" & objGrupoAtributo.selectSingleNode("TP_MESG").Text & "', " & _
                        objGrupoAtributo.selectSingleNode("TP_FORM_MESG_SAID").Text & ", " & _
                        "'" & objGrupoAtributo.selectSingleNode("NO_ATRB_MESG").Text & "', " & _
                        objGrupoAtributo.selectSingleNode("TP_FORM_MESG").Text & ", " & _
                        "'" & objGrupoAtributo.selectSingleNode("IN_OBRI_ATRB").Text & "', " & _
                        objGrupoAtributo.selectSingleNode("NU_ORDE_AGRU_ATRB").Text & ", " & vbCrLf & _
                        objGrupoAtributo.selectSingleNode("QT_REPE").Text & ", " & vbCrLf & _
                        objGrupoAtributo.selectSingleNode("NU_NIVE_MESG_ATRB").Text & ");" & vbCrLf
                        
                                                    
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

'Verificar as regras de negócio pertinentes a alteração / inclusão ou exclusão de tipos de mensagem.
Private Function flRegrasNegocio(ByRef pxmlTipoMensagem As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    If pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/@Operacao").Text = "Incluir" Then
        
        If VerificaTipoMensagem(pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text, _
                                CLng(pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/TP_FORM_MESG_SAID").Text)) Then
            'Tipo de Mensagem já cadastrado.
            lngCodigoErroNegocio = 1007
            GoTo ErrorHandler
        End If
        
    End If
    
    If pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/@Operacao").Text = "Excluir" Then
        
        If Not IsNumeric(pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text) Then
            'Tipo de Mensagem alfanumérico (catalogo SPB) não pode ser excluído.
            lngCodigoErroNegocio = 1013
            GoTo ErrorHandler
        End If
        
    End If
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flRegrasNegocio Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verificar a existência de um tipo de mensagem
Public Function VerificaTipoMensagem(ByVal pstrTipoMensagem As String, _
                                     ByVal plngTipoFormatoMensagemSaida As Long) As Boolean

Dim strSQL                                   As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    
    strSQL = " SELECT TP_MESG " & _
             "   FROM A7.TB_TIPO_MESG " & _
             "  WHERE TP_MESG = '" & pstrTipoMensagem & "'" & _
             "    AND TP_FORM_MESG_SAID = " & plngTipoFormatoMensagemSaida

    Set objRS = fgQuerySQL(strSQL)
    
    VerificaTipoMensagem = Not objRS.EOF
    
    objRS.Close
    
    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "VerificaTipoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Verificar a existência de regra de transporte cadastrada para o tipo de mensagem.
Private Function flExisteRegraTraducao(ByRef pxmlTipoMensagem As MSXML2.DOMDocument40) As Boolean

Dim objRegraTraducao                        As A7Server.clsRegraTransporte

On Error GoTo ErrorHandler


    Set objRegraTraducao = CreateObject("A7Server.clsRegraTransporte")
    flExisteRegraTraducao = objRegraTraducao.ExisteRegraTransporte(pxmlTipoMensagem.selectSingleNode("//TP_MESG").Text, _
                                                                   CLng(pxmlTipoMensagem.selectSingleNode("//TP_FORM_MESG_SAID").Text))
    Set objRegraTraducao = Nothing
    
    Exit Function

ErrorHandler:
    
    Set objRegraTraducao = Nothing
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExisteRegraTraducao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Promover a exclusão de um tipo de mensagem existente
Private Function Excluir(ByRef pxmlTipoMensagem As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    flRegrasNegocio pxmlTipoMensagem
       
    If flExisteRegraTraducao(pxmlTipoMensagem) Then
        lngCodigoErroNegocio = 1005
        GoTo ErrorHandler
    End If
    
    strSQL = " DELETE FROM A7.TB_TIPO_MESG_ATRB " & _
             " WHERE TP_MESG = '" & pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text & "'" & _
             "   AND TP_FORM_MESG_SAID = " & pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/TP_FORM_MESG_SAID").Text
    
    fgExecuteSQL strSQL
    
    strSQL = " DELETE FROM A7.TB_TIPO_MESG " & _
             " WHERE TP_MESG = '" & pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text & "'" & _
             "   AND TP_FORM_MESG_SAID = " & pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/TP_FORM_MESG_SAID").Text & _
             "   AND DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/DH_ULTI_ATLZ").Text)
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001-Registro foi alterado ou excluido por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If

    
    Excluir = True
    
    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Excluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Verificar se existe tipo de mensagem vigente
Public Function ExisteTipoMensagem(ByVal pstrTipoMensagem As String, _
                          Optional ByVal plngTipoFormatoMensagemSaida As Long = 0) As Boolean

Dim strSQL                                   As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    pstrTipoMensagem = fgLimpaCaracterEspecial(pstrTipoMensagem)
    
    strSQL = " SELECT  TP_MESG " & _
             "   FROM  A7.TB_TIPO_MESG " & _
             "  WHERE  LTRIM(RTRIM(TP_MESG)) = '" & IIf(IsNumeric(pstrTipoMensagem), Val(pstrTipoMensagem), Trim(pstrTipoMensagem)) & "'" & _
             "    AND  (DT_FIM_VIGE_MESG    IS NULL " & _
             "     OR   DT_FIM_VIGE_MESG >=  " & fgDataHoraServidor_To_Oracle & ")"
             
    If plngTipoFormatoMensagemSaida <> 0 Then
        strSQL = strSQL & "    AND  TP_FORM_MESG_SAID = " & plngTipoFormatoMensagemSaida
    End If

    Set objRS = fgQuerySQL(strSQL)
    
    ExisteTipoMensagem = Not objRS.EOF
    
    objRS.Close
  
    Set objRS = Nothing
    
    Exit Function

ErrorHandler:
    Set objRS = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExisteTipoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Promover a alteração de um tipo de mensagem existente.
Private Function flAlterarNomeTitulo(ByRef pxmlTipoMensagem As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler
    
    strSQL = "UPDATE A7.TB_TIPO_MESG " & _
             "SET NO_TIPO_MESG           = '" & fgLimpaCaracterEspecial(pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/NO_TIPO_MESG").Text) & "'," & vbCrLf & _
                 "NO_TITU_MESG           = '" & pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/NO_TITU_MESG").Text & "'," & vbCrLf & _
                 "DH_ULTI_ATLZ           = SYSDATE" & vbCrLf & _
             " WHERE TP_MESG             = '" & pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text & "'" & _
             "   AND TP_FORM_MESG_SAID   = " & pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/TP_FORM_MESG_SAID").Text & _
             "   AND DH_ULTI_ATLZ        = " & fgDtHrXML_To_Oracle(pxmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/DH_ULTI_ATLZ").Text)

    If fgExecuteSQL(strSQL) = 0 Then
       '4001 - Registro foi alterado ou excluído por outro usuário.
       lngCodigoErroNegocio = 4001
       GoTo ErrorHandler
    End If

Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterarNomeTitulo Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

