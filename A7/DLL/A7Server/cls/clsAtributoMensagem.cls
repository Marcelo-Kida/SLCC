VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAtributoMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EFC4644003E"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Objeto responsável pelo fornecimento de informações de atributos de mensagem, bem como a manipulação dos dados cadastrados
Option Explicit

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

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

'Consultar todos os atributos de mensagem cadastrados.
Public Function LerTodos(Optional ByVal pstrIndicadorVigencia As String = "N") As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = "  SELECT   NO_ATRB_MESG,           " & vbNewLine & _
             "           DT_FIM_VIGE_ATRB_MESG,  " & vbNewLine & _
             "           QT_CASA_DECI_ATRB,      " & vbNewLine & _
             "           TP_DADO_ATRB_MESG,      " & vbNewLine & _
             "           DE_ATRB_MESG,           " & vbNewLine & _
             "           NO_TRAP_ATRB,           " & vbNewLine & _
             "           DT_INIC_VIGE_ATRB_MESG, " & vbNewLine & _
             "           IN_ATRB_PRMT_VALO_NEGT, " & vbNewLine & _
             "           QT_CTER_ATRB,           " & vbNewLine & _
             "           CO_USUA_ULTI_ATLZ,      " & vbNewLine & _
             "           CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
             "           DH_ULTI_ATLZ            " & vbNewLine & _
             "     FROM   A7.TB_ATRB_MESG        "
    
    If pstrIndicadorVigencia = "S" Then
        strSQL = strSQL & " WHERE  DT_INIC_VIGE_ATRB_MESG <= " & fgDataHoraServidor_To_Oracle & _
                          "   AND (DT_FIM_VIGE_ATRB_MESG  IS NULL " & _
                          "    OR  DT_FIM_VIGE_ATRB_MESG  >=  " & fgDataHoraServidor_To_Oracle & ")"
        
    End If
    
    strSQL = strSQL & " ORDER BY   NO_ATRB_MESG ASC "

    LerTodos = fgQueryXMLLerTodos("AtributoMensagem", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a estrutura da tabela TB_ATRB_MESG.
Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = "SELECT " & _
                " NO_ATRB_MESG, " & _
                " DE_ATRB_MESG, " & _
                " NO_TRAP_ATRB, " & _
                " TP_DADO_ATRB_MESG, " & _
                " QT_CASA_DECI_ATRB, " & _
                " QT_CTER_ATRB, " & _
                " DT_INIC_VIGE_ATRB_MESG, " & _
                " DT_FIM_VIGE_ATRB_MESG, " & _
                " CO_USUA_ULTI_ATLZ, " & _
                " CO_ETCA_TRAB_ULTI_ATLZ, " & _
                " DH_ULTI_ATLZ, " & _
                " 'N' AS IN_VIGE,  " & _
                " IN_ATRB_PRMT_VALO_NEGT " & _
             " FROM A7.TB_ATRB_MESG "

    ObterPropriedades = fgPropriedades("Grupo_AtributoMensagem", strSQL, "A7Server.clsAtributoMensagem")

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Controlar o roteamento e execução dos métodos privados, tais como:
' -  Incluir
' -  Alterar
' -  Excluir
' -  LerTodos
' -  Ler
Public Function Executar(ByVal pxmlAtributoMensagem As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = pxmlAtributoMensagem.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pxmlAtributoMensagem)
        Case "Ler"
            Executar = Ler(pxmlAtributoMensagem.selectSingleNode("//NO_ATRB_MESG").Text)
        Case "LerTodos"
            Executar = LerTodos(pxmlAtributoMensagem.selectSingleNode("//IN_VIGE").Text)
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

'Promove a inclusão de um novo atributo de mensagem.
Private Function flIncluir(ByRef pxmlAtributoMensagem As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

    On Error GoTo ErrorHandler
    
    With pxmlAtributoMensagem
        
        strSQL = " INSERT INTO A7.TB_ATRB_MESG (" & _
                "           NO_ATRB_MESG, " & _
                "           NO_TRAP_ATRB, " & _
                "           DE_ATRB_MESG, " & _
                "           TP_DADO_ATRB_MESG, " & _
                "           QT_CTER_ATRB, " & _
                "           QT_CASA_DECI_ATRB, " & _
                "           IN_ATRB_PRMT_VALO_NEGT, " & _
                "           DT_INIC_VIGE_ATRB_MESG, " & _
                "           DT_FIM_VIGE_ATRB_MESG, " & _
                "           CO_USUA_ULTI_ATLZ, " & _
                "           CO_ETCA_TRAB_ULTI_ATLZ, " & _
                "           DH_ULTI_ATLZ "
        strSQL = strSQL & _
                " ) VALUES ( " & _
                "'" & .selectSingleNode("//NO_ATRB_MESG").Text & "'," & _
                "'" & .selectSingleNode("//NO_TRAP_ATRB").Text & "'," & _
                "'" & .selectSingleNode("//DE_ATRB_MESG").Text & "'," & _
                CLng(.selectSingleNode("//TP_DADO_ATRB_MESG").Text) & "," & _
                CLng(.selectSingleNode("//QT_CTER_ATRB").Text) & "," & _
                CLng(.selectSingleNode("//QT_CASA_DECI_ATRB").Text) & "," & _
                CLng(.selectSingleNode("//IN_ATRB_PRMT_VALO_NEGT").Text) & "," & _
                fgDtXML_To_Oracle(.selectSingleNode("//DT_INIC_VIGE_ATRB_MESG").Text) & "," & _
                fgDtXML_To_Oracle(.selectSingleNode("//DT_FIM_VIGE_ATRB_MESG").Text) & "," & _
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

'Promove a alteração de um atributo de mensagem cadastrado.
Private Function flAlterar(ByRef pxmlAtributoMensagem As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

    On Error GoTo ErrorHandler
                
    With pxmlAtributoMensagem
                
        strSQL = " UPDATE A7.TB_ATRB_MESG " & _
                "  SET    NO_TRAP_ATRB            = '" & .selectSingleNode("//NO_TRAP_ATRB").Text & "'," & _
                "         DE_ATRB_MESG            = '" & .selectSingleNode("//DE_ATRB_MESG").Text & "'," & _
                "         TP_DADO_ATRB_MESG       =  " & CLng(.selectSingleNode("//TP_DADO_ATRB_MESG").Text) & "," & _
                "         QT_CTER_ATRB            =  " & CLng(.selectSingleNode("//QT_CTER_ATRB").Text) & "," & _
                "         QT_CASA_DECI_ATRB       =  " & CLng(.selectSingleNode("//QT_CASA_DECI_ATRB").Text) & "," & _
                "         IN_ATRB_PRMT_VALO_NEGT  =  " & CLng(.selectSingleNode("//IN_ATRB_PRMT_VALO_NEGT").Text) & "," & _
                "         DT_INIC_VIGE_ATRB_MESG  = " & fgDtXML_To_Oracle(.selectSingleNode("//DT_INIC_VIGE_ATRB_MESG").Text) & "," & _
                "         DT_FIM_VIGE_ATRB_MESG   = " & fgDtXML_To_Oracle(.selectSingleNode("//DT_FIM_VIGE_ATRB_MESG").Text) & "," & _
                "         CO_USUA_ULTI_ATLZ       = '" & fgObterUsuarioRede & "'," & _
                "         CO_ETCA_TRAB_ULTI_ATLZ  = '" & fgObterEstacaoTrabalhoUsuario & "'," & _
                "         DH_ULTI_ATLZ            =  SYSDATE " & _
                " WHERE   NO_ATRB_MESG            = '" & .selectSingleNode("//NO_ATRB_MESG").Text & "'" & vbNewLine & _
                " AND     DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(.selectSingleNode("//DH_ULTI_ATLZ").Text)
    
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

'Promove a alteração de dados específicos quando o atributo esta vigente e em uso.
Private Function flAlterarNomeDescricaoDatas(ByRef pxmlAtributoMensagem As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler
            
    With pxmlAtributoMensagem.documentElement
            
        strSQL = " UPDATE A7.TB_ATRB_MESG " & _
                "  SET    NO_TRAP_ATRB               = '" & .selectSingleNode("NO_TRAP_ATRB").Text & "'," & _
                "         DE_ATRB_MESG               = '" & .selectSingleNode("DE_ATRB_MESG").Text & "'," & _
                "         CO_USUA_ULTI_ATLZ          = '" & fgObterUsuarioRede & "'," & _
                "         CO_ETCA_TRAB_ULTI_ATLZ     = '" & fgObterEstacaoTrabalhoUsuario & "'," & _
                "         DT_FIM_VIGE_ATRB_MESG      = " & fgDtXML_To_Oracle(.selectSingleNode("DT_FIM_VIGE_ATRB_MESG").Text) & "," & _
                "         DH_ULTI_ATLZ               =  SYSDATE " & _
                " WHERE   NO_ATRB_MESG               = '" & .selectSingleNode("NO_ATRB_MESG").Text & "'" & _
                " AND     DH_ULTI_ATLZ               = " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text)
                
        If fgExecuteSQL(strSQL) = 0 Then
           '4001 - Registro foi alterado ou excluído por outro usuário.
           lngCodigoErroNegocio = 4001
           GoTo ErrorHandler
        End If
      
    End With
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterarNomeDescricaoDatas Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Consultar todos os atributos de um atributo de mensagem cadastrado.
Public Function Ler(ByVal pstrNomeAtributoMensagem As String) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   NO_ATRB_MESG, " & _
            "           NO_TRAP_ATRB, " & _
            "           DE_ATRB_MESG, " & _
            "           TP_DADO_ATRB_MESG, " & _
            "           QT_CTER_ATRB, " & _
            "           QT_CASA_DECI_ATRB, " & _
            "           IN_ATRB_PRMT_VALO_NEGT, " & _
            "           DT_INIC_VIGE_ATRB_MESG, " & _
            "           DT_FIM_VIGE_ATRB_MESG, " & _
            "           CO_USUA_ULTI_ATLZ, " & _
            "           CO_ETCA_TRAB_ULTI_ATLZ, " & _
            "           DH_ULTI_ATLZ " & _
            " FROM      A7.TB_ATRB_MESG " & _
            " WHERE     NO_ATRB_MESG ='" & pstrNomeAtributoMensagem & "'"
    
    Ler = fgQueryXMLLer("Grupo_AtributoMensagem", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Len(Ler) = 0 Then
       '4002 - Registro Inexistente ou excluído por outro usuário.
       lngCodigoErroNegocio = 4002
       GoTo ErrorHandler
    End If

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Rotear o comando salvar para as rotinas privadas flIncluir ou flAlterar.
Public Function Salvar(ByVal pxmlAtributoMensagem As MSXML2.DOMDocument40) As Boolean

    On Error GoTo ErrorHandler
    
    flRegrasNegocio pxmlAtributoMensagem
            
    With pxmlAtributoMensagem
            
        If .selectSingleNode("//@Operacao").Text = "Incluir" Then
            flIncluir pxmlAtributoMensagem
        
        ElseIf .selectSingleNode("//@Operacao").Text = "Alterar" Then
            If flVerificaAtributoTipoMensagem(.selectSingleNode("//NO_ATRB_MESG").Text) Then
                'Atributo não pode ser alterado.Atributo está associado a um tipo de mensagem,
                'Pode alterar somente o nome logico e descricao e desativar
                flAlterarNomeDescricaoDatas pxmlAtributoMensagem
            Else
                flAlterar pxmlAtributoMensagem
            End If
            
        ElseIf .selectSingleNode("//@Operacao").Text = "Excluir" Then
            flExcluir pxmlAtributoMensagem
        
        Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
        
        End If
    
    End With
    
    Salvar = True
      
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Consistir as regras de negócio pertinentes a operações efetuadas com atributos de mensagem.
Private Function flRegrasNegocio(ByRef pxmlAtributoMensagem As MSXML2.DOMDocument40) As Boolean

    On Error GoTo ErrorHandler
    
    With pxmlAtributoMensagem
        
        If .selectSingleNode("//@Operacao").Text = "Incluir" Then
            If flVerificaAtributoCadastrado(.selectSingleNode("//NO_ATRB_MESG").Text) Then
                'Atributo já está cadastrado
                lngCodigoErroNegocio = 1001
                GoTo ErrorHandler
            End If
            
        ElseIf .selectSingleNode("//@Operacao").Text = "Excluir" Then
            If flVerificaAtributoTipoMensagem(.selectSingleNode("//NO_ATRB_MESG").Text) Then
                'Atributo não pode ser excluído.Atributo está associado a um tipo de mensagem
                lngCodigoErroNegocio = 1011
                GoTo ErrorHandler
            End If
        
        End If
    
    End With

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flRegrasNegocio Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verificar se um atributo já está cadastrado.
Private Function flVerificaAtributoCadastrado(ByVal pstrNomeAtributoMensagem As String) As Boolean

Dim strSQL                                   As String
Dim objRS                                    As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT  COUNT(NO_ATRB_MESG) " & _
             "   FROM  A7.TB_ATRB_MESG " & _
             "  WHERE  UPPER(NO_ATRB_MESG) = '" & UCase(Trim(pstrNomeAtributoMensagem)) & "'"
    
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.fields(0) = 0 Then
        flVerificaAtributoCadastrado = False
    Else
        flVerificaAtributoCadastrado = True
    End If
    
    objRS.Close

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificaAtributoCadastrado Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verificar se já existe associação do atributo com algum tipo de mensagem.
Private Function flVerificaAtributoTipoMensagem(ByVal pstrNomeAtributoMensagem As String) As Boolean

Dim objTipoMemensagemAtributo               As A7Server.clsTipoMesgAtributo
Dim strTipoMemensagemAtributo               As String

On Error GoTo ErrorHandler
    
    Set objTipoMemensagemAtributo = CreateObject("A7Server.clsTipoMesgAtributo")
        
    strTipoMemensagemAtributo = objTipoMemensagemAtributo.Ler(pstrNomeAtributoMensagem)
        
    If Trim(strTipoMemensagemAtributo) = "" Then
        flVerificaAtributoTipoMensagem = False
    Else
        flVerificaAtributoTipoMensagem = True
    End If
        
    Set objTipoMemensagemAtributo = Nothing

    Exit Function

ErrorHandler:
    
    Set objTipoMemensagemAtributo = Nothing
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificaAtributoTipoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Promover a exclusão de um atributo de mensagem cadastrado.
Private Function flExcluir(ByRef pxmlAtributoMensagem As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler
    
    With pxmlAtributoMensagem

        strSQL = " DELETE FROM A7.TB_ATRB_MESG " & _
                 " WHERE  UPPER(NO_ATRB_MESG) = '" & UCase(Trim(.selectSingleNode("//NO_ATRB_MESG").Text)) & "'" & vbNewLine & _
                 " AND    DH_ULTI_ATLZ        =  " & fgDtHrXML_To_Oracle(.selectSingleNode("//DH_ULTI_ATLZ").Text)

    End With

    If fgExecuteSQL(strSQL) = 0 Then
       '4001 - Registro foi alterado ou excluído por outro usuário.
       lngCodigoErroNegocio = 4001
       GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
