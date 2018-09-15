VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrupoVeiculoLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA\r\nPacote         :\r\nClasse         : clsGrupoVeiculoLegal\r\nData Criação   : 22/07/2003\r\nObjetivo       :\r\n\r\nAnalista       : Carlos Fortes\r\n\r\nProgramador    : Cassiano Nicolosi\r\nData           : 22/07/2003\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F20393D019E"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"

'Objeto responsável pela manutenção do cadastro de grupo de veiculo legal

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

'Verifica se um grupo possue veiculos legais associados

Private Sub flVerificaVigenciaVeiculoLegal(ByVal plngGrupoVeiculoLegal As Long, _
                                           ByVal pstrDataFinalVigencia As String)
                                                          
Dim strSQL                                  As String
Dim strLer                                  As String

On Error GoTo ErrorHandler

    If pstrDataFinalVigencia <> vbNullString Then
    
       strSQL = " Select    CO_VEIC_LEGA        " & vbNewLine & _
                " From      A8.TB_VEIC_LEGA     " & vbNewLine & _
                " Where     CO_GRUP_VEIC_LEGA = " & plngGrupoVeiculoLegal & vbNewLine & _
                " AND       DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & vbNewLine & _
                " AND      (DT_FIM_VIGE    IS  NULL       " & vbNewLine & _
                " OR        DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
                
       strLer = fgQueryXMLLer("flRegradeNegocio", strSQL, App.EXEName & "." & TypeName(Me))
    
        If Len(strLer) <> 0 Then
           '4071 - Existem Veículos Legais Associados a este Grupo de Veículo Legal. Ele não pode ser desativado.
           lngCodigoErroNegocio = 4071
           GoTo ErrorHandler
        End If
                
    End If
                     
    Exit Sub
             
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificaVigenciaVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Recuperar os registros da tabela A8.TB_GRUP_VEIC_LEGA
'pstrIndicadorVigencia      -   Indicador de Vigencia (S-Sim ; N-Não)
'pintTipoBackOffice         -   Tipo de BackOffice
'pstrSegregaDados As String -   Indicador de Segregação de dados (S-Sim ; N-Não)
'prsGrupoVeiculoLegal       -   Retorno ADODB.Recordset

Public Function LerTodos(Optional ByVal pstrIndicadorVigencia As String = "S", _
                         Optional ByVal pintTipoBackOffice As enumTipoBackOffice, _
                         Optional ByVal pstrSegregaDados As String = "S", _
                         Optional ByRef prsGrupoVeiculoLegal As ADODB.Recordset) As String

Dim strSQL                                  As String
Dim strSQL02                                As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_GRUP_VEIC_LEGA, " & vbCrLf & _
            "           TP_BKOF, " & _
            "           NO_GRUP_VEIC_LEGA, " & vbCrLf & _
            "           DT_INIC_VIGE, " & vbCrLf & _
            "           DT_FIM_VIGE, " & vbCrLf & _
            "           CO_USUA_ULTI_ATLZ, " & vbCrLf & _
            "           CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
            "           DH_ULTI_ATLZ, " & vbCrLf & _
            "           'S' AS TP_VIGE, " & vbCrLf & _
            "           'S' AS TP_SEGR  " & _
            " FROM      A8.TB_GRUP_VEIC_LEGA A"
  
    If pstrIndicadorVigencia = "S" Then
       strSQL02 = " WHERE A.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & vbCrLf & _
                  " AND  (A.DT_FIM_VIGE    IS  NULL       " & vbCrLf & _
                  " OR    A.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
    End If
    
    If pintTipoBackOffice <> 0 And pintTipoBackOffice <> 99 And pstrIndicadorVigencia <> "S" Then
       strSQL02 = strSQL02 & " Where TP_BKOF = " & pintTipoBackOffice
    ElseIf pintTipoBackOffice <> 0 And pintTipoBackOffice <> 99 And pstrIndicadorVigencia = "S" Then
       strSQL02 = strSQL02 & " AND TP_BKOF = " & pintTipoBackOffice
    End If
    
    If pstrSegregaDados = "S" Then
       If pintTipoBackOffice <> 0 Then
          strSQL02 = strSQL02 & fgSegregaDados("A8.TB_GRUP_VEIC_LEGA", _
                                               False, _
                                               "A", _
                                               "A", _
                                               Len(strSQL02) = 0, _
                                               False, _
                                               False, _
                                               False, _
                                               True, _
                                               False)
       Else
          strSQL02 = strSQL02 & fgSegregaDados("A8.TB_GRUP_VEIC_LEGA", _
                                               False, _
                                               "A", _
                                               "A", _
                                               Len(strSQL02) = 0, _
                                               False, _
                                               True, _
                                               False, _
                                               True, _
                                               False)
       End If
    End If
    strSQL = strSQL & strSQL02 & " ORDER BY   A.CO_GRUP_VEIC_LEGA ASC "
    
    If prsGrupoVeiculoLegal Is Nothing Then
       LerTodos = fgQueryXMLLerTodos("GrupoVeiculoLegal", strSQL, App.EXEName & "." & TypeName(Me))
       If LerTodos = "" Then
            If flInserirGrupoPadrao Then
                LerTodos = fgQueryXMLLerTodos("GrupoVeiculoLegal", strSQL, App.EXEName & "." & TypeName(Me))
            End If
       End If
    Else
       Set prsGrupoVeiculoLegal = fgQuerySQL(strSQL)
       If prsGrupoVeiculoLegal.EOF Then
          flInserirGrupoPadrao
          Set prsGrupoVeiculoLegal = fgQuerySQL(strSQL)
       End If
       LerTodos = ""
    End If

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar os nomes das colunas da tabela A8.TB_GRUP_VEIC_LEGA


Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_GRUP_VEIC_LEGA, " & _
            "           TP_BKOF, " & _
            "           NO_GRUP_VEIC_LEGA, " & _
            "           DT_INIC_VIGE, " & _
            "           DT_FIM_VIGE, " & _
            "           CO_USUA_ULTI_ATLZ, " & _
            "           CO_ETCA_TRAB_ULTI_ATLZ, " & _
            "           DH_ULTI_ATLZ, " & _
            "           '' AS TP_VIGE, " & _
            "           '' AS TP_SEGR  " & _
            " FROM      A8.TB_GRUP_VEIC_LEGA "

    ObterPropriedades = fgPropriedades("Grupo_GrupoVeiculoLegal", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Executar as operações de Inclusão, Alteração, Exclusão , Ler, Lertodos

Public Function Executar(ByVal pxmlGrupoVeiculoLegal As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pxmlGrupoVeiculoLegal)
        Case "Ler"
            Executar = Ler(pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("//CO_GRUP_VEIC_LEGA").Text)
        Case "LerTodos"
            If pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("//TP_BKOF") Is Nothing Then
                Executar = LerTodos
            Else
                If pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("//TP_BKOF").Text <> "" Then
                    Executar = LerTodos(pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("//TP_VIGE").Text, _
                                        pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("//TP_BKOF").Text, _
                                        pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("//TP_SEGR").Text)
                Else
                    Executar = LerTodos(pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("//TP_VIGE").Text, 0, _
                                        pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("//TP_SEGR").Text)
                End If
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

'Incluir um registro na tabela A8.TB_GRUP_VEIC_LEGA

Private Function flIncluir(ByRef pxmlGrupoVeiculoLegal As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    flVerificarDescicaoIgual "Incluir", _
                             pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("TP_BKOF").Text, _
                             pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("CO_GRUP_VEIC_LEGA").Text, _
                             pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("NO_GRUP_VEIC_LEGA").Text
    
    strSQL = " INSERT INTO A8.TB_GRUP_VEIC_LEGA (" & _
             " CO_GRUP_VEIC_LEGA, " & _
             " TP_BKOF, " & _
             " NO_GRUP_VEIC_LEGA, " & _
             " DT_INIC_VIGE, " & _
             " DT_FIM_VIGE, " & _
             " CO_USUA_ULTI_ATLZ, " & _
             " CO_ETCA_TRAB_ULTI_ATLZ, " & _
             "  DH_ULTI_ATLZ " & _
             " ) VALUES ( " & _
             pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("CO_GRUP_VEIC_LEGA").Text & "," & _
             pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("TP_BKOF").Text & "," & _
             "'" & pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("NO_GRUP_VEIC_LEGA").Text & "'," & _
             fgDtXML_To_Oracle(pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("DT_INIC_VIGE").Text) & "," & _
             fgDtXML_To_Oracle(pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("DT_FIM_VIGE").Text) & "," & _
             "'" & fgObterUsuarioRede & "'," & _
             "'" & fgObterEstacaoTrabalhoUsuario & "'," & _
             "SYSDATE)"

    fgExecuteSQL strSQL
     
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Alterar um registro na tabela A8.TB_GRUP_VEIC_LEGA

Private Function flAlterar(ByRef pxmlGrupoVeiculoLegal As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    flVerificarDescicaoIgual "Alterar", _
                             pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("TP_BKOF").Text, _
                             pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("CO_GRUP_VEIC_LEGA").Text, _
                             pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("NO_GRUP_VEIC_LEGA").Text

    flVerificaVigenciaVeiculoLegal CLng(pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("CO_GRUP_VEIC_LEGA").Text), _
                                   pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("DT_FIM_VIGE").Text
            
    strSQL = " UPDATE A8.TB_GRUP_VEIC_LEGA " & _
            "  SET    NO_GRUP_VEIC_LEGA       = '" & pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("NO_GRUP_VEIC_LEGA").Text & "'," & _
            "         TP_BKOF                 = " & pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("TP_BKOF").Text & "," & _
            "         DT_INIC_VIGE            = " & fgDtXML_To_Oracle(pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("DT_INIC_VIGE").Text) & "," & _
            "         DT_FIM_VIGE             = " & fgDtXML_To_Oracle(pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("DT_FIM_VIGE").Text) & "," & _
            "         CO_USUA_ULTI_ATLZ       = '" & fgObterUsuarioRede & "'," & _
            "         CO_ETCA_TRAB_ULTI_ATLZ  = '" & fgObterEstacaoTrabalhoUsuario & "'," & _
            "         DH_ULTI_ATLZ            =  SYSDATE " & _
            " WHERE   CO_GRUP_VEIC_LEGA       = " & pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("CO_GRUP_VEIC_LEGA").Text & _
            " AND     DH_ULTI_ATLZ            = " & fgDtHrXML_To_Oracle(pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)
    
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

'Recuperar um registro na tabela A8.TB_GRUP_VEIC_LEGA

Public Function Ler(ByVal plngCO_GRUP_VEIC_LEGA As Long) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_GRUP_VEIC_LEGA, " & _
            "           NO_GRUP_VEIC_LEGA, " & _
            "           TP_BKOF, " & _
            "           DT_INIC_VIGE, " & _
            "           DT_FIM_VIGE, " & _
            "           CO_USUA_ULTI_ATLZ, " & _
            "           CO_ETCA_TRAB_ULTI_ATLZ, " & _
            "           DH_ULTI_ATLZ " & _
            " FROM      A8.TB_GRUP_VEIC_LEGA " & _
            " WHERE     CO_GRUP_VEIC_LEGA = " & plngCO_GRUP_VEIC_LEGA
    
    Ler = fgQueryXMLLer("Grupo_GrupoVeiculoLegal", strSQL, App.EXEName & "." & TypeName(Me))
    
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

'Salvar um registro (Incluir,Alterar,Exluir)

Public Function Salvar(ByVal pxmlGrupoVeiculoLegal As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
       
    If pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pxmlGrupoVeiculoLegal
    ElseIf pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        flAlterar pxmlGrupoVeiculoLegal
    ElseIf pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        flExcluir pxmlGrupoVeiculoLegal
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

'Exluir um registro na tabela A8.TB_GRUP_VEIC_LEGA

Private Function flExcluir(ByRef pxmlGrupoVeiculoLegal As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " DELETE FROM A8.TB_GRUP_VEIC_LEGA " & _
             " WHERE CO_GRUP_VEIC_LEGA  = " & pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("CO_GRUP_VEIC_LEGA").Text & _
             " AND DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pxmlGrupoVeiculoLegal.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)

    If fgExecuteSQL(strSQL) = 0 Then
        '4002-Registro foi exccluido por outro usuário
        lngCodigoErroNegocio = 4002
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Inserir o registro de grupo Padrão para cada tipo de backoffice na tabela A8.TB_GRUP_VEIC_LEGA

Private Function flInserirGrupoPadrao() As Boolean

Dim rsDisconect                             As ADODB.Recordset
Dim xmlGrupoVeiculoLegal                    As MSXML2.DOMDocument40
Dim rsGrupoVeiculoLegal                     As ADODB.Recordset
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = "Select Count(*) QtdeGrupoVeiculoLegal From A8.TB_GRUP_VEIC_LEGA "
    Set rsDisconect = fgQuerySQL(strSQL)
    'Somente inserir Grupo Padrão para cada Tipo de BackOffice se não houver nenhum Grupo na Tabela
    If rsDisconect!QtdeGrupoVeiculoLegal > 0 Then
        flInserirGrupoPadrao = False
        Exit Function
    End If
    
    strSQL = "Select TP_BKOF from A8.TB_TIPO_BKOF"
    Set rsDisconect = fgQuerySQL(strSQL)

    Do Until rsDisconect.EOF

        strSQL = "SELECT " & vbCrLf & _
                 "   max(CO_GRUP_VEIC_LEGA)+1 As CO_GRUP_VEIC_LEGA " & vbCrLf & _
                 "FROM " & vbCrLf & _
                 "   A8.TB_GRUP_VEIC_LEGA "
    
        Set rsGrupoVeiculoLegal = CreateObject("ADODB.Recordset")
        Set rsGrupoVeiculoLegal = fgQuerySQL(strSQL)
        
        Set xmlGrupoVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
        xmlGrupoVeiculoLegal.loadXML (ObterPropriedades)
        
        With xmlGrupoVeiculoLegal.documentElement
            .selectSingleNode("//@Operacao").Text = "Incluir"
            .selectSingleNode("//DT_INIC_VIGE").Text = Format(fgDataHoraServidor(Data), "YYYYMMDD")
            .selectSingleNode("//DT_FIM_VIGE").Text = vbNullString
            .selectSingleNode("//TP_BKOF").Text = rsDisconect.Fields!TP_BKOF
            .selectSingleNode("//CO_GRUP_VEIC_LEGA").Text = IIf(IsNull(rsGrupoVeiculoLegal("CO_GRUP_VEIC_LEGA").Value), 1, rsGrupoVeiculoLegal("CO_GRUP_VEIC_LEGA").Value)
            .selectSingleNode("//NO_GRUP_VEIC_LEGA").Text = strGRUPO_VEIC_LEGA_PADRAO
        End With
        
        flIncluir xmlGrupoVeiculoLegal
        
        rsDisconect.MoveNext
        
    Loop
        
    flInserirGrupoPadrao = True
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function

ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificaExistenciaGruposPadroes Function", lngCodigoErroNegocio, intNumeroSequencialErro)
        
End Function

'Verifica se já existe um grupo de veiculo legal com o mesma descrição

Private Function flVerificarDescicaoIgual(ByVal pstrOperacao As String, _
                                          ByVal penumTipoBackOffice As enumTipoBackOffice, _
                                          ByVal plngCodigoGrupoVeiculoLegal As Long, _
                                          ByVal pstrNomeGrupoVeiculoLegal As String)

Dim strSQL                                  As String
Dim rsGrupoVeiculoLeagal                    As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " Select       CO_GRUP_VEIC_LEGA,        " & vbNewLine & _
             "              NO_GRUP_VEIC_LEGA         " & vbNewLine & _
             " From      A8.TB_GRUP_VEIC_LEGA         " & vbNewLine & _
             " Where  UPPER(NO_GRUP_VEIC_LEGA) = '" & UCase(Trim(pstrNomeGrupoVeiculoLegal)) & "'" & vbNewLine & _
             "   And        TP_BKOF            = " & penumTipoBackOffice

    Set rsGrupoVeiculoLeagal = fgQuerySQL(strSQL)
    
    If Not rsGrupoVeiculoLeagal.EOF Then
        If pstrOperacao = "Incluir" Then
            '4073 - Já Existe um Grupo de Veículo Legal com este nome, para este Tipo de BackOffice
            lngCodigoErroNegocio = 4073
            GoTo ErrorHandler
        ElseIf pstrOperacao = "Alterar" And rsGrupoVeiculoLeagal.Fields!CO_GRUP_VEIC_LEGA <> plngCodigoGrupoVeiculoLegal Then
            '4073 - Já Existe um Grupo de Veículo Legal com este nome, para este Tipo de BackOffice
            lngCodigoErroNegocio = 4073
            GoTo ErrorHandler
        End If
    End If
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificarDescicaoIgual Function", lngCodigoErroNegocio, intNumeroSequencialErro, pstrNomeGrupoVeiculoLegal)
        
End Function
