VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControleAcessDado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável pelo mapeamento de informações (DE/PARA), bem como o
'' controle de acesso entre os sistemas:
''      MBS     (Módulo Básico de Segurança)    e
''      SLCC    (Sistema de Liquidação e Controle de Câmaras)

'
Option Explicit
'
'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private objContext                          As COMSVCSLib.ObjectContext

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

'' Consultar uma associação DE/PARA já efetuada para outro código MBS, através dos
'' parâmetros: código e tipo de informação fornecidos
Private Sub flVerificarAssociacao(ByVal pintTP_INFO As Integer, ByVal plngCO_INFO As Long)

Dim strSQL                                  As String
Dim strLer                                  As String
Dim rsLer                                   As ADODB.Recordset
Dim strComplemento                          As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_GRUP_GS,             " & vbNewLine & _
             "          NO_GRUP_GS              " & vbNewLine & _
             " FROM     A8.TB_CTRL_ACES_DADO    " & vbNewLine & _
             " WHERE    CO_INFO = " & plngCO_INFO & _
             " AND      TP_INFO = " & pintTP_INFO
    
    Set rsLer = fgQuerySQL(strSQL)
    
    If Not rsLer.EOF Then
       '3033 - Informação já está associada ao Código de Grupo GS:
       lngCodigoErroNegocio = 3033
       strComplemento = rsLer("CO_GRUP_GS").Value & " - " & rsLer("NO_GRUP_GS").Value
       GoTo ErrorHandler
    End If

    Exit Sub
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificarAssociacao", lngCodigoErroNegocio, intNumeroSequencialErro, strComplemento)


End Sub

'' Consultar todos as informações DE/PARA cadastradas por tipo de informação,
'' sendo:
''      Grupo de Veículo Legal
''      Grupo de Usuário
''      Tipo de BackOffice              ou
''      Local de Liquidação
''
''      ou Todas as informações
Public Function LerTodos(Optional ByVal penumTipoInformacao As enumTipoInformacao, _
                         Optional ByVal pstrLerTodos As String = "N", _
                         Optional ByVal pstrVerificaControleAcesDado As String = "N", _
                         Optional ByRef prsControleAcessDado As ADODB.Recordset, _
                         Optional ByVal penumTipoInformacaoAux As enumTipoInformacao) As String


Dim strSQL                                  As String

Dim strSQL01                                As String
Dim strSQL02                                As String
Dim strSQL03                                As String
Dim strSQL04                                As String
Dim strSQL05                                As String

Dim strSQLUnionAll                          As String

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   A.CO_GRUP_GS as CO_GRUP_GS, " & vbNewLine & _
             "          A.NO_GRUP_GS,               " & vbNewLine & _
             "          A.CO_INFO,                  " & vbNewLine & _
             "          A.TP_INFO,                  " & vbNewLine & _
             "          A.CO_USUA_ULTI_ATLZ,        " & vbNewLine & _
             "          A.CO_ETCA_TRAB_ULTI_ATLZ,   "

    If pstrLerTodos = "N" Then
    
        strSQL = strSQL & "          A.DH_ULTI_ATLZ,           "
    
        If penumTipoInformacao <> 0 Then
           strSQL = strSQL & "          B.DT_INIC_VIGE,         " & vbNewLine & _
                             "          B.DT_FIM_VIGE,          " & vbNewLine & _
                             "          '' AS TP_VIGE,          "
            Select Case penumTipoInformacao
                   Case enumTipoInformacao.GrupoVeiculoLegal
                        strSQL01 = "       B.NO_GRUP_VEIC_LEGA As DE_INFO  " & vbNewLine & _
                                   " FROM  A8.TB_CTRL_ACES_DADO A,         " & vbNewLine & _
                                   "       A8.TB_GRUP_VEIC_LEGA B          " & vbNewLine & _
                                   " WHERE B.CO_GRUP_VEIC_LEGA = A.CO_INFO " & vbNewLine & _
                                   " AND   B.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & vbCrLf & _
                                   " AND  (B.DT_FIM_VIGE    IS  NULL       " & vbCrLf & _
                                   " OR    B.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
                                          
                        strSQL = strSQL & strSQL01
                        
                   Case enumTipoInformacao.GrupoUsuario
                        strSQL02 = "          B.NO_GRUP_USUA  As DE_INFO " & vbNewLine & _
                                   " FROM     A8.TB_CTRL_ACES_DADO A,    " & vbNewLine & _
                                   "          A8.TB_GRUP_USUA B          " & vbNewLine & _
                                   " WHERE    B.CO_GRUP_USUA = A.CO_INFO " & vbNewLine & _
                                   " AND   B.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & vbCrLf & _
                                   " AND  (B.DT_FIM_VIGE    IS  NULL       " & vbCrLf & _
                                   " OR    B.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
                        
                        strSQL = strSQL & strSQL02
                                          
                   
                   Case enumTipoInformacao.TipoBackOffice
                        strSQL03 = "          B.DE_BKOF As DE_INFO    " & vbNewLine & _
                                   " FROM    A8.TB_CTRL_ACES_DADO A, " & vbNewLine & _
                                   "         A8.TB_TIPO_BKOF B       " & vbNewLine & _
                                   " WHERE   B.TP_BKOF = A.CO_INFO   " & vbNewLine & _
                                   " AND   B.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & vbCrLf & _
                                   " AND  (B.DT_FIM_VIGE    IS  NULL       " & vbCrLf & _
                                   " OR    B.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
                                          
                        strSQL = strSQL & strSQL03
                   
                   Case enumTipoInformacao.LocalLiquidacao
                        strSQL04 = "         B.DE_LOCA_LIQU As DE_INFO  " & vbNewLine & _
                                   " FROM     A8.TB_CTRL_ACES_DADO A,    " & vbNewLine & _
                                   "          A8.TB_LOCAL_LIQUIDACAO B      " & vbNewLine & _
                                   " WHERE    B.CO_LOCA_LIQU = A.CO_INFO " & vbNewLine & _
                                   " AND      B.CO_EMPR_FUSI = 1 "
                        
                        strSQL = strSQL & strSQL04
                                                              
            End Select
            
            If penumTipoInformacao = TodasInformacoes Then
                
                strSQLUnionAll = vbCrLf & " Union All " & vbCrLf
                
                strSQL05 = strSQL05 & LerTodos(GrupoUsuario, , , , TodasInformacoes) & _
                                strSQLUnionAll & _
                                LerTodos(GrupoVeiculoLegal, , , , TodasInformacoes) & _
                                strSQLUnionAll & _
                                LerTodos(LocalLiquidacao, , , , TodasInformacoes) & _
                                strSQLUnionAll & _
                                LerTodos(TipoBackOffice, , , , TodasInformacoes)
                                
                strSQL = strSQL05
                
            Else
            
                strSQL = strSQL & " AND A.TP_INFO = " & penumTipoInformacao
                
            End If
        Else
            strSQL = strSQL & vbNewLine & _
                            "          '' AS TP_VIGE           " & vbNewLine & _
                            " FROM     A8.TB_CTRL_ACES_DADO A  "
        End If
    
    Else
       strSQL = strSQL & "          A.DH_ULTI_ATLZ           " & vbNewLine & _
                         " FROM     A8.TB_CTRL_ACES_DADO A   " & vbNewLine & _
                         " WHERE A.TP_INFO = " & penumTipoInformacao
       
    End If
    
    If Not penumTipoInformacaoAux > 0 Then
        If pstrVerificaControleAcesDado = "N" Then
           strSQL = strSQL & vbCrLf & " ORDER BY CO_GRUP_GS "
        Else
           strSQL = strSQL & vbCrLf & " ORDER BY A.CO_INFO "
        End If
    End If
    
    If penumTipoInformacaoAux > 0 Then
        LerTodos = strSQL
        Exit Function
    End If
    
    If prsControleAcessDado Is Nothing Then
       LerTodos = fgQueryXMLLerTodos("ControleAcessDado", strSQL, App.EXEName & "." & TypeName(Me))
    Else
       Set prsControleAcessDado = fgQuerySQL(strSQL)
       LerTodos = ""
    End If

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a
'' estrutura da tabela:
''      TB_CTRL_ACES_DADO
Public Function ObterPropriedades() As String
Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_GRUP_GS,             " & vbNewLine & _
             "          NO_GRUP_GS,             " & vbNewLine & _
             "          CO_INFO,                " & vbNewLine & _
             "          TP_INFO,                " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ,      " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
             "          DH_ULTI_ATLZ,           " & vbNewLine & _
             "          '' AS TP_VIGE           " & vbNewLine & _
             " FROM     A8.TB_CTRL_ACES_DADO    "

    ObterPropriedades = fgPropriedades("Grupo_ControleAcessDado", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Controlar o roteamento e execução dos métodos privados, tais como:
''      Incluir
''      Alterar
''      Excluir
''      Ler     e
''      LerTodos
Public Function Executar(ByVal pxmlControleAcessDado As MSXML2.DOMDocument40) As String
Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlControleAcessDado.documentElement.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pxmlControleAcessDado)
        Case "Ler"
            Executar = Ler(pxmlControleAcessDado.documentElement.selectSingleNode("//CO_GRUP_GS").Text)
            
        Case "LerTodos"
              If pxmlControleAcessDado.documentElement.selectSingleNode("//TP_INFO").Text = "" Then
                 Executar = LerTodos()
              Else
                 Executar = LerTodos(pxmlControleAcessDado.documentElement.selectSingleNode("//TP_INFO").Text)
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
'' Promover a inclusão de um novo DE/PARA (MBS x SLCC)
Private Function flIncluir(ByRef pxmlControleAcessDado As MSXML2.DOMDocument40) As Boolean
Dim strSQL                                   As String
Dim strNO_GRUP_GS                            As String
Dim strCO_GRUP_GS                            As String

On Error GoTo ErrorHandler

    flVerificarAssociacao pxmlControleAcessDado.documentElement.selectSingleNode("TP_INFO").Text, _
                          pxmlControleAcessDado.documentElement.selectSingleNode("CO_INFO").Text

    If pxmlControleAcessDado.documentElement.selectSingleNode("NO_GRUP_GS").Text = "" Then
       strNO_GRUP_GS = "Null"
    Else
       strNO_GRUP_GS = "'" & pxmlControleAcessDado.documentElement.selectSingleNode("NO_GRUP_GS").Text & "'"
    End If
       
    If pxmlControleAcessDado.documentElement.selectSingleNode("CO_GRUP_GS").Text = "" Then
       strCO_GRUP_GS = "Null"
    Else
       strCO_GRUP_GS = pxmlControleAcessDado.documentElement.selectSingleNode("CO_GRUP_GS").Text
    End If

    strSQL = " INSERT INTO A8.TB_CTRL_ACES_DADO (" & _
             " CO_GRUP_GS, " & _
             " NO_GRUP_GS, " & _
             " CO_INFO, " & _
             " TP_INFO, " & _
             " CO_USUA_ULTI_ATLZ, " & _
             " CO_ETCA_TRAB_ULTI_ATLZ, " & _
             "  DH_ULTI_ATLZ " & _
             " ) VALUES ( " & _
             strCO_GRUP_GS & "," & _
             strNO_GRUP_GS & "," & _
             pxmlControleAcessDado.documentElement.selectSingleNode("CO_INFO").Text & "," & _
             pxmlControleAcessDado.documentElement.selectSingleNode("TP_INFO").Text & "," & _
             "'" & fgUsuarioRede & "'," & _
             "'" & fgEstacaoTrabalhoUsuario & "'," & _
             "SYSDATE)"

    fgExecuteSQL strSQL
     
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
'' Promover a alteração de um DE/PARA (MBS x SLCC) existente
Private Function flAlterar(ByRef pxmlControleAcessDado As MSXML2.DOMDocument40) As Boolean
Dim strSQL                                   As String
Dim strNO_GRUP_GS                            As String
Dim strCO_GRUP_GS                            As String

On Error GoTo ErrorHandler

    flVerificarAssociacao pxmlControleAcessDado.documentElement.selectSingleNode("TP_INFO").Text, _
                          pxmlControleAcessDado.documentElement.selectSingleNode("CO_INFO").Text

    If pxmlControleAcessDado.documentElement.selectSingleNode("NO_GRUP_GS").Text = "" Then
       strNO_GRUP_GS = "Null"
    Else
       strNO_GRUP_GS = "'" & pxmlControleAcessDado.documentElement.selectSingleNode("NO_GRUP_GS").Text & "'"
    End If
       
    If pxmlControleAcessDado.documentElement.selectSingleNode("CO_GRUP_GS").Text = "" Then
       strCO_GRUP_GS = "Null"
    Else
       strCO_GRUP_GS = pxmlControleAcessDado.documentElement.selectSingleNode("CO_GRUP_GS").Text
    End If
            
    strSQL = " UPDATE A8.TB_CTRL_ACES_DADO    " & _
            "  SET    NO_GRUP_GS              =  '" & pxmlControleAcessDado.documentElement.selectSingleNode("NO_GRUP_GS").Text & "'," & _
            "         CO_INFO                 =  " & pxmlControleAcessDado.documentElement.selectSingleNode("CO_INFO").Text & "," & _
            "         TP_INFO                 =  " & pxmlControleAcessDado.documentElement.selectSingleNode("TP_INFO").Text & "," & _
            "         CO_USUA_ULTI_ATLZ       = '" & fgUsuarioRede & "'," & _
            "         CO_ETCA_TRAB_ULTI_ATLZ  = '" & fgEstacaoTrabalhoUsuario & "'," & _
            "         DH_ULTI_ATLZ            =  SYSDATE " & _
            " WHERE   CO_GRUP_GS              = " & CLng(pxmlControleAcessDado.documentElement.selectSingleNode("CO_GRUP_GS").Text) & _
            " AND     DH_ULTI_ATLZ            = " & fgDtHrXML_To_Oracle(pxmlControleAcessDado.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)
    
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
'' Consultar todos os atributos de um DE/PARA (MBS x SLCC) específico
Public Function Ler(ByVal plngCO_GRUP_GS As Long) As String
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_GRUP_GS,             " & vbNewLine & _
             "          NO_GRUP_GS,             " & vbNewLine & _
             "          CO_INFO,                " & vbNewLine & _
             "          TP_INFO,                " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ,      " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
             "          DH_ULTI_ATLZ,           " & vbNewLine & _
             "          '' AS TP_VIGE           " & vbNewLine & _
             " FROM     A8.TB_CTRL_ACES_DADO    " & vbNewLine & _
             " WHERE    CO_GRUP_GS = " & plngCO_GRUP_GS
    
    
    Ler = fgQueryXMLLer("Grupo_ControleAcessDado", strSQL, App.EXEName & "." & TypeName(Me))
    
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
'' Controlar o roteamento e execução dos métodos privados de manipulação de dados,
'' tais como:
''      Incluir
''      Alterar e
''      Excluir
Public Function Salvar(ByVal pxmlControleAcessDado As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    If pxmlControleAcessDado.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pxmlControleAcessDado
    ElseIf pxmlControleAcessDado.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        flAlterar pxmlControleAcessDado
    ElseIf pxmlControleAcessDado.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        flExcluir pxmlControleAcessDado
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

'' Promover a exclusão de um DE/PARA (MBS x SLCC) existente
Private Function flExcluir(ByRef pxmlControleAcessDado As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " DELETE " & strSQL & _
             " FROM A8.TB_CTRL_ACES_DADO " & _
             " WHERE CO_GRUP_GS   = " & CLng(pxmlControleAcessDado.documentElement.selectSingleNode("CO_GRUP_GS").Text) & _
             " AND   DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pxmlControleAcessDado.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)

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

