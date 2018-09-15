VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Objeto responsável pela rastreabilidade de erros ocorridos no sistema.

Option Explicit

'Variaveis de Erro
Private ErrLastDllError                     As Long
Private ErrHelpContext                      As Long
Private ErrHelpFile                         As String

Private objContext                          As COMSVCSLib.ObjectContext
Private objSecurityCallContext              As COMSVCSLib.SecurityCallContext

Public Enum enumTipoErro
    ErroNegocio = 1
    ErroInterno = 2
End Enum

Implements COMSVCSLib.ObjectControl

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
    Set objSecurityCallContext = COMSVCSLib.GetSecurityCallContext
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
    Set objSecurityCallContext = Nothing
End Sub

'Adiciona um erro de negócio e sua descrição em um XML
Public Sub AdicionaErroNegocio(ByRef xmlDOMErro As MSXML2.DOMDocument40, _
                               ByVal lngErroNegocio As Long, _
                               Optional ByVal lngCodigoJustificativa As Long = 0)
                            
Dim objXMLElement                           As MSXML2.IXMLDOMElement
Dim strDescrErroNegocio                     As String
                            
On Error GoTo ErrorHandler

    If xmlDOMErro.xml = Empty Then
        fgAppendNode xmlDOMErro, "", "Erro", ""
    End If

    Set objXMLElement = xmlDOMErro.createElement("Grupo_ErrorInfo")
    
    objXMLElement.appendChild xmlDOMErro.createElement("Number")
    objXMLElement.selectSingleNode("Number").Text = lngErroNegocio
    
    objXMLElement.appendChild xmlDOMErro.createElement("Description")
    objXMLElement.selectSingleNode("Description").Text = ObterDescErroNegocio(CStr(lngErroNegocio))
    
    If lngCodigoJustificativa <> 0 Then
        objXMLElement.appendChild xmlDOMErro.createElement("CodigoJustificativa")
        objXMLElement.selectSingleNode("CodigoJustificativa").Text = lngCodigoJustificativa
    End If
    
    xmlDOMErro.selectSingleNode("//Erro").appendChild objXMLElement
    
    Set objXMLElement = Nothing
    
    Exit Sub

ErrorHandler:
    Set objXMLElement = Nothing
    Me.RaiseError App.EXEName, TypeName(Me), "AdicionaColecaoErrosNegocio", 0, Err.Number, Err.Source, Err.Description
    
End Sub

'Recuperar o erro de negócio

Public Sub RaiseError(ByVal psComponente As String, _
                      ByVal psClasse As String, _
                      ByVal psMetodo As String, _
                      ByRef plCodigoErroNegocio As Long, _
                      ByRef ErrNumber As Long, _
                      ByRef ErrSource As String, _
                      ByRef ErrDescription As String, _
             Optional ByRef piNumeroSequencialErro As Integer = 0, _
             Optional ByVal psComplemento As String = "", _
             Optional ByRef objErro As ErrObject)

Dim objDomErro                              As MSXML2.DOMDocument40
Dim lbErroOrigem                            As Boolean
    
    ErrNumber = objErro.Number
    ErrDescription = objErro.Description
    ErrSource = objErro.Source
    ErrLastDllError = objErro.LastDllError
    ErrHelpContext = objErro.HelpContext
    ErrHelpFile = objErro.HelpFile
    
    
On Error GoTo ErrHandler
    
    Set objDomErro = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not objDomErro.loadXML(ErrDescription) Then
        Call fgAppendNode(objDomErro, "", "Erro", "")
        lbErroOrigem = True
    End If
    
    If plCodigoErroNegocio = 0 Then
        'Erro Run-time interno
        If ErrNumber = 0 Then
            If lbErroOrigem Then
                Call GetErroUsoIndevido(objDomErro, psComponente & " - " & psClasse & " - " & psMetodo)
                Call fgAppendNode(objDomErro, "Erro", "Repet_Origem", "")
                
                ErrNumber = vbObjectError + 513
                ErrSource = psComponente & " - " & psClasse & " - " & psMetodo
            End If
        Else
            If lbErroOrigem Then
                Call GetVBErrorXML(objDomErro, ErrNumber, ErrSource, ErrDescription)
                Call GetContextXML(objDomErro)
                Call fgAppendNode(objDomErro, "Erro", "Repet_Origem", "")
            End If
            
        End If
    Else
        If lbErroOrigem Then
            Call GetErroNegocioXML(objDomErro, plCodigoErroNegocio, psComponente & " - " & psClasse & " - " & psMetodo)
            Call fgAppendNode(objDomErro, "Erro", "Repet_Origem", "")
        End If
                        
        ErrNumber = vbObjectError + 513 + plCodigoErroNegocio
        ErrSource = psComponente & " - " & psClasse & " - " & psMetodo
    End If
    
    If lbErroOrigem Then
        If piNumeroSequencialErro <> 0 Then
            fgAppendNode objDomErro, "Grupo_ErrorInfo", "Sequencial", piNumeroSequencialErro
        End If
    End If
    
    Call fgAppendNode(objDomErro, "Repet_Origem", "Grupo_Origem", "")
    Call fgAppendNode(objDomErro, "Grupo_Origem", "Origem", psComponente & " - " & psClasse & " - " & psMetodo, "Repet_Origem")
    Call fgAppendNode(objDomErro, "Grupo_Origem", "Complemento", psComplemento, "Repet_Origem")
        
    ErrDescription = objDomErro.xml
            
    Set objDomErro = Nothing
        
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, psComponente & " - " & psClasse & " - " & psMetodo, Err.Description
End Sub

'Obter a descrição do erro de negócio ou erro BACEN , tabelas A8.TB_MESG_ERRO e A8.TB_ERRO_BACEN

Public Function ObterDescErroNegocio(ByVal psCodigoErroNegocio As String) As String

Dim objConsulta                             As A6A7A8CA.clsConsulta
Dim rsErro                                  As ADODB.Recordset
Dim strSQL                                  As String

On Error GoTo ErrHandler
    
    Set objConsulta = CreateObject("A6A7A8CA.clsConsulta")
    
    'Verifica se é um erro de Negócio Interno da aplicação
    If IsNumeric(psCodigoErroNegocio) Then
        strSQL = " SELECT DE_ERRO ERRO " & _
                 "   FROM A8.TB_MESG_ERRO " & _
                "   WHERE CO_ERRO = " & psCodigoErroNegocio
        
        Set rsErro = objConsulta.QuerySQL(strSQL, 1)
    
    Else '...ou se é um erro proveniente do BACEN
        
        strSQL = " SELECT DE_ERRO_BACEN ERRO " & _
                 "   FROM A8.TB_ERRO_BACEN " & _
                 "  WHERE CO_ERRO_BACEN = '" & psCodigoErroNegocio & "'"
        
        Set rsErro = objConsulta.QuerySQL(strSQL, 1)
        
    End If
                               
    If rsErro.EOF Then
        ObterDescErroNegocio = "Erro não cadastrado na tabela de erros"
    Else
        ObterDescErroNegocio = rsErro.fields("Erro")
    End If
    
    Set rsErro = Nothing
    Set objConsulta = Nothing
    
    Exit Function

ErrHandler:
    
    Set rsErro = Nothing
    Set objConsulta = Nothing
    
    Err.Raise Err.Number, "CA-clsLogErro-ObterDescErroNegocio", Err.Description
        
End Function


Private Function GetContextXML(ByRef objDOMDocument As MSXML2.DOMDocument40) As String

    If objContext Is Nothing Then
        Exit Function
    End If
        
    Call fgAppendNode(objDOMDocument, "Erro", "Grupo_ObjectContext", "")
    
    Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "Time", Now)
    'Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "Description", "XML")
    Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "Application", App.Title)
    Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "ApplicationVersion", App.Major & "." & App.Minor & "." & App.Revision)
    'Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "UnattendedAndRetained", -(App.UnattendedApp And App.RetainedProject))
    'Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "Thread", App.ThreadID)

    Exit Function

    If objContext Is Nothing Then
        Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "Valid", "0")
    Else
        Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "Valid", "1")
        With objContext

            With .ContextInfo
                Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "ActivityId", .GetActivityId)
                Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "ContextId", .GetContextId)
                Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "IsInTransaction", .IsInTransaction)

                If .IsInTransaction Then
                    Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "TransactionId", .GetTransactionId)
                End If

            End With

            Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "IsSecurityEnabled", .IsSecurityEnabled)
            Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "ItemCount", .Count)

            If .IsSecurityEnabled Then
                With .Security
                    Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "DirectCallerName", .GetDirectCallerName)
                    Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "DirectCreatorName", .GetDirectCreatorName)
                    Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "OriginalCallerName", .GetOriginalCallerName)
                    Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "OriginalCreatorName", .GetOriginalCreatorName)
                End With
            End If

        End With
        
        If objSecurityCallContext.IsSecurityEnabled Then
            With objSecurityCallContext

                Dim i As Integer
                Dim oCallers As SecurityCallers
                
                Set oCallers = .Item("Callers")
                
                Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "CallerCount", oCallers.Count)
                
                If oCallers.Count > 0 Then
                    Call fgAppendNode(objDOMDocument, "Grupo_ObjectContext", "Repeat_Caller", "")
                End If
                
                For i = 1 To oCallers.Count
                    With oCallers.Item(i)
                        
                        Call fgAppendNode(objDOMDocument, "Repeat_Caller", "Grupo_Caller", "")
                        
                        Call fgAppendNode(objDOMDocument, "Grupo_Caller", "AccountName", .Item("AccountName"))
                        Call fgAppendNode(objDOMDocument, "Grupo_Caller", "AuthenticationService", .Item("AuthenticationService"))
                        Call fgAppendNode(objDOMDocument, "Grupo_Caller", "ImpersonationLevel", _
                            .Item(" ImpersonationLevel"))
                        Call fgAppendNode(objDOMDocument, "Grupo_Caller", "AuthenticationLevel", _
                            .Item("AuthenticationLevel"))
                    End With
                Next i
                Set oCallers = Nothing
            End With
        End If

    End If

End Function

Private Function GetVBErrorXML(ByRef objDOMDocument As MSXML2.DOMDocument40, _
                               ByVal ErrNumber As Long, _
                               ByVal ErrSource As String, _
                               ByVal ErrDescription As String) As String

Dim lsNomeComputador                        As String
Dim objWinNTSystemInfo                      As ActiveDs.WinNTSystemInfo
    
    Set objWinNTSystemInfo = New ActiveDs.WinNTSystemInfo
    lsNomeComputador = objWinNTSystemInfo.ComputerName
    Set objWinNTSystemInfo = Nothing
    
    Call fgAppendNode(objDOMDocument, "Erro", "Grupo_ErrorInfo", "")
    
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "Number", CStr(ErrNumber))
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "Description", ErrDescription)
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "ComputerName", lsNomeComputador)
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "Source", ErrSource)
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "ErrorType", enumTipoErro.ErroInterno)
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "Time", Now)

End Function

Private Function GetErroNegocioXML(ByRef objDOMDocument As MSXML2.DOMDocument40, _
                                   ByVal plCodigoErroNegocio As String, _
                                   ByVal ErrSource As String) As String

Dim lsNomeComputador                        As String
Dim objWinNTSystemInfo                      As ActiveDs.WinNTSystemInfo
    
    Set objWinNTSystemInfo = New ActiveDs.WinNTSystemInfo
    lsNomeComputador = objWinNTSystemInfo.ComputerName
    Set objWinNTSystemInfo = Nothing
    
    Call fgAppendNode(objDOMDocument, "Erro", "Grupo_ErrorInfo", "")
    
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "Number", CStr(plCodigoErroNegocio))
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "Description", ObterDescErroNegocio(plCodigoErroNegocio))
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "ComputerName", lsNomeComputador)
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "Source", ErrSource)
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "ErrorType", enumTipoErro.ErroNegocio)
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "Time", Now)

End Function

Private Function GetErroUsoIndevido(ByRef objDOMDocument As MSXML2.DOMDocument40, _
                                    ByVal ErrSource As String) As String
 
Dim lsNomeComputador                        As String
Dim objWinNTSystemInfo                      As ActiveDs.WinNTSystemInfo
    
    Set objWinNTSystemInfo = New ActiveDs.WinNTSystemInfo
    lsNomeComputador = objWinNTSystemInfo.ComputerName
    Set objWinNTSystemInfo = Nothing
    
    Call fgAppendNode(objDOMDocument, "Erro", "Grupo_ErrorInfo", "")
    
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "Number", 0)
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "Description", "Erro não tratado. A mensagem foi enviada para a fila de erro para um reprocessamento comandado pelo usuário!")
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "ComputerName", lsNomeComputador)
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "Source", ErrSource)
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "ErrorType", enumTipoErro.ErroInterno)
    Call fgAppendNode(objDOMDocument, "Grupo_ErrorInfo", "Time", Now)

End Function

'baseado na tabela PJ.TB_ERRO_SPB
Public Function ObterDescricaoErroNZPKPJ(ByVal pstrCodigoErro As String) As String

    On Error GoTo ErrorHandler

    Select Case pstrCodigoErro
        Case "254": ObterDescricaoErroNZPKPJ = " - ISPB não possui empresa cadastrada"
        Case "5001": ObterDescricaoErroNZPKPJ = " - Data da Remessa Inválida"
        Case "5002": ObterDescricaoErroNZPKPJ = " - Data da Remessa não é a data Vigente"
        Case "5003": ObterDescricaoErroNZPKPJ = " - Código da Empresa Inválido"
        Case "5004": ObterDescricaoErroNZPKPJ = " - Código do Sistema Inválido"
        Case "5005": ObterDescricaoErroNZPKPJ = " - Código da Moeda Inválido"
        Case "5006": ObterDescricaoErroNZPKPJ = " - Sigla do Sistema Inválido"
        Case "5007": ObterDescricaoErroNZPKPJ = " - Controle Remessa - Data Corrente Inválida"
        Case "5008": ObterDescricaoErroNZPKPJ = " - Controle Remessa - Sigla Sistema não Informada"
        Case "5009": ObterDescricaoErroNZPKPJ = " - Controle Remessa - Número Sequência Inválido"
        Case "5010": ObterDescricaoErroNZPKPJ = " - Valor Lançamento Inválido"
        Case "5011": ObterDescricaoErroNZPKPJ = " - Número Sequencia Mensagem Inválido"
        Case "5012": ObterDescricaoErroNZPKPJ = " - Quantidade de Ocorrências Mensagem Inválido"
        Case "5013": ObterDescricaoErroNZPKPJ = " - Identificador da Mensagem (Correlation ID) não Informado"
        Case "5014": ObterDescricaoErroNZPKPJ = " - Identificador da Mensagem (Correlation ID) Inválido"
        Case "5015": ObterDescricaoErroNZPKPJ = " - Identificador do Formato da Mensagem Inválido"
        Case "5016": ObterDescricaoErroNZPKPJ = " - Código da Mensagem Inválido"
        Case "5017": ObterDescricaoErroNZPKPJ = " - Controle Remessa - Sigla Sistema Inválido"
        Case "5018": ObterDescricaoErroNZPKPJ = " - Indicador de Impacto na Reserva Inválido"
        Case "5073": ObterDescricaoErroNZPKPJ = " - Número de repetições da mensagem e maior que o cadastrado"
        Case "5100": ObterDescricaoErroNZPKPJ = " - Documento XML inválido"
        Case "5101": ObterDescricaoErroNZPKPJ = " - Documento não válido - SISMSG"
        Case "5102": ObterDescricaoErroNZPKPJ = " - Erro no Formato XML"
        Case "5103": ObterDescricaoErroNZPKPJ = " - Erro no formato XML do banco de dados"
        Case "5104": ObterDescricaoErroNZPKPJ = " - Mensagem não encontrada no banco de dados"
        Case "5105": ObterDescricaoErroNZPKPJ = " - XML da mensagem inválido"
        Case "5106": ObterDescricaoErroNZPKPJ = " - XML do banco de dados inválido"
        Case "5107": ObterDescricaoErroNZPKPJ = " - Tag requerida inexistente"
        Case "5108": ObterDescricaoErroNZPKPJ = " - Tag com tamanho inválido"
        Case "5109": ObterDescricaoErroNZPKPJ = " - Tag numérica com conteúdo alfanumérico"
        Case "5110": ObterDescricaoErroNZPKPJ = " - Tag sem conteúdo"
        Case "5112": ObterDescricaoErroNZPKPJ = " - BCMSG Inválido"
        Case "5114": ObterDescricaoErroNZPKPJ = " - Mensagem sem conteúdo"
        Case "5115": ObterDescricaoErroNZPKPJ = " - Mensagem R1 recebida sem requisição"
        Case "5116": ObterDescricaoErroNZPKPJ = " - Domínio Inválido para Tipo Tag"
        Case "5117": ObterDescricaoErroNZPKPJ = " - Indicador de Formato Mensagem não cadastrado para Sistema - Mensagem"
        Case "5118": ObterDescricaoErroNZPKPJ = " - Formato não cadastrado na tabela de sistema mensagem"
        Case "5120": ObterDescricaoErroNZPKPJ = " - Erro Conversão de XML para Copy Fixo"
        Case "5137": ObterDescricaoErroNZPKPJ = " - Mensagem Rejeitada por Regra Impeditiva de  WorkFLow"
        Case "5138": ObterDescricaoErroNZPKPJ = " - Regras do Saldo Reserva com XML incorreto"
        Case "5156": ObterDescricaoErroNZPKPJ = " - Mensagem Cancelada - Piloto."
        Case "5157": ObterDescricaoErroNZPKPJ = " - Mensagem Cancelada."
        Case "5158": ObterDescricaoErroNZPKPJ = " - Mensagem não esta em aprovação."
        Case "5159": ObterDescricaoErroNZPKPJ = " - Usuário que envia não pode aprovar a mensagem."
        Case "5160": ObterDescricaoErroNZPKPJ = " - Cancelamento/Aprovação já realizado por este usuário."
        Case "5161": ObterDescricaoErroNZPKPJ = " - Data da mensagem inválida."
        Case "5162": ObterDescricaoErroNZPKPJ = " - Usuário não pertence ao grupo de envio."
        Case "5172": ObterDescricaoErroNZPKPJ = " - Mensagem Duplicada no Controle de Envio"
        Case "5173": ObterDescricaoErroNZPKPJ = " - Erro Fila sem Mensagem - (Relacionado Saldo)"
        Case "5178": ObterDescricaoErroNZPKPJ = " - Mensagem Cancelada - Análise GOE"
        Case Else: ObterDescricaoErroNZPKPJ = vbNullString
    End Select

Exit Function
ErrorHandler:
    Err.Raise Err.Number, "A6A7A8CA-clsLogErro-ObterDescricaoErroNZPKPJ", Err.Description
End Function

