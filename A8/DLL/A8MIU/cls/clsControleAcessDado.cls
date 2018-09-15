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

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito ao Controle de Acesso de Usuários / Segregação de Dados.

Option Explicit

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

' Este método aciona a leitura da descrição do tipo de backoffice a partir do código passado como parâmetro,
' na camada de negócios, e retorna para a camada de interface.

Public Function ObterDescricaoTipoBackoffice(ByVal strTipoBackoffice As String) As String

Dim objUsuario                              As A6A7A8.clsControleAcesso

    On Error GoTo ErrorHandler
    
    Set objUsuario = CreateObject("A6A7A8.clsControleAcesso")
    ObterDescricaoTipoBackoffice = objUsuario.ObterDescricaoTipoBackoffice(CLng(strTipoBackoffice))
    Set objUsuario = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

ErrorHandler:
    Set objUsuario = Nothing

    If Not objContext Is Nothing Then
        'ObjectContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDescricaoTipoBackOffice Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona a obtenção do tipo de backoffice a partir do código do usuário logado no sistema,
' na camada de negócios, e retorna para a camada de interface.
Public Function ObterTipoBackOfficeUsuario(ByRef pvntCodErro As Variant, _
                                           ByRef pvntMensagemErro As Variant) As String

Dim objUsuario                           As A6A7A8.clsControleAcesso

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objUsuario = CreateObject("A6A7A8.clsControleAcesso")
    
    ObterTipoBackOfficeUsuario = objUsuario.ObterTipoBackOfficeUsuarioLogado
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objUsuario = Nothing

Exit Function
ErrorHandler:

    Set objUsuario = Nothing

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoBackOffice Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura dos diversos itens associados à segregação de dados para usuários
' na camada de negócios, monta um XML com os resultados, e retorna para a camada de interface.

Public Function VerificaControleAcessDado(penumTipoInformacao As Integer, _
                                          ByRef pvntCodErro As Variant, _
                                          ByRef pvntMensagemErro As Variant) As String

Dim xmlOcorrencias                          As MSXML2.DOMDocument40
Dim blnExisteOcorrencia                     As Boolean
    
Dim objControleAcessDado                    As A8LQS.clsControleAcessDado
Dim objGrupoVeiculoLegal                    As A6A7A8.clsGrupoVeiculoLegal
Dim objGrupoUsuario                         As A6A7A8.clsGrupoUsuario
Dim objTipoBackOffice                       As A8LQS.clsTipoBackOffice
Dim objLocalLiquidacao                      As A6A7A8.clsLocalLiquidacao

Dim rsTipoInformacao                        As ADODB.Recordset
Dim rsControleAcesso                        As ADODB.Recordset

Dim strCodigoA8                             As String
Dim strCodigoA8Ant                          As String
Dim strDescricaoA8                          As String
Dim strDescricaoTipoInfo                    As String

Dim strOcorrencia                           As String
Dim intCodOcorrencia                        As Integer

Dim strCodInformacao                        As String
Dim strDesInformacao                        As String
    
On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    blnExisteOcorrencia = False
    
    Select Case penumTipoInformacao
           Case enumTipoInformacao.GrupoVeiculoLegal
           
                Set objGrupoVeiculoLegal = CreateObject("A6A7A8.clsGrupoVeiculoLegal")
                Set rsTipoInformacao = CreateObject("ADODB.Recordset")
                Call objGrupoVeiculoLegal.LerTodos("N", 0, "N", rsTipoInformacao)
                Set objGrupoVeiculoLegal = Nothing
                
                strCodigoA8 = "CO_GRUP_VEIC_LEGA"
                strDescricaoA8 = "NO_GRUP_VEIC_LEGA"
                strDescricaoTipoInfo = "Grupo de Veículo Legal"
           
           Case enumTipoInformacao.GrupoUsuario
           
                Set objGrupoUsuario = CreateObject("A6A7A8.clsGrupoUsuario")
                Set rsTipoInformacao = CreateObject("ADODB.Recordset")
                Call objGrupoUsuario.LerTodos("N", "N", rsTipoInformacao)
                Set objGrupoUsuario = Nothing
           
                strCodigoA8 = "CO_GRUP_USUA"
                strDescricaoA8 = "NO_GRUP_USUA"
                strDescricaoTipoInfo = "Grupo de Usuário"
           
           Case enumTipoInformacao.TipoBackOffice
           
                Set objTipoBackOffice = CreateObject("A8LQS.clsTipoBackOffice")
                Set rsTipoInformacao = CreateObject("ADODB.Recordset")
                Call objTipoBackOffice.LerTodos("N", "N", rsTipoInformacao)
                Set objTipoBackOffice = Nothing
           
                strCodigoA8 = "TP_BKOF"
                strDescricaoA8 = "DE_BKOF"
                strDescricaoTipoInfo = "Tipo BackOffice"
           
           Case enumTipoInformacao.LocalLiquidacao
           
                Set objLocalLiquidacao = CreateObject("A6A7A8.clsLocalLiquidacao")
                Set rsTipoInformacao = CreateObject("ADODB.Recordset")
                Call objLocalLiquidacao.LerTodos("N", "N", rsTipoInformacao)
                Set objLocalLiquidacao = Nothing
           
                strCodigoA8 = "CO_LOCA_LIQU"
                strDescricaoA8 = "DE_LOCA_LIQU"
                strDescricaoTipoInfo = "Local Liquidação"
    
    End Select

    Set objControleAcessDado = CreateObject("A8LQS.clsControleAcessDado")
    Set rsControleAcesso = CreateObject("ADODB.Recordset")
    Call objControleAcessDado.LerTodos(penumTipoInformacao, "S", "S", rsControleAcesso)
    Set objControleAcessDado = Nothing
    
    Set xmlOcorrencias = CreateObject("MSXML2.DOMDocument.4.0")
    
    fgAppendNode xmlOcorrencias, "", "Repeat_TipoInformacao", ""
    fgAppendNode xmlOcorrencias, "Repeat_TipoInformacao", "Grupo_TipoInformacao", ""
    fgAppendNode xmlOcorrencias, "Grupo_TipoInformacao", "TP_INFO", penumTipoInformacao, ""
    fgAppendNode xmlOcorrencias, "Grupo_TipoInformacao", "DE_INFO", strDescricaoTipoInfo, ""
    fgAppendNode xmlOcorrencias, "Grupo_TipoInformacao", "Repeat_Ocorrencias", ""

    'Se não existir Tipo de Informação então vai para o Loop do Controle de Acesso.
    If Not rsTipoInformacao.EOF Then
        Do Until rsTipoInformacao.EOF
           'Se não existir Controle de Acesso então vai para o Loop do Tipo de Informação.
           If Not rsControleAcesso.EOF Then
               Do Until rsControleAcesso.EOF
                
                  intCodOcorrencia = 0
               
                  If Not rsTipoInformacao.EOF Then
                    'Se existe em Ambas as Tabelas.
                    If rsControleAcesso("CO_INFO").Value = rsTipoInformacao(strCodigoA8).Value Then
                             
                       If rsTipoInformacao("DT_FIM_VIGE").Value <> Null Then
                          'Tipo Informação não está vigente mas sua associação está.
                          'Ação 03
                              
                          intCodOcorrencia = 3
                          strOcorrencia = strDescricaoTipoInfo & " não está vigente mas sua associação está."
                          
                          strCodInformacao = CStr(rsTipoInformacao(strCodigoA8).Value)
                          strDesInformacao = CStr(rsTipoInformacao(strDescricaoA8).Value)
                              
                       End If
                       
                       rsTipoInformacao.MoveNext
                       rsControleAcesso.MoveNext
                       
                    ElseIf rsControleAcesso("CO_INFO").Value < rsTipoInformacao(strCodigoA8).Value Then
                        'Tipo Informação não existe mas esta associado a um Grupo do Sistema MBS.
                        'Ação 01
                        
                        intCodOcorrencia = 1
                        strOcorrencia = strDescricaoTipoInfo & " não existe mas esta associado a um Grupo do Sistema MBS."
                        
                        strCodInformacao = rsControleAcesso("CO_INFO").Value
                        strDesInformacao = ""
                        
                        rsControleAcesso.MoveNext
                        
                    ElseIf rsControleAcesso("CO_INFO").Value > rsTipoInformacao(strCodigoA8).Value Then
                        'Tipo Informação existe mas não esta associado a um Grupo do Sistema MBS.
                        'Ação 02
                        
                        intCodOcorrencia = 2
                        strOcorrencia = strDescricaoTipoInfo & " existe mas não esta associado a um Grupo do Sistema MBS."
                        
                        strCodInformacao = CStr(rsTipoInformacao(strCodigoA8).Value)
                        strDesInformacao = CStr(rsTipoInformacao(strDescricaoA8).Value)
                        
                        rsTipoInformacao.MoveNext
                        
                    End If
                    
                    If intCodOcorrencia > 0 Then
                       blnExisteOcorrencia = True
                       If strCodInformacao <> strCodigoA8Ant Then
                          fgAppendNode xmlOcorrencias, "Repeat_Ocorrencias", "Grupo_Ocorrencias", ""
                          fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", strCodigoA8, strCodInformacao, "Repeat_Ocorrencias"
                          fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", strDescricaoA8, strDesInformacao, "Repeat_Ocorrencias"
                          fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", "CO_OCOR", CStr(intCodOcorrencia), "Repeat_Ocorrencias"
                          fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", "DE_OCOR", strOcorrencia, "Repeat_Ocorrencias"
                       End If
                    End If
                    
                  Else
                      'Tipo Informação não existe mas esta associado a um Grupo do Sistema MBS.
                      'Ação 01
                     
                      intCodOcorrencia = 1
                      strOcorrencia = strDescricaoTipoInfo & " não existe mas esta associado a um Grupo do Sistema MBS."
                    
                      strCodInformacao = rsControleAcesso("CO_INFO").Value
                      strDesInformacao = ""
                      
                      If intCodOcorrencia > 0 Then
                         blnExisteOcorrencia = True
                         If strCodigoA8Ant <> strCodInformacao Then
                            fgAppendNode xmlOcorrencias, "Repeat_Ocorrencias", "Grupo_Ocorrencias", ""
                            fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", strCodigoA8, strCodInformacao, "Repeat_Ocorrencias"
                            fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", strDescricaoA8, strDesInformacao, "Repeat_Ocorrencias"
                            fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", "CO_OCOR", CStr(intCodOcorrencia), "Repeat_Ocorrencias"
                            fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", "DE_OCOR", strOcorrencia, "Repeat_Ocorrencias"
                         End If
                      End If
                      
                      rsControleAcesso.MoveNext
                      
                  End If
                  
                  strCodigoA8Ant = strCodInformacao
                    
               Loop
           ElseIf rsControleAcesso.EOF And Not rsTipoInformacao.EOF Then
               'Tipo Informação existe mas não esta associado a um Grupo do Sistema MBS.
               'Ação 02
               
               If rsTipoInformacao.EOF Then
                  Exit Do
               End If
             
               intCodOcorrencia = 2
               strOcorrencia = strDescricaoTipoInfo & " existe mas não esta associado a um Grupo do Sistema MBS."
             
               strCodInformacao = CStr(rsTipoInformacao(strCodigoA8).Value)
               strDesInformacao = CStr(rsTipoInformacao(strDescricaoA8).Value)
           
               If intCodOcorrencia > 0 Then
                  blnExisteOcorrencia = True
                  If strCodigoA8Ant <> strCodInformacao Then
                     fgAppendNode xmlOcorrencias, "Repeat_Ocorrencias", "Grupo_Ocorrencias", ""
                     fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", strCodigoA8, strCodInformacao, "Repeat_Ocorrencias"
                     fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", strDescricaoA8, strDesInformacao, "Repeat_Ocorrencias"
                     fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", "CO_OCOR", CStr(intCodOcorrencia), "Repeat_Ocorrencias"
                     fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", "DE_OCOR", strOcorrencia, "Repeat_Ocorrencias"
                  End If
               End If
               
               rsTipoInformacao.MoveNext
               
           End If
           
           strCodigoA8Ant = strCodInformacao
           
        Loop
        
    Else
    
       Do Until rsControleAcesso.EOF
          'Tipo Informação não existe mas esta associado a um Grupo do Sistema MBS.
          'Ação 01
         
          intCodOcorrencia = 1
          strOcorrencia = strDescricaoTipoInfo & " não existe mas esta associado a um Grupo do Sistema MBS."
        
          strCodInformacao = rsControleAcesso("CO_INFO").Value
          strDesInformacao = ""
          
          If intCodOcorrencia > 0 Then
             blnExisteOcorrencia = True
             If strCodigoA8Ant <> strCodInformacao Then
                fgAppendNode xmlOcorrencias, "Repeat_Ocorrencias", "Grupo_Ocorrencias", ""
                fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", strCodigoA8, strCodInformacao, "Repeat_Ocorrencias"
                fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", strDescricaoA8, strDesInformacao, "Repeat_Ocorrencias"
                fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", "CO_OCOR", CStr(intCodOcorrencia), "Repeat_Ocorrencias"
                fgAppendNode xmlOcorrencias, "Grupo_Ocorrencias", "DE_OCOR", strOcorrencia, "Repeat_Ocorrencias"
             End If
          End If
          
          strCodigoA8Ant = strCodInformacao
          
          rsControleAcesso.MoveNext
          
       Loop
       
    End If
    
    If blnExisteOcorrencia Then
        VerificaControleAcessDado = xmlOcorrencias.xml
    Else
       VerificaControleAcessDado = ""
    End If
    
    Set xmlOcorrencias = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "VerificaControleAcessDado Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function
