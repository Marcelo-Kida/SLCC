VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente         : BJMIU\r\nClasse             : clsAjusteAgendamento\r\nData Criação       : 24/07/2003\r\nObjetivo           : Iniciar a classes não transasionais para a tela de Ajuste e Agendamento\r\n                   :\r\nAnalista           : Carlos Fortes\r\n\r\nProgramador        : \r\nData               : \r\n\r\nData Teste         :\r\nAutor              :\r\n\r\nData Alteração     :\r\nAutor              :\r\nObjetivo           :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F1DC0F80091"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito ao Ajuste de Movimentação.

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

' Este método aciona a leitura da data da última posição de um determinado veículo legal, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterDataCaixaVeiculoLegal(ByVal pstrCodigoVeiculoLegal As String, _
                                           ByVal pstrSiglaSistema As String, _
                                           ByRef pvntCodErro As Variant, _
                                           ByRef pvntMensagemErro As Variant) As String
                                           
Dim objCaixaSubReserva                      As A6SubReserva.clsCaixaSubReserva
Dim xmlPosicaoCaixa                         As MSXML2.DOMDocument40
Dim datObterDataCaixaVeiculoLegal           As Date
                                           
On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
                                           
    Set xmlPosicaoCaixa = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objCaixaSubReserva = CreateObject("A6SubReserva.clsCaixaSubReserva")
    Call xmlPosicaoCaixa.loadXML(objCaixaSubReserva.ObterPosicaoCaixaSubReserva(pstrCodigoVeiculoLegal, pstrSiglaSistema))
    Set objCaixaSubReserva = Nothing
    
    datObterDataCaixaVeiculoLegal = fgDtXML_To_Date(xmlPosicaoCaixa.documentElement.selectSingleNode("Grupo_PosicaoCaixaSubReserva/DT_CAIX_SUB_RESE").Text)
    
    ObterDataCaixaVeiculoLegal = CStr(Format(datObterDataCaixaVeiculoLegal, "DD/MM/YYYY"))
    
    Set xmlPosicaoCaixa = Nothing
                                           
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objCaixaSubReserva = Nothing
    Set xmlPosicaoCaixa = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDataCaixaVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
                                           
End Function

' Este método aciona a leitura de intens de caixa associados a grupos de veículos legais, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterVeiculoLegalItemCaixa(ByVal plngGrupoVeiculoLegal As Long, _
                                           ByRef pvntCodErro As Variant, _
                                           ByRef pvntMensagemErro As Variant) As String
    
Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal
Dim objItemCaixa                            As A6SubReserva.clsItemCaixaGrupoVeicLegal
Dim strXMLDoc                               As String
    
On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objVeiculoLegal = CreateObject("A6a7a8.clsVeiculoLegal")
    Set objItemCaixa = CreateObject("A6SubReserva.clsItemCaixaGrupoVeicLegal")
    
    strXMLDoc = "<frmAjuste><Grupo_Dados>"
    strXMLDoc = strXMLDoc & objVeiculoLegal.LerTodos(, plngGrupoVeiculoLegal)
    strXMLDoc = strXMLDoc & objItemCaixa.ObterItensCaixaPorGrupoVeicLegal(plngGrupoVeiculoLegal)
    strXMLDoc = strXMLDoc & "</Grupo_Dados></frmAjuste>"
    
    Set objVeiculoLegal = Nothing
    Set objItemCaixa = Nothing

    ObterVeiculoLegalItemCaixa = strXMLDoc

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objVeiculoLegal = Nothing
    Set objItemCaixa = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterVeiculoLegalItemCaixa Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function
