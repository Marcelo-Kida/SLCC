VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTCaixaSubReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente         : BJMIU\r\nClasse             : clsTAberturaCaixa\r\nData Criação       : 24/07/2003\r\nObjetivo           : Iniciar a classes transasionais para a tela de Abertura do Caixa\r\n                   :\r\nAnalista           : Carlos Fortes\r\n\r\nProgramador        : \r\nData               : \r\n\r\nData Teste         :\r\nAutor              :\r\n\r\nData Alteração     :\r\nAutor              :\r\nObjetivo           :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F1C73230291"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

' Este componente tem como objetivo, agrupar os métodos responsáveis pelo controle transacional nas operações
' de atualização de tabelas, referentes à Abertura e Fechamento de Caixa.
' É responsável também, pela ligação da camada intermediária, à camada de negócios.

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

' Este método aciona a abertura de caixa, na camada de negócios.

Public Function AbrirCaixa(ByVal pxmlDomNode As MSXML2.IXMLDOMNode) As String

Dim objCaixaReserva                         As A6SubReserva.clsCaixaSubReserva

    On Error GoTo ErrorHandler

    Set objCaixaReserva = CreateObject("A6SubReserva.clsCaixaSubReserva")

    Call objCaixaReserva.AbrirCaixa(pxmlDomNode)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    AbrirCaixa = vbNullString
    
    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        objContext.DisableCommit
    End If

    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    AbrirCaixa = Err.Description

End Function

' Este método aciona o fechamento de caixa, na camada de negócios.

Public Function FecharCaixa(ByVal pstrCodigoVeiculoLegal As String, _
                            ByVal pstrSiglaSistema As String) As String
    
Dim objCaixaSubReserva                      As A6SubReserva.clsCaixaSubReserva
    
    On Error GoTo ErrorHandler

    Set objCaixaSubReserva = CreateObject("A6SubReserva.clsCaixaSubReserva")

    objCaixaSubReserva.FecharCaixa pstrCodigoVeiculoLegal, pstrSiglaSistema
    
    Set objCaixaSubReserva = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    FecharCaixa = vbNullString
    
    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        objContext.DisableCommit
    End If

    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    FecharCaixa = Err.Description

End Function
