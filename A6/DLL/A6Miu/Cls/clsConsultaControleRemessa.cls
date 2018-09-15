VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsultaControleRemessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente         : BJMIU\r\nClasse             : clsConsultaControleRemessa\r\nData Criação       : 24/07/2003\r\nObjetivo           : Iniciar a classes não transasionais para a tela de Consulta Controle da Remessa\r\n                   :\r\nAnalista           : Carlos Fortes\r\n\r\nProgramador        : \r\nData               : \r\n\r\nData Teste         :\r\nAutor              :\r\n\r\nData Alteração     :\r\nAutor              :\r\nObjetivo           :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F18559F0140"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito às Consultas de Controle de Remessas.

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

' Este método aciona a leitura ao cadastro de controle de remessas enviadas ao A6, na camada de negócios,
' e retorna para a camada de interface.

Public Function LerTodos(pxmlDocFiltros As String, _
                         pvntCodErro As Variant, _
                         pvntMensagemErro As Variant) As String

Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim objControleRemessa                      As A6SubReserva.clsControleRemessa

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Set objControleRemessa = CreateObject("A6SubReserva.clsControleRemessa")

    If Not xmlFiltros.loadXML(pxmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If

    LerTodos = objControleRemessa.LerTodos(xmlFiltros)

    Set xmlFiltros = Nothing
    Set objControleRemessa = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function
