VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Empresa            : Regerbanc - Partticipações , Negócios e Serviços LTDA\r\nComponente         : MIU\r\nClasse             : clsTransacao\r\nData Criação       : 07/10/2002\r\nObjetivo           : Iniciar a Transação\r\n                   :\r\nAnalista           : Adilson G. Damasceno/Marcelo Kida\r\n\r\nProgramador        : Marcelo Kida\r\nData               : 07/10/2002\r\n\r\nData Teste         :\r\nAutor              :\r\n\r\nData Alteração     :\r\nAutor              :\r\nObjetivo           :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F0F215A0215"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e outras classes transacionais genéricas do objeto A8MIU.dll.

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

' Este método aciona genericamente a execução de um outro método na camada de negócios em um contexto transacional,
' e retorna para a camada de interface.
' O Objeto e o Método a serem executados são passados via XML.
Public Function Executar(ByVal pxmlPropriedades As MSXML2.DOMDocument40) As String
    
Dim objExecutar                             As Object
Dim strNomeObjeto                           As String

On Error GoTo ErrorHandler

    strNomeObjeto = Trim(pxmlPropriedades.selectSingleNode("//*/@Objeto").Text)

    Set objExecutar = CreateObject(strNomeObjeto)
    Executar = objExecutar.Executar(pxmlPropriedades)
    Set objExecutar = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    Set objExecutar = Nothing

    If Not objContext Is Nothing Then
        objContext.DisableCommit
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona genericamente a execução de um outro método na camada de negócios em um contexto transacional,
' e retorna para a camada de interface.
' O Objeto e o Método a serem executados são passados via XML.
Public Function ProcessarItemLote(ByVal pxmlPropriedades As MSXML2.DOMDocument40) As String
    
Dim objExecutar                             As Object
Dim strNomeObjeto                           As String

    On Error GoTo ErrorHandler

    strNomeObjeto = Trim(pxmlPropriedades.selectSingleNode("//*/@Objeto").Text)

    Set objExecutar = CreateObject(strNomeObjeto)
    ProcessarItemLote = objExecutar.Executar(pxmlPropriedades)
    Set objExecutar = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objExecutar = Nothing

    If Not objContext Is Nothing Then
        objContext.DisableCommit
    End If

    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    ProcessarItemLote = Err.Description
    
End Function
