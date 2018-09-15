VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA6A7A8Funcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa            : Regerbanc - Participações , Negócios e Serviços LTDA
'Componente         : A8MIU
'Classe             : clsA6A7A8Funcoes
'Data Criação       : 16/12/2003
'Objetivo           :
'                   :
'Analista           : Carlos Fortes
'
'Programador        : Eder Andrade
'Data               : 16/12/2003
'
'Data Teste         :
'Autor              :
'
'Data Alteração     :
'Autor              :
'Objetivo           :

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

Public Function AdicionarDiasUteis(ByVal pdatData As String, _
                                   ByVal pintQtdeDias As Integer, _
                                   ByVal plngMovimento As Long, _
                                   ByRef pvntCodErro As Variant, _
                                   ByRef pvntMensagemErro As Variant) As String

Dim objA6A7A8Funcoes                        As A6A7A8.clsA6A7A8Funcoes
Dim datData                                 As Date
Dim datDataAux                              As Date

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    If Trim(pdatData) <> "" Then
       datData = CDate(Format(pdatData, "DD/MM/YYYY"))
    End If
    
    Set objA6A7A8Funcoes = CreateObject("A6A7a8.clsA6A7A8Funcoes")
    datDataAux = objA6A7A8Funcoes.AdicionarDiasUteis(datData, _
                                                     pintQtdeDias, _
                                                     plngMovimento)
                                                  
    AdicionarDiasUteis = CStr(Format(datDataAux, "DD/MM/YYYY"))
                                                             
    Set objA6A7A8Funcoes = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set objA6A7A8Funcoes = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "AdicionarDiasUteis Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

Public Function ObterDominioSPB(ByVal pstrNomeTipoTag As String, _
                       Optional ByVal pstrTagRetorno As String = "") As String
On Error GoTo ErrorHandler
Dim objDominioSPB                           As A8LQS.clsMensagemSLCC

    Set objDominioSPB = CreateObject("A8LQS.clsMensagemSLCC")
    ObterDominioSPB = objDominioSPB.ObterDominioSPB(pstrNomeTipoTag, pstrTagRetorno)
    Set objDominioSPB = Nothing
    
    Exit Function
ErrorHandler:
    Set objDominioSPB = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDominioSPB Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
