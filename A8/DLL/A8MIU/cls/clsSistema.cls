VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa            : Regerbanc - Participações , Negócios e Serviços LTDA
'Componente         :
'Classe             : clsSistema
'Data Criação       : 16/09/2003
'Objetivo           :
'                   :
'Analista           : Andrea Bolcero
'
'Programador        : Cassiano
'Data               : 16/09/2003
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

Public Function LerTodos(Optional ByVal pintBancoLiqu As Integer = 0) As String

Dim objSistema                              As A6A7A8.clsSistema

    On Error GoTo ErrorHandler
    
    Set objSistema = CreateObject("A6A7A8.clsSistema")
    LerTodos = objSistema.LerTodos(pintBancoLiqu)
    Set objSistema = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    Set objSistema = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
