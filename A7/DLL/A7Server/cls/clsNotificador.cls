VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNotificador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF768DD02B4"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA
'Pacote         :
'Classe         : clsNotificar
'Data Criação   : 01/07/2003
'Objetivo       :
'
'Analista       : Marcelo Kida
'
'Programador    : Marcelo Kida
'Data           : 01/07/2003
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :


Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

Private ObjectContext                       As COMSVCSLib.ObjectContext
Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
End Sub

Public Function Notificar(ByRef pobjDomMensagem As MSXML2.DOMDocument40) As Boolean


On Error GoTo ErrorHandler
  
    
    



    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Notificar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function ObterParametroNotificacao(ByRef pobjDomMensagem As MSXML2.DOMDocument40) As String


On Error GoTo ErrorHandler


    Exit Function
ErrorHandler:
    


    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterParametroNotificacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function PostarEmail(ByRef pobjDomMensagem As MSXML2.DOMDocument40) As Boolean
    On Error GoTo ErrorHandler


    '## your code goes here...

    Exit Function
ErrorHandler:


    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "PostarEmail Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

