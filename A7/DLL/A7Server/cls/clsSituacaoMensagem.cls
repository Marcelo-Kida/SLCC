VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSituacaoMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF758FB039E"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Objeto responsável pela histórico das situações das mensagens transitadas pelo sistema A7.
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

'Inclui nova Situação para a mensagem transitada no sistema A7.
Public Function SalvarSituacaoMensagem(ByVal pxmlMensagem As MSXML2.DOMDocument40, _
                                       ByVal plngCodigoOcorrencia As Long) As Long

Dim strSQL                                   As String
Dim lngProxNumeroOrdem                       As Long
Dim objRS                                    As ADODB.Recordset
Dim objRespostaOcorrencia                    As A7Server.clsRespostaOcorrencia

On Error GoTo ErrorHandler
    
    strSQL = " SELECT  MAX(NU_SEQU_SITU_MESG) " & _
             "   FROM  A7.TB_SITU_MESG " & _
            "   WHERE  CO_MESG = " & CLng(pxmlMensagem.documentElement.selectSingleNode("CO_MESG").Text)

    Set objRS = fgQuerySQL(strSQL)
    
    If IsNull(objRS.fields(0)) Then
        lngProxNumeroOrdem = 1
    Else
        lngProxNumeroOrdem = objRS.fields(0) + 1
    End If
        
    objRS.Close
        
    strSQL = " INSERT INTO  A7.TB_SITU_MESG (" & _
            "   NU_SEQU_SITU_MESG , " & _
            "   CO_MESG, " & _
            "   CO_OCOR_MESG, " & _
            "   DH_OCOR_MESG, " & _
            "   TX_DTLH_OCOR_ERRO " & _
            "   ) VALUES ( " & _
            lngProxNumeroOrdem & ", " & _
            CLng(pxmlMensagem.documentElement.selectSingleNode("CO_MESG").Text) & ", " & _
            plngCodigoOcorrencia & "," & _
            "SYSDATE" & ","
                    
    If Trim(pxmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").Text) = "" Then
        strSQL = strSQL & "NULL)"
    Else
        strSQL = strSQL & "'" & Replace(Trim(pxmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").Text), "'", " ") & "')"
    End If

    SalvarSituacaoMensagem = fgExecuteSQL(strSQL)
    
    Set objRespostaOcorrencia = CreateObject("A7Server.clsRespostaOcorrencia")
    objRespostaOcorrencia.EnviarRespostaOcorrencia pxmlMensagem, plngCodigoOcorrencia
    Set objRespostaOcorrencia = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objRespostaOcorrencia = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "SalvarSituacaoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
