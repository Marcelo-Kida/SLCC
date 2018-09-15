VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParametrosGerais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa        : Regerbanc - Participações , Negócios e Serviços LTDA
'Pacote         :
'Classe         : clsParametrosGerais
'Data Criação   : 01/06/2004
'Objetivo       :
'
'Analista       : Michel P. Barros / Marcelo Kida
'
'Programador    : Cassiano Nicolosi
'Data           : 01/06/2004
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :

Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Implements COMSVCSLib.ObjectControl
Private objContext                          As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

Public Function Executar(ByRef xmlParametrizacao As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String
Dim strCampoSequencia                       As String
Dim strCampoVarchar                         As String

    On Error GoTo ErrorHandler

    strOperacao = xmlParametrizacao.selectSingleNode("//*/@Operacao").Text
    strCampoSequencia = xmlParametrizacao.selectSingleNode("//*/CO_TEXT_XML").Text
    
    'If Not xmlParametrizacao.selectSingleNode("//*/TX_XML") Is Nothing Then
     strCampoVarchar = xmlParametrizacao.xml
    'End If

    Select Case strOperacao
        Case "Alterar"
            Executar = flAlterar(strCampoVarchar)
            
        Case "Ler"
            Executar = flLer(strCampoSequencia)
            
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
            
    End Select

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flAlterar(ByVal pstrConteudoCampoVarchar As String) As Long

Dim strSQL                                  As String
    
    On Error GoTo ErrorHandler

    strSQL = "DELETE FROM A8.TB_TEXT_XML WHERE CO_TEXT_XML = 0"
    
    Call fgExecuteSQL(strSQL)
    
    flAlterar = fgInsertVarchar4000(pstrConteudoCampoVarchar, "0", False)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function flLer(ByVal pstrConteudoCampoSequencia As String) As String

Dim xmlParamGeral                           As MSXML2.DOMDocument40
Dim strParamGeral                           As String
    
On Error GoTo ErrorHandler

    strParamGeral = fgSelectVarchar4000(pstrConteudoCampoSequencia, False)
    
    If strParamGeral = vbNullString Then
        Set xmlParamGeral = CreateObject("MSXML2.DOMDocument.4.0")
        
        xmlParamGeral.preserveWhiteSpace = True
        xmlParamGeral.Load App.Path & "/ParametrosGerais.XML"
        strParamGeral = xmlParamGeral.xml
        
        Set xmlParamGeral = Nothing
    End If
    
    flLer = strParamGeral
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set xmlParamGeral = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLer", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
