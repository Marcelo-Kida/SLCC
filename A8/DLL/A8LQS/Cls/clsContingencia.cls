VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
 
Attribute VB_Name = "clsContingencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : \r\nData Criação   : \r\nObjetivo       : \r\n\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F3A4E620157"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'' Objeto responsável pelo fornecimento de informações do estado de contingência
'' dos sistemas legados

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private objContext                       As COMSVCSLib.ObjectContext
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

'Verificar o sistema está em contingencia

Public Function VerificarSistContingencia(ByVal pstrSiglaSistema As String) As Boolean

Dim strSQL                                  As String
Dim rsContingencia                          As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsContingencia = CreateObject("ADODB.Recordset")
 
    strSQL = "Select    IN_SIST_SITU_CNTG " & vbCrLf & _
             "  From    A8.TB_SITU_CNTG_SIST " & vbCrLf & _
             " Where    SG_SIST =   '" & pstrSiglaSistema & "'"

    Set rsContingencia = fgQuerySQL(strSQL)

    If Not rsContingencia.EOF Then
        If rsContingencia!IN_SIST_SITU_CNTG = enumIndicadorSimNao.sim Then
            VerificarSistContingencia = True
            Exit Function
        End If
    End If

    VerificarSistContingencia = False

    Set rsContingencia = Nothing

    Exit Function
ErrorHandler:

    Set rsContingencia = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "VerificarSistContingencia Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
