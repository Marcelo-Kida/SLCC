VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOcorrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa        : Regerbanc - Participações , Negócios e Serviços LTDA
'Pacote         :
'Classe         : clsOcorrencia
'Data Criação   : 08/08/2003
'Objetivo       :
'
'Analista       : Maurício
'
'Programador    : Cassiano Nicolosi
'Data           : 08/08/2003
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

Public Function Executar(ByRef pobjOcorrencia As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = pobjOcorrencia.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar"
            Executar = flAtualizar(pobjOcorrencia)
        
        Case "Ler"
            Executar = Ler(pobjOcorrencia.selectSingleNode("//SG_SIST").Text, _
                  CLng(Val(pobjOcorrencia.selectSingleNode("//CO_BANC_LIQU").Text)))
        
        Case "LerTodos"
            Executar = LerTodos()
        
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
    
    End Select

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flAtualizar(ByRef pobjOcorrencia As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String
Dim larrOcorrencias()                       As String
Dim liCount                                 As Integer

    On Error GoTo ErrorHandler

    With pobjOcorrencia
        strSQL = "DELETE FROM A7.TB_RESPOSTA_OCORRENCIA " & _
                " WHERE  SG_SIST = '" & .selectSingleNode("//Grupo_Ocorrencia/SG_SIST").Text & "'" & _
                " AND    CO_BANC_LIQU = " & .selectSingleNode("//Grupo_Ocorrencia/CO_BANC_LIQU").Text
        
        Call fgExecuteSQL(strSQL)
          
        If .selectSingleNode("//Grupo_Ocorrencia/CO_OCOR").Text <> vbNullString Then
            larrOcorrencias = Split(.selectSingleNode("//Grupo_Ocorrencia/CO_OCOR").Text, ";")
            
            For liCount = LBound(larrOcorrencias) To UBound(larrOcorrencias)
                strSQL = " INSERT INTO A7.TB_RESPOSTA_OCORRENCIA ( " & vbCrLf & _
                        "             CO_OCOR,                    " & vbCrLf & _
                        "             SG_SIST,                    " & vbCrLf & _
                        "             CO_BANC_LIQU,               " & vbCrLf & _
                        "             IN_ENVI_RSPT,               " & vbCrLf & _
                        "             CO_USUA_ULTI_ATLZ,          " & vbCrLf & _
                        "             DH_ULTI_ATLZ                " & vbCrLf & _
                        " ) VALUES ( " & vbCrLf & _
                        "" & larrOcorrencias(liCount) & ",                                          " & vbCrLf & _
                        "'" & .selectSingleNode("//Grupo_Ocorrencia/SG_SIST").Text & "',            " & vbCrLf & _
                        "" & .selectSingleNode("//Grupo_Ocorrencia/CO_BANC_LIQU").Text & ",         " & vbCrLf & _
                        "'" & .selectSingleNode("//Grupo_Ocorrencia/IN_ENVI_RSPT").Text & "',       " & vbCrLf & _
                        "'" & .selectSingleNode("//Grupo_Ocorrencia/CO_USUA_ULTI_ATLZ").Text & "',  " & vbCrLf & _
                        "SYSDATE)"
        
                Call fgExecuteSQL(strSQL)
            Next
        End If
    End With
      
    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Atualizar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function Ler(ByVal psSiglaSistema As String, _
                    ByVal plCodigoBancoLiquidante As Long) As String

Dim strSQL                                   As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT    CO_OCOR                           " & _
            " FROM      A7.TB_RESPOSTA_OCORRENCIA         " & _
            " WHERE     SG_SIST = '" & psSiglaSistema & "'" & _
            " AND       CO_BANC_LIQU = " & plCodigoBancoLiquidante
    
    Ler = fgQueryXMLLerTodos("Ocorrencia", strSQL)

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function LerTodos() As String
    
Dim strSQL                                   As String
    
    On Error GoTo ErrorHandler

    strSQL = " Select    CO_OCOR,       " & vbCrLf & _
            "           DE_OCOR        " & vbCrLf & _
            " From      TB_OCORRENCIA  " & vbCrLf & _
            " Order by  CO_OCOR ASC    "
            
    LerTodos = fgQueryXMLLerTodos("Ocorrencia", strSQL)

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT    *                         " & vbCrLf & _
            " FROM      A7.TB_RESPOSTA_OCORRENCIA "

    ObterPropriedades = fgPropriedades("Grupo_Ocorrencia", strSQL, "A7BusServer.clsOcorrencia")

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
