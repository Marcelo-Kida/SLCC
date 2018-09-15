VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIdPartCamaraVeicLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Implementado por Eduardo Quaresma em 03/01/05

'Objeto responsável pela manutenção do cadastro de veiculo legal

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
 
'Executar a operação : LerTodos

Public Function Executar(ByRef pxmlPropriedades As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlPropriedades.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pxmlPropriedades)
        Case "Propriedades"
            Executar = ObterPropriedades()
        Case "LerTodos"
            Executar = LerTodos()
        
        Case Else
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

Public Function LerTodos() As String
                         
Dim strSQL                                  As String
Dim strCondicao                             As String
    
On Error GoTo ErrorHandler

    strSQL = " SELECT    CO_VEIC_LEGA,                            " & vbNewLine & _
             "           SG_SIST,                                 " & vbNewLine & _
             "           CO_LOCA_LIQU,                            " & vbNewLine & _
             "           TP_IDEF_PART_CAMR,                       " & vbNewLine & _
             "           ID_PART_CAMR                             " & vbNewLine & _
             " FROM      A8.TB_IDEF_PART_CAMR_VEIC_LEGA             " & vbNewLine
    
    LerTodos = fgQueryXMLLerTodos("IdentificadorCamara", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar um registro da tabela A8.TB_VEIC_LEGA

Public Function Ler(ByVal pstrCodigoVeiculoLegal As String, _
                    ByVal pstrSiglaSistema As String, _
                    ByVal pintCodigoLocalLiquidante As Integer, _
                    ByVal pintTipoIdentificadorPartCamara As Integer) As String
                         
Dim strSQL                                   As String
    
    On Error GoTo ErrorHandler

    strSQL = " SELECT    CO_VEIC_LEGA,                            " & vbNewLine & _
             "           SG_SIST,                                 " & vbNewLine & _
             "           CO_LOCA_LIQU,                            " & vbNewLine & _
             "           TP_IDEF_PART_CAMR,                       " & vbNewLine & _
             "           ID_PART_CAMR                             " & vbNewLine & _
             " FROM      A8.TB_IDEF_PART_CAMR_VEIC_LEGA           " & vbNewLine & _
             " WHERE     CO_VEIC_LEGA = '" & pstrCodigoVeiculoLegal & "'" & vbNewLine & _
             "   AND     SG_SIST      = '" & pstrSiglaSistema & "'" & vbNewLine
          
    If pintCodigoLocalLiquidante <> 0 Then
       strSQL = strSQL & " AND CO_LOCA_LIQU = " & pintCodigoLocalLiquidante
    End If
    
    If pintTipoIdentificadorPartCamara <> 0 Then
       strSQL = strSQL & " AND TP_IDEF_PART_CAMR = " & pintTipoIdentificadorPartCamara
    End If
    
    Ler = fgQueryXMLLerTodos("IdentificadorCamara", strSQL, App.EXEName & "." & TypeName(Me))
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Incluir um registro na tabela de tipo de operação
Private Function flIncluir(ByRef pxmlVeiculoLegal As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A8.TB_IDEF_PART_CAMR_VEIC_LEGA (" & _
                "CO_VEIC_LEGA, " & _
                "SG_SIST, " & _
                "CO_LOCA_LIQU, " & _
                "TP_IDEF_PART_CAMR, " & _
                "ID_PART_CAMR  "

    With pxmlVeiculoLegal.documentElement
        strSQL = strSQL & " ) VALUES ( " & _
               "'" & .selectSingleNode("CO_VEIC_LEGA").Text & "', " & _
               "'" & .selectSingleNode("SG_SIST").Text & "', " & _
               .selectSingleNode("CO_LOCA_LIQU").Text & ", " & _
               .selectSingleNode("TP_IDEF_PART_CAMR").Text & ", " & _
               "'" & .selectSingleNode("ID_PART_CAMR").Text & "') "
               
    End With
    
    fgExecuteSQL strSQL
     
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Alterar um registro na tabela de tipo de operação
Private Function flAlterar(ByRef pxmlVeiculoLegal As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
On Error GoTo ErrorHandler

    With pxmlVeiculoLegal.documentElement
        strSQL = " UPDATE A8.TB_IDEF_PART_CAMR_VEIC_LEGA " & _
                "  SET      ID_PART_CAMR            ='" & .selectSingleNode("ID_PART_CAMR").Text & "' " & _
                " WHERE    CO_VEIC_LEGA  ='" & .selectSingleNode("CO_VEIC_LEGA").Text & "'" & _
                " AND      SG_SIST       ='" & .selectSingleNode("SG_SIST").Text & "'" & _
                " AND      CO_LOCA_LIQU  =" & .selectSingleNode("CO_LOCA_LIQU").Text & _
                " AND      TP_IDEF_PART_CAMR =" & .selectSingleNode("TP_IDEF_PART_CAMR").Text
    End With
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4002-Registro inexistente ou excluído por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Excluir um registro na tabela de tipo de operação

Private Function flExcluir(ByRef pxmlVeiculoLegal As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    With pxmlVeiculoLegal.documentElement
        strSQL = " DELETE FROM A8.TB_IDEF_PART_CAMR_VEIC_LEGA " & _
                 " WHERE    CO_VEIC_LEGA  ='" & .selectSingleNode("CO_VEIC_LEGA").Text & "'" & _
                 " AND      SG_SIST       ='" & .selectSingleNode("SG_SIST").Text & "'" & _
                 " AND      CO_LOCA_LIQU  =" & .selectSingleNode("CO_LOCA_LIQU").Text & _
                 " AND      TP_IDEF_PART_CAMR =" & .selectSingleNode("TP_IDEF_PART_CAMR").Text
    End With
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4002-Registro inexistente ou excluído por outro usuário
        lngCodigoErroNegocio = 4002
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter os nomes da colunas da tabela de tipo de operação

Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT * FROM A8.TB_IDEF_PART_CAMR_VEIC_LEGA "
    
    ObterPropriedades = fgPropriedades("IdentificadorCamara", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Gerenciar as operações de Inclusão,Alteração e Exclusão
Public Function Salvar(ByVal pxmlVeiculoLegal As MSXML2.DOMDocument40) As Boolean
On Error GoTo ErrorHandler
    
    'flRegrasNegocio pxmlVeiculoLegal
            
    If pxmlVeiculoLegal.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pxmlVeiculoLegal
    ElseIf pxmlVeiculoLegal.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        flAlterar pxmlVeiculoLegal
    ElseIf pxmlVeiculoLegal.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        flExcluir pxmlVeiculoLegal
    Else
        ' 7 - Operação Inválida
        lngCodigoErroNegocio = 7
        GoTo ErrorHandler
    End If
        
    Salvar = True
      
    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar um registro da tabela A8.TB_VEIC_LEGA atraves do
'Identificador Partipante Camara e Local de Liquidação

Public Function ObterVeiculoLegal(ByVal pvntIdentParticipanteCamara As Variant, _
                                  ByRef pstrCodigoVeiculoLegal As String, _
                                  ByRef pstrSiglaSistema As String, _
                         Optional ByVal plngCodLocalLiquidacao As Long = 0) As Boolean
                         
Dim strSQL                                   As String
Dim objRS                                    As ADODB.Recordset
    
On Error GoTo ErrorHandler

    strSQL = " SELECT    CO_VEIC_LEGA,                            " & vbNewLine & _
             "           SG_SIST,                                 " & vbNewLine & _
             "           CO_LOCA_LIQU,                            " & vbNewLine & _
             "           TP_IDEF_PART_CAMR,                       " & vbNewLine & _
             "           ID_PART_CAMR                             " & vbNewLine & _
             " FROM      A8.TB_IDEF_PART_CAMR_VEIC_LEGA           " & vbNewLine & _
             " WHERE     ID_PART_CAMR = '" & pvntIdentParticipanteCamara & "'" & vbNewLine
          
    If plngCodLocalLiquidacao <> 0 Then
       strSQL = strSQL & " AND CO_LOCA_LIQU = " & plngCodLocalLiquidacao
    End If
    
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.EOF Then
        ObterVeiculoLegal = False
        pstrCodigoVeiculoLegal = vbNullString
        pstrSiglaSistema = vbNullString
    Else
        ObterVeiculoLegal = True
        pstrCodigoVeiculoLegal = objRS!CO_VEIC_LEGA
        pstrSiglaSistema = objRS!SG_SIST
    End If
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


