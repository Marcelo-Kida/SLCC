VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParamCtrlPostagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF773940141"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Objeto responsável pelo fornecimento de informações de parâmetros de postagem, bem como a manipulação dos dados cadastrados.
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

'Consultar todos os atributos de um parâmetro de postagem específico.
Public Function Ler(ByVal pdtmDataHoraParametro As String) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler
    
    strSQL = " SELECT  DH_PARM,          " & _
             "         QT_TENT_ENVI_MESG," & _
             "         QT_SEGU_FREQ_VERI," & _
             "         QT_TEMP_ENTG_MESG," & _
             "         QT_TEMP_RETI_MESG," & _
             "         CO_USUA_ULTI_ATLZ," & _
             "         CO_ETCA_ULTI_ATLZ," & _
             "         DH_ULTI_ATLZ      " & _
             "   FROM  A7.TB_PARM_CTRL_PSTA " & _
             "  WHERE  DH_PARM = " & fgDtXML_To_Oracle(Format(fgDtHrStr_To_DateTime(pdtmDataHoraParametro), "YYYYMMDD")) & ")"
    
    Ler = fgQueryXMLLer("ParametroControlePostagem", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Consultar todos os parâmetros de controle de postagem cadastrados
Public Function LerTodos(Optional ByVal pdtmDataHoraParametro As String) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler
 
    strSQL = " SELECT  QT_TENT_ENVI_MESG," & _
             "         DH_PARM,          " & _
             "         QT_FREQ_VERI,     " & _
             "         QT_TEMP_ENTG_MESG," & _
             "         QT_TEMP_RETI_MESG," & _
             "         CO_USUA_ULTI_ATLZ," & _
             "         CO_ETCA_TRAB_ULTI_ATLZ," & _
             "         DH_ULTI_ATLZ      " & _
             "   FROM  A7.TB_PARM_CTRL_PSTA "
    
    If Trim(pdtmDataHoraParametro) <> vbNullString Then
        strSQL = strSQL & " WHERE     DH_PARM = " & fgDtXML_To_Oracle(Format(fgDtHrStr_To_DateTime(pdtmDataHoraParametro), "YYYYMMDD")) & ") "
    End If
    
    strSQL = strSQL & " ORDER BY DH_PARM DESC"

    LerTodos = fgQueryXMLLerTodos("ParametroControlePostagem", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Controlar o roteamento e execução dos métodos privados, tais como:
' -  Incluir
' -  Alterar
' -  Excluir
' -  LerTodos
' -  Ler
Public Function Executar(ByRef pobjPropriedades As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pobjPropriedades.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar"
            Executar = Salvar(pobjPropriedades)
        Case "Ler"
            Executar = Ler(pobjPropriedades.selectSingleNode("//DH_PARM").Text)
        Case "LerTodos"
            Executar = LerTodos()
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
    End Select

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a estrutura da tabela TB_PARM_CTRL_PSTA.
Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT * " & _
             "   FROM A7.TB_PARM_CTRL_PSTA "

    ObterPropriedades = fgPropriedades("Grupo_ParametroControlePostagem", strSQL, "A7Server.clsParamCtrlPostagem")

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Promove a inclusão de um novo parâmetro de controle de postagem.
Private Function flIncluir(ByRef objDOMDocument As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler
            
    strSQL = " INSERT INTO A7.TB_PARM_CTRL_PSTA ( " & _
            " DH_PARM, " & _
            " QT_TEMP_RETI_MESG, " & _
            " QT_TENT_ENVI_MESG, " & _
            " QT_TEMP_ENTG_MESG, " & _
            " QT_FREQ_VERI, " & _
            " CO_USUA_ULTI_ATLZ, " & _
            " CO_ETCA_TRAB_ULTI_ATLZ, " & _
            " DH_ULTI_ATLZ " & _
            " ) VALUES ( " & _
            " SYSDATE , " & _
            CLng(objDOMDocument.selectSingleNode("//Grupo_ParametroControlePostagem/QT_TEMP_RETI_MESG").Text) & ", " & _
            CLng(objDOMDocument.selectSingleNode("//Grupo_ParametroControlePostagem/QT_TENT_ENVI_MESG").Text) & ", " & _
            CLng(objDOMDocument.selectSingleNode("//Grupo_ParametroControlePostagem/QT_TEMP_ENTG_MESG").Text) & ", " & _
            CLng(objDOMDocument.selectSingleNode("//Grupo_ParametroControlePostagem/QT_FREQ_VERI").Text) & ", " & _
            "'" & fgObterUsuarioRede & "', " & _
            "'" & fgObterEstacaoTrabalhoUsuario & "', " & _
            "SYSDATE)"
            
    fgExecuteSQL strSQL
      
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function

'Rotear o comando salvar para os método privado flIncluir.
Public Function Salvar(ByRef pobjPropriedades As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
            
    If pobjPropriedades.selectSingleNode("//Grupo_ParametroControlePostagem/@Operacao").Text = "Incluir" Then
        flIncluir pobjPropriedades
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


