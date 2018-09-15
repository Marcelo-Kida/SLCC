VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFatoGeradorAlertaTipoBackOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa        : Regerbanc - Participações , Negócios e Serviços LTDA
'Pacote         :
'Classe         : clsFatoGeradorAlertaTipoBackOffice
'Data Criação   : 22/09/2003
'Objetivo       :
'
'Analista       : Andrea Bolcero
'
'Programador    : Cassiano Nicolosi
'Data           : 22/09/2003
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

Public Function LerTodos(ByVal pintTipoBackOffice As enumTipoBackOffice) As String

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String
Dim xmlSegregacaoDados                      As MSXML2.DOMDocument40

    'Este trecho será substituido posteriormente por uma funcao Obter..."
    Set xmlSegregacaoDados = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlSegregacaoDados.Load("C:\SegregaAcessoDados.xml")
    'Final trecho
    
    On Error GoTo ErrorHandler

    strSQL = " SELECT    B.DE_FATO_GERA_ALER,                           " & vbNewLine & _
             "           A.IN_ENVI_EMAIL,                               " & vbNewLine & _
             "           A.DE_EMAIL_ENVI_ALER,                          " & vbNewLine & _
             "           A.IN_EMIS_ALER_TELA                            " & vbNewLine & _
             " FROM      A8.TB_FATO_GERA_ALER_BKOF A,                   " & vbNewLine & _
             "           A8.TB_FATO_GERA_ALER B                         " & vbNewLine & _
             " WHERE     A.CO_FATO_GERA_ALER = B.CO_FATO_GERA_ALER      " & vbNewLine & _
             " AND       A.TP_BKOF           = " & pintTipoBackOffice & vbNewLine

    strWhereSegregacao = fgSegregaDados("A8.TB_FATO_GERA_ALER_BKOF", _
                                        xmlSegregacaoDados, _
                                        False, _
                                        "A", _
                                        "A", _
                                        False, _
                                        False, _
                                        True, _
                                        False, _
                                        False)
    
    strSQL = strSQL & strWhereSegregacao & vbNewLine
    strSQL = strSQL & " ORDER BY  B.DE_FATO_GERA_ALER"

    LerTodos = fgQueryXMLLerTodos("FatoGeradorAlertaTipoBackOffice", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function ObterPropriedades() As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT * FROM A8.TB_FATO_GERA_ALER_BKOF"
    ObterPropriedades = fgPropriedades("Grupo_FatoGeradorAlertaTipoBackOffice", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
