VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLiquidacaoFutura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Oferece os métodos para permitir o gerenciamento do fluxo normal de operações registradas no sistema
'' com data de liquidação futura.

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

'' Busca operações com status de liquidação futura, para providenciar o envio de
'' previsões, conta corrente e alertas
Public Function GerenciarFluxoOperacoesLiquidacaoFutura() As String

Dim strSQL                                  As String
Dim Rs                                      As Recordset
Dim objRemessa                              As clsTRemessa
    
    On Error GoTo ErrorHandler

    strSQL = " SELECT  NU_SEQU_OPER_ATIV                                          " & vbCrLf & _
             " FROM    A8.TB_OPER_ATIV                                            " & vbCrLf & _
             " WHERE   CO_ULTI_SITU_PROC =  " & enumStatusOperacao.LiquidacaoFutura & vbCrLf & _
             " AND     DT_OPER_ATIV BETWEEN " & fgDtOracle_To_OracleBetween( _
                                                fgFormataDataHoraOracle(fgDataHoraServidor(DataAux)))

    Set Rs = fgQuerySQL(strSQL)
    
    Do While Not Rs.EOF
        Set objRemessa = CreateObject("A8LQS.clsTRemessa")
        objRemessa.GerenciarOperacaoLiquidacaoFutura Rs("NU_SEQU_OPER_ATIV")
        Set objRemessa = Nothing
        
        Rs.MoveNext
    Loop
    
    Set Rs = Nothing
    
    GerenciarFluxoOperacoesLiquidacaoFutura = vbNullString
    
    Exit Function

ErrorHandler:
    GerenciarFluxoOperacoesLiquidacaoFutura = Err.Description
    Set objRemessa = Nothing

    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarFluxoOperacoesLiquidacaoFutura Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
