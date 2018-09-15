VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
 
Attribute VB_Name = "clsLiquidacaoTQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' Oferece os métodos para permitir a transição da liquidação das operações CETIP
'' do TQ para o SLCC

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



'' Busca operações que foram liquidadas pelo TQ, para providenciar o envio de
'' previsões, conta corrente e alertas
Public Function LiquidarOperacoesTQ() As String



    'Verifica quais operações estão pendentes para liquidação

On Error GoTo ErrorHandler
    Dim strSQL                                  As String
    Dim Rs                                      As Recordset
    Dim objRemessa                              As clsTRemessa

    strSQL = " SELECT  NU_SEQU_OPER_ATIV                                     " & vbCrLf & _
             " FROM    A8.TB_OPER_ATIV                                       " & vbCrLf & _
             " WHERE   CO_ULTI_SITU_PROC = " & enumStatusOperacao.LiberadaTQ & vbCrLf & _
             " AND     DT_OPER_ATIV BETWEEN " & fgDtOracle_To_OracleBetween( _
                                                fgFormataDataHoraOracle( _
                                                fgDataHoraServidor(DataAux) _
                                                ) _
                                                ) & vbCrLf

    Set Rs = fgQuerySQL(strSQL)
    
    Do While Not Rs.EOF
        'Libera cada uma
        Set objRemessa = CreateObject("A8LQS.clsTRemessa")
        objRemessa.LiquidarOperacaoTQ Rs("NU_SEQU_OPER_ATIV")
        Set objRemessa = Nothing
        
        Rs.MoveNext
    Loop
    
    Set Rs = Nothing
    
    LiquidarOperacoesTQ = vbNullString
    
    Exit Function
ErrorHandler:
    
    LiquidarOperacoesTQ = Err.Description
    Set objRemessa = Nothing

    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarOperacoesTQ Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function




 