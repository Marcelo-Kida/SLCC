VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGerenciadorReprocCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'   Objeto responsável pelos métodos de controle do parâmetro de reprocessamento de Conta Corrente.

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

'Obter todas as parametrizações de histórico de conta corrente, utilizando a segregação de acesso
Private Function ObterParamReprocCC() As ADODB.Recordset

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String

On Error GoTo ErrorHandler

    strSQL = " SELECT CD_EMPR           ," & vbNewLine & _
             " TP_BKOF                  ," & vbNewLine & _
             " TP_OPER                  ," & vbNewLine & _
             " IN_LANC_DEBT_CRED        ," & vbNewLine & _
             " TP_CNAL_VEND             ," & vbNewLine & _
             " QT_REPR_CNTA_CRRT        ," & vbNewLine & _
             " QT_HORA_INTL_REPR        ," & vbNewLine & _
             " IN_ATIV                  ," & vbNewLine & _
             " CD_USUA_ULTI_ATLZ        ," & vbNewLine & _
             " CD_ETCA_USUA_ULTI_ATLZ   ," & vbNewLine & _
             " DH_ULTI_ATLZ             ," & vbNewLine & _
             " HR_LIMI_ITGR AS HO_LIMI_ITGR " & vbNewLine & _
             " FROM   A8.TB_PARM_REPR_CNTA_CRRT   " & vbNewLine & _
             " WHERE  IN_ATIV = " & enumIndicadorSimNao.Sim

    Set ObterParamReprocCC = fgQuerySQL(strSQL)

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterParamReprocCC Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function Processar() As String

Dim objRSParamReprocCC                      As ADODB.Recordset
Dim objRSLanc                               As ADODB.Recordset
Dim objGerenciadorReprocCCT                 As A8LQS.clsGerenciadorReprocCCT
Dim strErro                                 As String
Dim xmlLancamentos                          As String
Dim datDataOper                             As Date
Dim datHoraUltiReproc                       As Date
Dim datHoraAtual                            As Date
Dim datHoraGradeHorario                     As Date

On Error GoTo ErrorHandler

    Set objRSParamReprocCC = ObterParamReprocCC()
    
    If objRSParamReprocCC.EOF Then
        Set objRSParamReprocCC = Nothing
        Processar = True
        Exit Function
    End If
    
    While Not objRSParamReprocCC.EOF
        
        'verifica se está dentro da grade de horario (da teimosinha ou de integração sistema),
        'e caso esteja verifica se está na hora de reprocessar de acordo com o intervalo configurado na teimosinha
        If VerificarGradeHorario(objRSParamReprocCC!CD_EMPR, IIf(IsNull(objRSParamReprocCC!HO_LIMI_ITGR), gstrDataVazia, objRSParamReprocCC!HO_LIMI_ITGR)) Then
            
            datDataOper = fgDataHoraServidor(DataAux)
            
            Set objRSLanc = ObterLancamentos(objRSParamReprocCC!CD_EMPR, _
                                             objRSParamReprocCC!TP_OPER, _
                                             objRSParamReprocCC!IN_LANC_DEBT_CRED, _
                                             objRSParamReprocCC!TP_BKOF, _
                                             IIf(IsNull(objRSParamReprocCC!TP_CNAL_VEND), 0, objRSParamReprocCC!TP_CNAL_VEND), _
                                             Format$(datDataOper, "YYYYMMDD"))
                        
            While Not objRSLanc.EOF
            
                datHoraUltiReproc = objRSLanc!DH_ULTI_REPR_CC
                'FREITAS - 14/07/2010 - o Controle de horário agora é por minuto.
                'HORAS
                'datHoraUltiReproc = DateAdd("h", objRSParamReprocCC!QT_HORA_INTL_REPR, datHoraUltiReproc)
                'MINUTOS
                datHoraUltiReproc = Format(DateAdd("n", objRSParamReprocCC!QT_HORA_INTL_REPR, datHoraUltiReproc), "hh:mm:ss")
                
                datHoraAtual = Format(fgDataHoraServidor(DataHoraAux), "hh:mm:ss")
                
                If datHoraAtual >= datHoraUltiReproc Then
                    
                    'FREITAS - 14/07/2010 - Não será mais necessário controlar a quantidade de vezes do reprocessamento,
                    'pois o sistema irá reprocessar o lançamento até que o horário limite seja excedido.
                    
                    'If objRSLanc!qt_repr_cc < objRSParamReprocCC!QT_REPR_CNTA_CRRT Then
                        Set objGerenciadorReprocCCT = CreateObject("A8LQS.clsGerenciadorReprocCCT")
                        strErro = objGerenciadorReprocCCT.Gerenciar(objRSLanc!NU_SEQU_OPER_ATIV, _
                                                                    objRSLanc!IN_LANC_DEBT_CRED, _
                                                                    objRSLanc!NR_SEQU_LANC)
                        Set objGerenciadorReprocCCT = Nothing
                    'End If
                End If
                
                objRSLanc.MoveNext
            Wend
        
        Else 'verifica quando foi feito o último reprocessamento, se ele tiver ocorrido antes do encerramento da grade, então permite-se que seja feito mais um reprocessamento final
        
            datDataOper = fgDataHoraServidor(DataAux)
            datHoraGradeHorario = BuscarGradeHorario(objRSParamReprocCC!CD_EMPR, IIf(IsNull(objRSParamReprocCC!HO_LIMI_ITGR), gstrDataVazia, objRSParamReprocCC!HO_LIMI_ITGR))
            
            Set objRSLanc = ObterLancamentos(objRSParamReprocCC!CD_EMPR, _
                                             objRSParamReprocCC!TP_OPER, _
                                             objRSParamReprocCC!IN_LANC_DEBT_CRED, _
                                             objRSParamReprocCC!TP_BKOF, _
                                             IIf(IsNull(objRSParamReprocCC!TP_CNAL_VEND), 0, objRSParamReprocCC!TP_CNAL_VEND), _
                                             Format$(datDataOper, "YYYYMMDD"))
            
            While Not objRSLanc.EOF
                
                datHoraUltiReproc = TimeSerial(Hour(objRSLanc!DH_ULTI_REPR_CC), Minute(objRSLanc!DH_ULTI_REPR_CC), Second(objRSLanc!DH_ULTI_REPR_CC))
            
                If datHoraGradeHorario >= datHoraUltiReproc Then
                    
                    Set objGerenciadorReprocCCT = CreateObject("A8LQS.clsGerenciadorReprocCCT")
                    strErro = objGerenciadorReprocCCT.Gerenciar(objRSLanc!NU_SEQU_OPER_ATIV, _
                                                                objRSLanc!IN_LANC_DEBT_CRED, _
                                                                objRSLanc!NR_SEQU_LANC)
                    Set objGerenciadorReprocCCT = Nothing
                    
                End If
            
                objRSLanc.MoveNext
            Wend
        End If
    
        objRSParamReprocCC.MoveNext
    Wend
    
    
    Processar = vbNullString

    Exit Function

ErrorHandler:
    
    strErro = Err.Description
    Err.Clear
    
    Set objGerenciadorReprocCCT = Nothing
    Set objRSParamReprocCC = Nothing
    Set objRSLanc = Nothing
    
    Processar = strErro
    
End Function

Private Function ObterLancamentos(ByVal pintCO_EMPR As Integer, _
                                  ByVal pintTP_OPER As Integer, _
                                  ByVal pintIN_LANC_DEBT_CRED As Integer, _
                                  ByVal pintTP_BKOF As Integer, _
                                  ByVal pintTP_CNAL_VEND As Integer, _
                                  ByVal pDT_OPER As String) As ADODB.Recordset

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " Select   A.NU_SEQU_OPER_ATIV, " & vbNewLine & _
             "          A.TP_LANC_ITGR, " & vbNewLine & _
             "          A.TP_BKOF, " & vbNewLine & _
             "          A.CO_EMPR, " & vbNewLine & _
             "          A.CO_VEIC_LEGA, " & vbNewLine & _
             "          A.TP_OPER, " & vbNewLine & _
             "          A.CO_LOCA_LIQU, " & vbNewLine & _
             "          A.IN_LANC_DEBT_CRED, " & vbNewLine & _
             "          A.DT_OPER, " & vbNewLine & _
             "          A.CO_BANC, " & vbNewLine & _
             "          A.CO_AGEN, " & vbNewLine & _
             "          A.NU_CC, " & vbNewLine & _
             "          A.NU_COMD_OPER, " & vbNewLine & _
             "          A.VA_LANC_CC, " & vbNewLine & _
             "          A.NU_CTRL_ITGR, " & vbNewLine & _
             "          A.CO_ULTI_SITU_PROC, " & vbNewLine & _
             "          A.TP_CNAL_VEND, " & vbNewLine & _
             "          A.DH_ULTI_ATLZ, " & vbNewLine & _
             "          NVL(A.DH_ULTI_REPR,A.DH_ULTI_ATLZ) AS DH_ULTI_REPR_CC, " & vbNewLine & _
             "          NVL(A.QT_REPR_CNTA_CRRT,0) AS qt_repr_cc, " & vbNewLine & _
             "          A.NR_SEQU_LANC " & vbNewLine & " "
    strSQL = strSQL & _
             " From     A8.TB_LANC_CC_CNTB A    " & vbNewLine & _
             " Where    A.CO_ULTI_SITU_PROC   = " & enumStatusIntegracao.ErroSaldoCC & vbNewLine & _
             " And      A.TP_LANC_ITGR        = " & enumTipoLancamentoIntegracao.Normal & vbNewLine & _
             " And      A.IN_LANC_DEBT_CRED   = " & pintIN_LANC_DEBT_CRED & vbNewLine & _
             " And      A.CO_EMPR             = " & pintCO_EMPR & vbNewLine & _
             " And      A.TP_OPER             = " & pintTP_OPER & vbNewLine & _
             " And      A.TP_BKOF             = " & pintTP_BKOF & vbNewLine & _
             " And      TRUNC(A.DH_ULTI_ATLZ) = " & fgDtXML_To_Oracle(pDT_OPER)
             
    If pintTP_CNAL_VEND = enumCanalDeVenda.Nenhum Then
        strSQL = strSQL & _
            " And      (A.TP_CNAL_VEND = 0 or A.TP_CNAL_VEND IS NULL)"
    Else
        strSQL = strSQL & _
            " And      A.TP_CNAL_VEND        = " & pintTP_CNAL_VEND & vbNewLine
    
    End If

    Set ObterLancamentos = fgQuerySQL(strSQL)

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
'retorna TRUE caso ainda esteja dentro do horário limite (da teimosinha ou de integração do BG)
'retorna FALSE se estiver fora do horário limite (da teimosinha ou de integração do BG)
Private Function VerificarGradeHorario(ByVal plngCodifgoEmpresa As Long, ByVal pdatHorarioLimiteTeimosinha As Date) As Boolean

Dim strSQL                                  As String
Dim objHoraLimite                           As ADODB.Recordset
Dim datHoraServidor                         As Date
Dim datHoraLimite                           As Date

On Error GoTo ErrorHandler

    If pdatHorarioLimiteTeimosinha = gstrDataVazia Then 'se nao existir horario limite teimosinha, entao obtem horario limite integracao BG
    
        strSQL = " SELECT   HO_LIMI_ENVI_ITGR         " & vbNewLine & _
                 " FROM     A8.TB_HORA_LIMI_ITGR_SIST " & vbNewLine & _
                 " WHERE    SG_SIST  = 'BG'" & _
                 " AND      CO_EMPR = " & plngCodifgoEmpresa
    
        Set objHoraLimite = fgQuerySQL(strSQL)
    
        If objHoraLimite.RecordCount = 0 Then
            VerificarGradeHorario = False
            Exit Function
        End If

        datHoraLimite = TimeSerial(Hour(objHoraLimite!HO_LIMI_ENVI_ITGR), Minute(objHoraLimite!HO_LIMI_ENVI_ITGR), 0)

    Else 'obtem horario limite teimosinha
    
        datHoraLimite = TimeSerial(Hour(pdatHorarioLimiteTeimosinha), Minute(pdatHorarioLimiteTeimosinha), 0)
    
    End If

    datHoraServidor = TimeSerial(Hour(fgDataHoraServidor(enumFormatoDataHora.Hora)), Minute(fgDataHoraServidor(enumFormatoDataHora.Hora)), 0)
  
    If datHoraLimite < datHoraServidor Then
        VerificarGradeHorario = False
    Else
        VerificarGradeHorario = True
    End If

    Exit Function

ErrorHandler:

    VerificarGradeHorario = False

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "VerificarGradeHorario Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function BuscarGradeHorario(ByVal plngCodigoEmpresa As Long, ByVal pdatHorarioLimiteTeimosinha As Date) As Date

Dim objHoraLimite                           As ADODB.Recordset
Dim strSQL                                  As String
Dim datHoraLimite                           As Date

On Error GoTo ErrorHandler

    If pdatHorarioLimiteTeimosinha = gstrDataVazia Then 'se nao existir horario limite teimosinha, entao obtem horario limite integracao sistema

        strSQL = " SELECT   HO_LIMI_ENVI_ITGR         " & vbNewLine & _
                 " FROM     A8.TB_HORA_LIMI_ITGR_SIST " & vbNewLine & _
                 " WHERE    SG_SIST  = 'BG'" & _
                 " AND      CO_EMPR = " & plngCodigoEmpresa
    
        Set objHoraLimite = fgQuerySQL(strSQL)
    
        If objHoraLimite.RecordCount = 0 Then
            datHoraLimite = TimeSerial(0, 0, 0)
        Else
            datHoraLimite = TimeSerial(Hour(objHoraLimite!HO_LIMI_ENVI_ITGR), Minute(objHoraLimite!HO_LIMI_ENVI_ITGR), Second(objHoraLimite!HO_LIMI_ENVI_ITGR))
        End If
    
    Else
    
        datHoraLimite = TimeSerial(Hour(pdatHorarioLimiteTeimosinha), Minute(pdatHorarioLimiteTeimosinha), 0)
    
    End If
    
    BuscarGradeHorario = datHoraLimite

Exit Function
ErrorHandler:

    BuscarGradeHorario = Null

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "BuscarGradeHorario Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

