VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA6A7A8Funcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objetivo       : Interface de Funções Genéricas para A6 , A7 e A8

Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Private strConstruct                        As String

Private udtDRAT                             As tA8_Feriado.DRAT
Private udtEntrada                          As tA8_Feriado.ENTRADA

Implements COMSVCSLib.ObjectControl
Implements COMSVCSLib.IObjectConstruct

Private objContext As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

Private Sub IObjectConstruct_Construct(ByVal pobjContruct As Object)
        
'Parametro Habilita/Desabilita chamada para a validação de data util
'Rotina : A6A7A8COMTI.clsConsulta , componete COMTI
        
    strConstruct = pobjContruct.ConstructString

End Sub

'---------------------------------------------------------------------------------------------
'Função que retorna a data (pdatData) a quantidade de dias passado no parametro 'pintQtdeDias'
'---------------------------------------------------------------------------------------------

Public Function AdicionarDiasUteis(ByVal pdatData As Date, _
                                   ByVal pintQtdeDias As Integer, _
                                   ByVal plngMovimento As enumPaginacao) As Date

Dim intQtdeDiasValidos                      As Integer
Dim intIncremento                           As Integer
Dim datRetorno                              As Date
Dim vntArray()                              As Variant
    
On Error GoTo ErrorHandler

    vntArray = Array(pdatData, pintQtdeDias, plngMovimento)
    datRetorno = fgExecuteCMD("A8PROC.A8P_ADICIONA_DIAS_UTEIS", _
                              3, _
                              vntArray)

    AdicionarDiasUteis = datRetorno
    
Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, _
                      TypeName(Me), _
                      "AdicionarDiasUteis Function", _
                      lngCodigoErroNegocio, _
                      intNumeroSequencialErro)

End Function

'Retorna o valor do XML de Parâmetros Gerais, a partir da Tag passada como parâmetro para a função
Public Function ObterValorParametrosGerais(ByVal pstrTagParametro As String) As String

Dim xmlParametroGeral                       As MSXML2.DOMDocument40
    
    On Error GoTo ErrorHandler

    Set xmlParametroGeral = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlParametroGeral.loadXML(fgSelectVarchar4000(0, False))

    If xmlParametroGeral.xml <> vbNullString Then
        If Not xmlParametroGeral.selectSingleNode("//" & pstrTagParametro) Is Nothing Then
            ObterValorParametrosGerais = xmlParametroGeral.selectSingleNode("//" & pstrTagParametro).Text
        End If
    End If
    Set xmlParametroGeral = Nothing
    
Exit Function
ErrorHandler:
    
    Set xmlParametroGeral = Nothing
    ObterValorParametrosGerais = vbNullString
    
End Function

'-------------------------------------------------------------------------------------
'------- METODOS PARA CHAMADA DO COMTI DE CONSULTA DE FERIADOS -----------------------
'-------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------
'Função que retorna a data util posterior a Data Base, considernado Banco e Agencia
'---------------------------------------------------------------------------------------------

Public Function ValidaDataUtil(ByVal pintCodigoBanco As Integer, _
                               ByVal plngAgencia As Long, _
                               ByVal pdatDataBase As Date, _
                               ByRef plngCodErro As Long, _
                               ByRef pstrMensagemErro As String) As Date

Dim strDataIni                              As String
Dim strRet                                  As String

    On Error GoTo ErrorHandler
    
    #If DebugMode = 1 Then
        ValidaDataUtil = pdatDataBase
    #Else

        If UCase(Trim(strConstruct)) = "OFF" Then
            ValidaDataUtil = pdatDataBase
        Else

            strDataIni = Format$(pdatDataBase, "DDMMYYYY")

            strRet = "S"
            Do While strRet <> "N"

                If Weekday(pdatDataBase) = vbSaturday Then
                    pdatDataBase = DateAdd("d", 2, pdatDataBase)
                ElseIf Weekday(pdatDataBase) = vbSunday Then
                    pdatDataBase = DateAdd("d", 1, pdatDataBase)
                End If

                strDataIni = Format$(pdatDataBase, "DDMMYYYY")

                strRet = flPesquisar(pintCodigoBanco, _
                                     plngAgencia, _
                                     strDataIni, _
                                     plngCodErro, _
                                     pstrMensagemErro)

                If Trim(pstrMensagemErro) <> "" Then
                    ValidaDataUtil = pdatDataBase
                    strRet = "N"
                Else
                    If strRet = "N" Then
                        ValidaDataUtil = pdatDataBase
                    Else
                        pdatDataBase = DateAdd("d", 1, pdatDataBase)
                    End If
                End If
            Loop
        End If
    #End If
    
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, _
                      TypeName(Me), _
                      "ValidaDataUtil", _
                      lngCodigoErroNegocio, _
                      intNumeroSequencialErro)
    
End Function

Private Function flObterDescricaoErroFX(ByVal plngCodErro As Long) As String

Dim strMsgErro                              As String

    Select Case plngCodErro
        Case 1: strMsgErro = "CODIGO DA TRANSACAO INVALIDO"
        Case 2: strMsgErro = "OPCAO DE ENTRADA INVALIDA"
        Case 3: strMsgErro = "BANCO INVALIDO"
        Case 4: strMsgErro = "CODIGO DE AREA GEOGRAFICA INVALIDO"
        Case 5: strMsgErro = "AREA GEOGRAFICA NAO CADASTRADA"
        Case 6: strMsgErro = "AGENCIA INVALIDA"
        Case 7: strMsgErro = "CODIGO MUNICIPIO BACEN NAO CADASTRADO"
        Case 8: strMsgErro = "DATA INICIAL NAO NUMERICA"
        Case 9: strMsgErro = "DATA INICIAL INVALIDA"
        Case 10: strMsgErro = "DATA FINAL NAO NUMERICA"
        Case 11: strMsgErro = "DATA FINAL INVALIDA"
        Case 12: strMsgErro = "DATA FINAL MENOR DO QUE A INICIAL"
        Case 13: strMsgErro = "AGEO NAO POSSUI FERIADO NO PERIODO SOLICITADO"
        Case 14: strMsgErro = "AGEO POSSUI MAIS DE 30 FERIADOS NO PERIODO SOLICITADO (INFORMADO OS 30 PRIMEIROS)"
        Case 99: strMsgErro = "ERRO DE ACESSO AO DB2"
        Case Else
            strMsgErro = "ERRO CONSULTA FERIADO COMTI"
    End Select
    
    flObterDescricaoErroFX = strMsgErro

End Function

Private Function flPesquisar(ByVal strBanco As String, _
                             ByVal strAgencia As String, _
                             ByVal strDataIni As String, _
                             ByRef plngCodErro As Long, _
                             ByRef pstrMsgErro As String) As String
   
Dim objA8             As Object 'tA8_Feriado.Servico
Dim strComplemento    As String
Dim ObjRsErro         As ADODB.Recordset
Dim ObjRet            As ADODB.Recordset

On Error GoTo ErrorHandler
    
    '--- inicializar as UDT's
    flInicializarUDTs
    
    '--- area para COMMAREA
    udtDRAT.DR_APLCINI_0084 = "S"
    udtDRAT.DR_PROGINI_0084 = "A8AC1005"  ' DK210_NM_PRG
    udtDRAT.DR_TRANINI_0084 = "A801"      ' DK210_TRANSID
    udtDRAT.DR_COMMINI_0084 = 70
    udtDRAT.DR_COMMTOT_0084 = 70
    
    '--- area comum para retorno de TS
    udtDRAT.DR_NUMTS_0084 = 1
    
    '--- area para a primeira TS
    'Este parametro determina se o nome da TS é gerado
    udtDRAT.DR_OPCTS1_0084 = "G"
    udtDRAT.DR_TAMIT1_0084 = 70
    udtDRAT.DR_MAXIT1_0084 = 99999
    udtDRAT.DR_DELFI1_0084 = "S"
    udtDRAT.DR_DELIN1_0084 = "S"
    
    '--- parametros para a rotina (DK210)
    udtEntrada.A8CE1005_E_CO_AGEN = strAgencia
    udtEntrada.A8CE1005_E_CO_BANC = strBanco
    udtEntrada.A8CE1005_E_DATA_FIM = strDataIni
    udtEntrada.A8CE1005_E_DATA_INICIAL = strDataIni
    
    '--- executar
    
    Set objA8 = CreateObject("tA8_Feriado.Servico")
    
    If objA8 Is Nothing Then
        strComplemento = "Componente 'tA8_Feriado.Servico - Valida Data Util' não instalado "
    End If

    Set ObjRet = objA8.A8AC1005(udtDRAT, _
                                udtEntrada, _
                                ObjRsErro)
    Set objA8 = Nothing
    
    If ObjRsErro(0) <> 0 Then
        If IsNumeric(ObjRsErro(0)) Then
            plngCodErro = ObjRsErro(0)
        Else
            plngCodErro = 0
        End If
        
        pstrMsgErro = flObterDescricaoErroFX(plngCodErro)
        flPesquisar = "N"
    Else
        pstrMsgErro = ""
        plngCodErro = 0
        If Trim(ObjRet.Fields(0)) = "" Then
            flPesquisar = "N"
        Else
            flPesquisar = "S"
        End If
    End If
    
    Set ObjRsErro = Nothing
    Set ObjRet = Nothing
    
Exit Function
ErrorHandler:
    
    Set ObjRsErro = Nothing
    Set ObjRet = Nothing
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    
    Call fgRaiseError(App.EXEName, _
                      TypeName(Me), _
                      "flPesquisar", _
                      lngCodigoErroNegocio, _
                      intNumeroSequencialErro, _
                      strComplemento)
    
End Function

Private Sub flInicializarUDTs()
    
    udtDRAT.DR_CODRET_0084 = 0
    udtDRAT.DR_CODPRB_0084 = Space(4)
    udtDRAT.DR_APLCINI_0084 = Space(1)
    udtDRAT.DR_PROGINI_0084 = Space(8)
    udtDRAT.DR_TRANINI_0084 = Space(4)
    udtDRAT.DR_COMMINI_0084 = 0
    udtDRAT.DR_NMSYINI_0084 = Space(8)
    udtDRAT.DR_CICSINI_0084 = Space(4)
    udtDRAT.DR_TIPCHMI_0084 = Space(1)
    udtDRAT.DR_PLATCALL_0084 = Space(1)
    udtDRAT.DR_BCOCALL_0084 = 0
    udtDRAT.DR_EXPIRATS_0084 = 0
    udtDRAT.DR_NUMTS_0084 = 0
    udtDRAT.DR_OPREGTS_0084 = 0
    udtDRAT.DR_NMSYTS_0084 = Space(8)
    udtDRAT.DR_CICSTS_0084 = Space(4)
    udtDRAT.DR_OPCTS1_0084 = Space(1)
    udtDRAT.DR_NOMETS1_0084 = Space(8)
    udtDRAT.DR_TAMIT1_0084 = 0
    udtDRAT.DR_MAXIT1_0084 = 0
    udtDRAT.DR_ULTIT1_0084 = 0
    udtDRAT.DR_DELFI1_0084 = Space(1)
    udtDRAT.DR_DELIN1_0084 = Space(1)
    udtDRAT.DR_RETTS1_0084 = 0
    udtDRAT.DR_NITEM1_0084 = 0
    udtDRAT.DR_NITBX1_0084 = 0
    udtDRAT.DR_OPCTS2_0084 = Space(1)
    udtDRAT.DR_NOMETS2_0084 = Space(8)
    udtDRAT.DR_TAMIT2_0084 = 0
    udtDRAT.DR_MAXIT2_0084 = 0
    udtDRAT.DR_ULTIT2_0084 = 0
    udtDRAT.DR_DELFI2_0084 = Space(1)
    udtDRAT.DR_DELIN2_0084 = Space(1)
    udtDRAT.DR_RETTS2_0084 = 0
    udtDRAT.DR_NITEM2_0084 = 0
    udtDRAT.DR_NITBX2_0084 = 0
    udtDRAT.DR_OPCTS3_0084 = Space(1)
    udtDRAT.DR_NOMETS3_0084 = Space(8)
    udtDRAT.DR_TAMIT3_0084 = 0
    udtDRAT.DR_MAXIT3_0084 = 0
    udtDRAT.DR_ULTIT3_0084 = 0
    udtDRAT.DR_DELFI3_0084 = Space(1)
    udtDRAT.DR_DELIN3_0084 = Space(1)
    udtDRAT.DR_RETTS3_0084 = 0
    udtDRAT.DR_NITEM3_0084 = 0
    udtDRAT.DR_NITBX3_0084 = 0
    udtDRAT.DR_GERATSO_0084 = Space(1)
    udtDRAT.DR_NOMETSO_0084 = Space(8)
    udtDRAT.DR_NMSYTSO_0084 = Space(8)
    udtDRAT.DR_CICSTSO_0084 = Space(4)
    udtDRAT.DR_NUMTSI_0084 = 0
    udtDRAT.DR_OPREGTSI_0084 = 0
    udtDRAT.DR_NMSYTSI_0084 = Space(8)
    udtDRAT.DR_CICSTSI_0084 = Space(4)
    udtDRAT.DR_DELFIMI_0084 = Space(1)
    udtDRAT.DR_OPCTS1I_0084 = Space(1)
    udtDRAT.DR_NOMETS1I_0084 = Space(8)
    udtDRAT.DR_TAMIT1I_0084 = 0
    udtDRAT.DR_QTDIT1I_0084 = 0
    udtDRAT.DR_OPCTS2I_0084 = Space(1)
    udtDRAT.DR_NOMETS2I_0084 = Space(8)
    udtDRAT.DR_TAMIT2I_0084 = 0
    udtDRAT.DR_QTDIT2I_0084 = 0
    udtDRAT.DR_APLCFIM_0084 = Space(1)
    udtDRAT.DR_PROGFIM_0084 = Space(8)
    udtDRAT.DR_TRANFIM_0084 = Space(4)
    udtDRAT.DR_COMMFIM_0084 = 0
    udtDRAT.DR_NMSYFIM_0084 = Space(8)
    udtDRAT.DR_CICSFIM_0084 = Space(4)
    udtDRAT.DR_TIPCHMF_0084 = Space(1)
    udtDRAT.DR_OBEXFIM_0084 = Space(1)
    udtDRAT.DR_EXECFIM_0084 = Space(1)
    udtDRAT.DR_OPCCOMM_0084 = Space(1)
    udtDRAT.DR_COMMTOT_0084 = 0
    
End Sub

Public Function ObterSYSDATE_Oracle() As Date


Dim strSQL                                   As String
Dim objRS                                    As ADODB.Recordset
Dim strRetorno                               As String
Dim strAux                                   As String
    
On Error GoTo ErrorHandler

    strSQL = "select to_char(sysdate,'yyyymmddhh24miss')  as DH_SERV from dual"

    Set objRS = fgQuerySQL(strSQL)
    
    
    strRetorno = objRS!DH_SERV
    
           
    ObterSYSDATE_Oracle = fgDtHrStr_To_DateTime(strRetorno)
             
             
    
    

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    
    Call fgRaiseError(App.EXEName, _
                      TypeName(Me), _
                      "ObterSYSDATE_Oracle Function", _
                      lngCodigoErroNegocio, _
                      intNumeroSequencialErro)
End Function
