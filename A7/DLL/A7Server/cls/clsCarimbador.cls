VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarimbador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo envio do "carimbo" ao monitor central do sistema NZ.
Option Explicit

'Public Declare Function GravaCarimbo Lib "Carimbo" (ByVal Msg As String) As Byte

Private Type udtCarimbo
    CO_EMPR                                 As String * 5
    NU_CTRL_REME                            As String * 20
    CO_MESG                                 As String * 9
    TP_FLUX_MESG                            As String * 1
    SITU_MESG                               As String * 3
    ORIG_MESG                               As String * 3
    DEST_MESG                               As String * 3
    NO_FILA_DEST                            As String * 48
    ID_SINA_VALO                            As String * 1
    TAG_VALOR                               As String * 18
    TAG_SITLANC                             As String * 20
    CO_ISPB_PKI                             As String * 8
    TAG_DTHR                                As String * 14
    TAG_DTMOVTO                             As String * 8
    FILLER                                  As String * 39
End Type

Private Type udtCarimboAux
    String                                  As String * 200
End Type

Public Enum enumStatusMonitor
    MONITOR_ENVIADA_NZ = 1
    MONITOR_ENVIADA_PK = 2
    MONITOR_ENVIADA_PILOTO = 3
    MONITOR_ERRO_PKI = 4
    MONITOR_ENVIADA_EXTERNO = 5
    MONITOR_RECEB_COA = 6
    MONITOR_RECEB_COD = 7
    MONITOR_ENVIADO_R1_NZ = 11
    MONITOR_ENVIADO_R1_LEGADO = 12
    MONITOR_RECEB_R1 = 13
    MONITOR_RECEB_MSG_REQ_CMAX = 14
    MONITOR_ENV_ERRO_EXT_NZ = 21
    MONITOR_ENV_ERRO_EXT_LEGADO = 22
    MONITOR_RECEB_ERRO_EXTERNO = 23
    MONITOR_ENV_CANCEL_PILOTO_NZ = 31
    MONITOR_ENV_CANCEL_PILOTO_LEGADO = 32
    MONITOR_RECEB_CANCEL_PILOTO = 33
    MONITOR_ENV_ERRO_PKNZ = 41
    MONITOR_ENV_ERRO_PKLEGADO = 42
    MONITOR_RECEB_ERRO_PK = 43
    MONITOR_ENV_ERRO_NZ = 51
    MONITOR_RECEB_ERRO_NZ = 52
    MONITOR_ENV_DUPLIC = 61
    MONITOR_RECEB_DUPLIC = 62
    MONITOR_ENV_MENS_EXTERNA_NZ = 71
    MONITOR_ENV_MENS_EXTERNA_LEGADO = 72
    MONITOR_RECEB_MENS_EXTERNA = 73
    MONITOR_ENV_CARGA_NZ = 81
    MONITOR_CARGA_EFETUADA = 82
    MONITOR_REGUL_ERRO_PKI = 0
    MONITOR_ERRO_PKI_MSG_EXT = 93
End Enum

Public Enum enumFluxoMonitor
    FLUXO_MONITOR_NORMAL = 1
    FLUXO_MONITOR_LEGADINHO = 2
    FLUXO_MONITOR_CONT_MAXIMA = 3
    FLUXO_MONITOR_MSG_EXTERNA = 4
    FLUXO_MONITOR_MSG_CARGA = 5
End Enum

Private llCodigoErroNegocio                 As Long
Private liNumeroSequencialErro              As Integer

Private lsConstruct                         As String

Private objContext                          As COMSVCSLib.ObjectContext

Implements COMSVCSLib.ObjectControl
Implements COMSVCSLib.IObjectConstruct
Private Sub IObjectConstruct_Construct(ByVal pCtorObj As Object)
On Error GoTo ErrHandler
    lsConstruct = pCtorObj.ConstructString
ErrHandler:
End Sub

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

'Enviar o carimbo ao sistema NZ através de put em fila MQSeries.
Public Sub IncluirAvisoMonitor(ByRef objMQAX200 As A6A7A8CA.clsMQAX200, _
                               ByVal plngCodigoEmpresa As Long, _
                               ByVal pstrNU_CTRL_REME As String, _
                               ByVal pstrCodigoMensagem As String, _
                               ByVal plngTipoFluxoMensagem As enumFluxoMonitor, _
                               ByVal plngSituacaoMensagem As enumStatusMonitor, _
                               ByVal pstrSistemaOrigem As String, _
                               ByVal pstrSistemaDestino As String, _
                               ByVal pstrNomeFilaDestino As String, _
                               ByVal pdatDataMovto As Date, _
                      Optional ByVal pvntValor As Variant, _
                      Optional ByVal pstrSituacaoLancamento As String, _
                      Optional ByVal pstrCodigoISPB_ErroPKI As String, _
                      Optional ByVal pdatDataHora_Ext_R1 As Date)

'ID_SINA_VALO           informar avaliando o valor informado
'pdDataHora_Ext_R1      Somente para Mensagens R1 ou Externas.
    
Dim udtCarimbo                                As udtCarimbo
Dim udtCarimboAux                             As udtCarimboAux
Dim xmlFilas                                  As MSXML2.DOMDocument40
Dim strNomeFilaMonitor                      As String
    
On Error Resume Next

    If InStr(1, UCase(lsConstruct), "OFF") > 0 Then Exit Sub

    udtCarimbo.CO_EMPR = Format(plngCodigoEmpresa, "00000")
    udtCarimbo.NU_CTRL_REME = Left(Trim(pstrNU_CTRL_REME) & String(20, " "), 20)
    udtCarimbo.CO_MESG = Left(Trim(pstrCodigoMensagem) & String(9, " "), 9)
    udtCarimbo.TP_FLUX_MESG = flFluxoMonitor(plngTipoFluxoMensagem)
    udtCarimbo.SITU_MESG = Format(plngSituacaoMensagem, "000")

    udtCarimbo.ORIG_MESG = Left(Trim(pstrSistemaOrigem) & String(3, " "), 3)
    udtCarimbo.DEST_MESG = Left(Trim(pstrSistemaDestino) & String(3, " "), 3)
    udtCarimbo.NO_FILA_DEST = Left(Trim(pstrNomeFilaDestino) & String(48, " "), 48)

    If Trim$(pvntValor) = vbNullString Then
        
        udtCarimbo.ID_SINA_VALO = "0"
        udtCarimbo.TAG_VALOR = String(18, "0")
        
    ElseIf Not IsMissing(pvntValor) Then
        
        If CDec(pvntValor) < 0 Then
            udtCarimbo.ID_SINA_VALO = "-"
        Else
            udtCarimbo.ID_SINA_VALO = "0"
        End If
        
        udtCarimbo.TAG_VALOR = Right(String(18, "0") & _
                           Replace(Replace( _
                           FormatNumber(CDec(pvntValor), 2, , , vbFalse), _
                           ",", vbNullString), ".", vbNullString), _
                           18)
    Else
        udtCarimbo.ID_SINA_VALO = "0"
        udtCarimbo.TAG_VALOR = String(18, "0")
    End If

    udtCarimbo.TAG_SITLANC = Left(Trim(pstrSituacaoLancamento) & String(20, " "), 20)
    udtCarimbo.CO_ISPB_PKI = Left(Trim(pstrCodigoISPB_ErroPKI) & String(8, " "), 8)

    'tratamento data hora externa
    If pdatDataHora_Ext_R1 = gstrDataVazia Then
        udtCarimbo.TAG_DTHR = "00000000000000"
    Else
        udtCarimbo.TAG_DTHR = Format(pdatDataHora_Ext_R1, "YYYYMMDDHHMMSS")
    End If

    'tratamento de data movimento
    If pdatDataMovto = gstrDataVazia Then
        udtCarimbo.TAG_DTMOVTO = "00000000"
    Else
        udtCarimbo.TAG_DTMOVTO = Format(pdatDataMovto, "YYYYMMDD")
    End If

    udtCarimbo.FILLER = String(39, " ")

    LSet udtCarimboAux = udtCarimbo
    
    If objMQAX200 Is Nothing Then
        ObjectControl_Activate
    End If
    
    Set xmlFilas = CreateObject("MSXML2.DOMDocument.4.0")
    
   'RATS 364 - Ocor 01
   'Paremetrizar Nome da fila com
        '<NomeFilaMonitor Empresa='523'>NZQ.E.MONITOR</SufixoFilaMonitor>
        '<NomeFilaMonitor Empresa='701'>NZQ.E.MONITOR</SufixoFilaMonitor>
    
    If xmlFilas.Load(App.Path & "\FilasNET.xml") Then
        If plngCodigoEmpresa = enumCodigoEmpresa.Banespa Then
            If Not xmlFilas.selectSingleNode("//NomeFilaMonitor[@Empresa='" & enumCodigoEmpresa.Banespa & "']") Is Nothing Then
                strNomeFilaMonitor = xmlFilas.selectSingleNode("//NomeFilaMonitor[@Empresa='" & plngCodigoEmpresa & "']").Text
            Else
                strNomeFilaMonitor = "NZQ.E.MONITOR"
            End If
        Else
            If Not xmlFilas.selectSingleNode("//NomeFilaMonitor[@Empresa='" & enumCodigoEmpresa.Santander & "']") Is Nothing Then
                strNomeFilaMonitor = xmlFilas.selectSingleNode("//NomeFilaMonitor[@Empresa='" & plngCodigoEmpresa & "']").Text
            Else
                strNomeFilaMonitor = "NZQ.E.MONITOR"
            End If
        End If
    Else
        strNomeFilaMonitor = "NZQ.E.MONITOR"
    End If
    
    Set xmlFilas = Nothing
    
    With objMQAX200
        .MQOpen strNomeFilaMonitor, MQOO_OUTPUT
        .MQPut udtCarimboAux.String, 200, MainFrame
        .MQClose
    End With

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    'GravaCarimbo udtCarimboAux.String
    
    Exit Sub
ErrorHandler:
    
    Set xmlFilas = Nothing
    
    If llCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluirAvisoMonitor Sub", llCodigoErroNegocio, liNumeroSequencialErro)
End Sub

'Converter o domínio de FluxoMonitor pelas literais esperadas pelo sistema NZ.
Private Function flFluxoMonitor(ByVal plFluxo As enumFluxoMonitor) As String

    Select Case plFluxo
        Case enumFluxoMonitor.FLUXO_MONITOR_NORMAL
            flFluxoMonitor = "N"
        Case enumFluxoMonitor.FLUXO_MONITOR_LEGADINHO
            flFluxoMonitor = "L"
        Case enumFluxoMonitor.FLUXO_MONITOR_CONT_MAXIMA
            flFluxoMonitor = "C"
        Case enumFluxoMonitor.FLUXO_MONITOR_MSG_EXTERNA
            flFluxoMonitor = "E"
        Case enumFluxoMonitor.FLUXO_MONITOR_MSG_CARGA
            flFluxoMonitor = "K"
    End Select

    'Public Const FLUXO_MONITOR_NORMAL           As String = "N"
    'Public Const FLUXO_MONITOR_LEGADINHO        As String = "L"
    'Public Const FLUXO_MONITOR_CONT_MAXIMA      As String = "C"
    'Public Const FLUXO_MONITOR_MSG_EXTERNA      As String = "E" '(R2, AVISO, INFORMA)
    'Public Const FLUXO_MONITOR_MSG_CARGA        As String = "K"

End Function




