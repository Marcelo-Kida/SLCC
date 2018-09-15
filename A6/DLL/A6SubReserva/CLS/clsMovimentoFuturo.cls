VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMovimentoFuturo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA\r\nPacote         :\r\nClasse         : clsMovimentoFuturo\r\nData Criação   : 14/07/2003\r\nObjetivo       :\r\n\r\nAnalista       : Carlos Fortes\r\n\r\nProgramador    : Carlos Fortes\r\nData           : 14/07/2003\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F105A9E0027"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela consulta e atualização
' do Cadastro de Movimento Futuro.

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

' Este método é responsável pela leitura de propriedades da tabela de movimento futuro,
' e pelo retorno à camada intermediária.

Public Function ObterPropriedades() As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   SG_SIST," & vbCrLf & _
             " CO_REME_MOVI_FUTU, " & vbCrLf & _
             " DH_REME, " & vbCrLf & _
             " CO_EMPR, " & vbCrLf & _
             " DT_FECH_PROC, " & vbCrLf & _
             " CO_LOCA_LIQU, " & vbCrLf & _
             " DE_TIPO_LIQU, " & vbCrLf & _
             " DT_LIQU_OPER, " & vbCrLf & _
             " IN_MOVI_ENTR_SAID, " & vbCrLf & _
             " VA_LIQU_OPER, " & vbCrLf & _
             " DE_ATIV, " & vbCrLf & _
             " CO_VEIC_LEGA, " & vbCrLf & _
             " CO_ITEM_CAIX, " & vbCrLf & _
             " CO_CNPJ_CNPT, " & vbCrLf & _
             " NO_CNPT, " & vbCrLf & _
             " TP_GERA_MOVI, " & vbCrLf & _
             " TX_MOTI_GERA_MOVI, " & vbCrLf & _
             " CO_ITEM_CAIX_NIVE_01, " & vbCrLf & _
             " CO_ITEM_CAIX_NIVE_02, " & vbCrLf & _
             " CO_ITEM_CAIX_NIVE_03, " & vbCrLf & _
             " CO_ITEM_CAIX_NIVE_04, " & vbCrLf & _
             " CO_ITEM_CAIX_NIVE_05, " & vbCrLf & _
             " CO_USUA_ULTI_ATLZ, " & vbCrLf & _
             " CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
             " DH_ULTI_ATLZ, " & vbCrLf
        strSQL = strSQL & " CO_PROD " & vbCrLf & _
            " FROM  A6.TB_MOVI_FUTU_OPER_ATIV "

    ObterPropriedades = fgPropriedades("Grupo_MovimentoFuturo", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela atualização geral da tabela de movimento futuro,
' e pelo retorno à camada intermediária.

Public Function Salvar(ByVal pxmlMoviFutu As MSXML2.DOMDocument40) As Boolean
    
    On Error GoTo ErrorHandler

    flIncluir pxmlMoviFutu

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é retorna se uma determinada data, é dia útil ou não.

Private Function flDiaUtil(ByVal pstrData As String) As Boolean

Dim strSQL                                  As String
Dim rsData                                  As ADODB.Recordset
Dim datData                                 As Date

On Error GoTo ErrorHandler

    datData = fgDtXML_To_Date(pstrData)
    
    If Weekday(datData) = vbSunday Or Weekday(datData) = vbSaturday Then
        flDiaUtil = False
        Exit Function
    End If

    strSQL = " SELECT DT_FERI FROM A8.TB_FERIADO_HO WHERE DT_FERI = " & fgDtXML_To_Oracle(pstrData)

    Set rsData = fgQuerySQL(strSQL)

    flDiaUtil = rsData.RecordCount < 1

    Set rsData = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set rsData = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flDiaUtil Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela inclusão de um lançamento no movimento futuro.

Private Function flIncluir(ByVal pxmlMoviFutu As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim strItemCaix                             As String
Dim strDataRemessa                          As String
Dim strDescricaoAtivo                       As String
Dim strCoRemeMoviFutu                       As String
Dim strDataMovimento                        As String
Dim strCNPJContraparte                      As String
Dim strNomeContraparte                      As String
Dim strAtivo                                As String
Dim strTipoLiqu                             As String

    On Error GoTo ErrorHandler

'    fgGravaArquivo "LOGA6_clsMovimentoFuturo_flIncluir_", vbNullString
    
    strSQL = " INSERT INTO A6.TB_MOVI_FUTU_OPER_ATIV ( " & vbNewLine & _
             " SG_SIST,                                " & vbNewLine & _
             " CO_REME_MOVI_FUTU,                      " & vbNewLine & _
             " DH_REME,                                " & vbNewLine & _
             " CO_EMPR,                                " & vbNewLine & _
             " DT_FECH_PROC,                           " & vbNewLine & _
             " CO_LOCA_LIQU,                           " & vbNewLine & _
             " DE_TIPO_LIQU,                           " & vbNewLine & _
             " DT_LIQU_OPER,                           " & vbNewLine & _
             " IN_MOVI_ENTR_SAID,                      " & vbNewLine & _
             " VA_LIQU_OPER,                           " & vbNewLine & _
             " DE_ATIV,                                " & vbNewLine & _
             " CO_VEIC_LEGA,                           " & vbNewLine & _
             " CO_ITEM_CAIX,                           " & vbNewLine & _
             " CO_CNPJ_CNPT,                           " & vbNewLine & _
             " NO_CNPT,                                " & vbNewLine & _
             " TP_GERA_MOVI,                           " & vbNewLine & _
             " TX_MOTI_GERA_MOVI,                      " & vbNewLine & _
             " CO_ITEM_CAIX_NIVE_01, CO_ITEM_CAIX_NIVE_02, CO_ITEM_CAIX_NIVE_03, CO_ITEM_CAIX_NIVE_04, CO_ITEM_CAIX_NIVE_05, " & vbNewLine & _
             " CO_USUA_ULTI_ATLZ,                      " & vbNewLine & _
             " CO_ETCA_TRAB_ULTI_ATLZ,                 " & vbNewLine & _
             " DH_ULTI_ATLZ,                           " & vbNewLine & _
             " CO_PROD                                 " & vbNewLine & _
             " ) VALUES (                              " & vbNewLine

    With pxmlMoviFutu.documentElement

        If .selectSingleNode("TP_MESG") Is Nothing Then
            Call fgAppendNode(pxmlMoviFutu, "Grupo_MovimentoFuturo", "TP_MESG", enumEvento.DetalheOperacaoCaixaFuturo)
        End If

        If Val(.selectSingleNode("TP_MESG").Text) = enumEvento.MovimentoSubReserva Then

            'Segurar o registro para não ter inclusão simultânea
            'flBloquearItemCaixaParaInclusao .selectSingleNode("CO_ITEM_CAIX").Text
            'Carlos 12/08/2004 - foi alterado, pois qdo a operacao de compromissada vem com mesmo identificador
            ' e com isso duplicava a chave
            flBloquearItemCaixaParaInclusao ""
            strDataRemessa = flObterDataGravacao(fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora)), _
                                                 .selectSingleNode("SG_SIST_ORIG").Text, _
                                                 .selectSingleNode("CO_OPER_ATIV").Text)

            strItemCaix = .selectSingleNode("CO_ITEM_CAIX").Text

            If .selectSingleNode("DT_OPER_ATIV") Is Nothing Then
                If .selectSingleNode("DT_LIQU_OPER_ATIV") Is Nothing Then
                    strDataMovimento = .selectSingleNode("DT_MESG").Text
                Else
                    strDataMovimento = .selectSingleNode("DT_LIQU_OPER_ATIV").Text
                End If
            Else
                strDataMovimento = .selectSingleNode("DT_OPER_ATIV").Text
            End If

            If .selectSingleNode("CO_CNPJ_CNPT") Is Nothing Then
                strCNPJContraparte = vbNullString
            Else
                strCNPJContraparte = .selectSingleNode("CO_CNPJ_CNPT").Text
            End If
            If .selectSingleNode("NO_CNPT") Is Nothing Then
                strNomeContraparte = vbNullString
            Else
                strNomeContraparte = .selectSingleNode("NO_CNPT").Text
            End If

            If .selectSingleNode("DE_TIPO_LIQU").Text = vbNullString Then
                strTipoLiqu = "NULL"
            Else
                strTipoLiqu = "'" & .selectSingleNode("DE_TIPO_LIQU").Text & "'"
            End If
            
            If .selectSingleNode("CO_LOCA_LIQU").Text <> enumLocalLiquidacao.CETIP Then
                If .selectSingleNode("DE_ATIV_MERC") Is Nothing Then
                    strAtivo = "NULL"
                Else
                    strAtivo = "'" & .selectSingleNode("DE_ATIV_MERC").Text & "'"
                End If
            
            Else
                If .selectSingleNode("CO_SUB_TIPO_ATIV") Is Nothing Then
                    strAtivo = "NULL"
                Else
                    strAtivo = "'" & .selectSingleNode("CO_SUB_TIPO_ATIV").Text & "'"
                End If
            
            End If
            
            strSQL = strSQL & _
                     " '" & .selectSingleNode("SG_SIST_ORIG").Text & "' ,            " & vbNewLine & _
                     " '" & .selectSingleNode("CO_OPER_ATIV").Text & "' ,            " & vbNewLine & _
                     fgDtHrXML_To_Oracle(strDataRemessa) & " ,                       " & vbNewLine & _
                     .selectSingleNode("CO_EMPR").Text & " ,                         " & vbNewLine & _
                     fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))) & " , " & vbNewLine & _
                     .selectSingleNode("CO_LOCA_LIQU").Text & " ,                    " & vbNewLine & _
                     strTipoLiqu & ",                                                " & vbNewLine & _
                     fgDtXML_To_Oracle(strDataMovimento) & ",                        " & vbNewLine & _
                     .selectSingleNode("IN_OPER_DEBT_CRED").Text & " ,               " & vbNewLine & _
                     fgVlrToDBServer(.selectSingleNode("//VA_OPER_ATIV").Text) & " ,   " & vbNewLine & _
                     strAtivo & ",                                                   " & vbNewLine & _
                     " '" & .selectSingleNode("//CO_VEIC_LEGA").Text & "' ,            " & vbNewLine & _
                     " '" & strItemCaix & "' ,                                       " & vbNewLine & _
                     IIf(strCNPJContraparte = vbNullString, " NULL ", strCNPJContraparte) & " , " & vbNewLine & _
                     IIf(strNomeContraparte = vbNullString, " NULL ", " '" & strNomeContraparte & "' ") & " , " & vbNewLine & _
                     .selectSingleNode("TP_GERA_MOVI").Text & " ,                 " & vbNewLine & _
                     IIf(.selectSingleNode("TX_MOTI_GERA_MOVI").Text = vbNullString, " NULL ", " '" & .selectSingleNode("TX_MOTI_GERA_MOVI").Text & "' ") & " , " & vbNewLine & _
                     " '" & Mid$(strItemCaix, 2, 3) & "' , '" & Mid$(strItemCaix, 5, 3) & "' , '" & Mid$(strItemCaix, 8, 3) & "' , '" & Mid$(strItemCaix, 11, 3) & "' , '" & Mid$(strItemCaix, 14, 3) & "' , " & vbNewLine & _
                     " '" & fgUsuarioRede & "' ,                                     " & vbNewLine & _
                     " '" & fgEstacaoTrabalhoUsuario & "' ,                          " & vbNewLine & _
                     " SYSDATE ,                                                     " & vbNewLine & _
                     IIf(.selectSingleNode("CO_PROD").Text = vbNullString, " NULL ", " '" & .selectSingleNode("CO_PROD").Text & "' ") & "                  )"
            
        Else
            strItemCaix = .selectSingleNode("CO_ITEM_CAIX").Text
            strCoRemeMoviFutu = .selectSingleNode("CO_REME_MOVI_FUTU").Text
            
            If strCoRemeMoviFutu = vbNullString Then
                strCoRemeMoviFutu = flObterCodigoRemessa
            End If
            
            If Not .selectSingleNode("DT_REME") Is Nothing Then
                strDataRemessa = .selectSingleNode("DT_REME").Text
            Else
                strDataRemessa = .selectSingleNode("DH_REME").Text
            End If
            
            If Not .selectSingleNode("DE_ATIV_MERC") Is Nothing Then
                strDescricaoAtivo = .selectSingleNode("DE_ATIV_MERC").Text
            Else
                strDescricaoAtivo = .selectSingleNode("DE_ATIV").Text
            End If
            
            If Not flDiaUtil(.selectSingleNode("DT_LIQU_OPER").Text) Then
                'Data do movimento não é um dia útil
                lngCodigoErroNegocio = 2027
                GoTo ErrorHandler
            End If
            
            strSQL = strSQL & _
                     " '" & .selectSingleNode("SG_SIST").Text & "' ,                 " & vbNewLine & _
                     " '" & strCoRemeMoviFutu & "' ,                                 " & vbNewLine & _
                     fgDtXML_To_Oracle(strDataRemessa) & " ,                         " & vbNewLine & _
                     .selectSingleNode("CO_EMPR").Text & " ,                         " & vbNewLine & _
                     fgDtXML_To_Oracle(.selectSingleNode("DT_FECH_PROC").Text) & " , " & vbNewLine & _
                     .selectSingleNode("CO_LOCA_LIQU").Text & " ,                    " & vbNewLine & _
                     IIf(.selectSingleNode("DE_TIPO_LIQU").Text = vbNullString, " NULL ", " '" & .selectSingleNode("DE_TIPO_LIQU").Text & "' ") & " , " & vbNewLine & _
                     fgDtXML_To_Oracle(.selectSingleNode("DT_LIQU_OPER").Text) & ",  " & vbNewLine & _
                     .selectSingleNode("IN_MOVI_ENTR_SAID").Text & " ,               " & vbNewLine & _
                     fgVlrToDBServer(.selectSingleNode("VA_LIQU_OPER").Text) & " ,   " & vbNewLine & _
                     " '" & strDescricaoAtivo & "' ,                                 " & vbNewLine & _
                     " '" & .selectSingleNode("CO_VEIC_LEGA").Text & "' ,            " & vbNewLine & _
                     " '" & .selectSingleNode("CO_ITEM_CAIX").Text & "' ,            " & vbNewLine & _
                     IIf(.selectSingleNode("CO_CNPJ_CNPT").Text = vbNullString, " NULL ", .selectSingleNode("CO_CNPJ_CNPT").Text) & " , " & vbNewLine & _
                     IIf(.selectSingleNode("NO_CNPT").Text = vbNullString, " NULL ", " '" & .selectSingleNode("NO_CNPT").Text & "' ") & " , " & vbNewLine & _
                     .selectSingleNode("TP_GERA_MOVI").Text & " ,                 " & vbNewLine & _
                     IIf(.selectSingleNode("TX_MOTI_GERA_MOVI").Text = vbNullString, " NULL ", " '" & .selectSingleNode("TX_MOTI_GERA_MOVI").Text & "' ") & " , " & vbNewLine & _
                     " '" & Mid$(strItemCaix, 2, 3) & "' , '" & Mid$(strItemCaix, 5, 3) & "' , '" & Mid$(strItemCaix, 8, 3) & "' , '" & Mid$(strItemCaix, 11, 3) & "' , '" & Mid$(strItemCaix, 14, 3) & "' , " & vbNewLine & _
                     " '" & fgUsuarioRede & "' ,                                     " & vbNewLine & _
                     " '" & fgEstacaoTrabalhoUsuario & "' ,                          " & vbNewLine & _
                     " SYSDATE ,                                                     " & vbNewLine & _
                     IIf(.selectSingleNode("CO_PROD").Text = vbNullString, " NULL ", " '" & .selectSingleNode("CO_PROD").Text & "' ") & "                  )"
                 
        End If
        
    End With

    fgExecuteSQL strSQL

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela atualização geral da tabela de movimento futuro,
' e pelo retorno à camada intermediária.

Public Function Executar(ByRef pxmlMoviFutu As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlMoviFutu.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Excluir"
            Executar = Salvar(pxmlMoviFutu)
        Case Else
            ' 7 - Operação Inválida
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

' Este método é responsável pela alteração de um lançamento no movimento futuro.

Private Function flAlterar(ByVal pxmlMoviFutu As MSXML2.DOMDocument40) As Boolean
Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = "UPDATE A6.TB_MOVIMENTO_FUTURO SET " & vbCrLf & _
            "   DT_MOVI =  " & fgDtXML_To_Oracle(pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/DT_MOVI").Text) & ", " & vbCrLf & _
            "   CO_VEIC_LEGA = '" & pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/CO_VEIC_LEGA").Text & "', " & vbCrLf & _
            "   CO_BANC_LIQU = " & CLng(pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/CO_BANC_LIQU").Text) & ", " & vbCrLf & _
            "   CO_ITEM_CAIX = '" & pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/CO_ITEM_CAIX").Text & "', " & vbCrLf & _
            "   HO_REME = TO_DATE('" & pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/HO_REME").Text & "','HH24MISS'), " & vbCrLf & _
            "   IN_ENTR_SAID = " & CLng(pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/IN_ENTR_SAID").Text) & ", " & vbCrLf & _
            "   VA_MOVI = " & fgVlrToDBServer(pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/VA_MOVI").Text) & ", " & vbCrLf & _
            "   DE_TITL = '" & pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/DE_TITL").Text & "', " & vbCrLf & _
            "   CNPJ_COTR_PART = " & CLng(pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/CNPJ_COTR_PART").Text) & ", " & vbCrLf & _
            "   DE_COTR_PART = '" & pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/DE_COTR_PART").Text & "', " & vbCrLf & _
            "   IN_TP_MOVI = " & CLng(pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/IN_TP_MOVI").Text) & ", " & vbCrLf & _
            "   MO_AJU = '" & pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/MO_AJU").Text & "', " & vbCrLf & _
            "   CO_USUA_ULTI_ATLZ = '" & pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/CO_USUA_ULTI_ATLZ").Text & "', " & vbCrLf & _
            "   DH_ULTI_ATLZ = SYSDATE" & vbCrLf & _
            "WHERE " & vbCrLf & _
            "   SQ_MOVI_FUTU = " & CLng(pxmlMoviFutu.selectSingleNode("//Grupo_MovimentoFuturo/SQ_MOVI_FUTU").Text) & " " & vbCrLf

    fgExecuteSQL strSQL

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela exclusão de um lançamento do movimento futuro.

Public Function Excluir(ByVal pxmlMoviFutu As MSXML2.DOMDocument40) As Boolean
Dim strSQL                                   As String

    On Error GoTo ErrorHandler

    With pxmlMoviFutu.documentElement
        strSQL = " DELETE FROM " & vbCrLf & _
                 "        A6.TB_MOVI_FUTU_OPER_ATIV " & vbCrLf & _
                 " WHERE  CO_REME_MOVI_FUTU   BETWEEN " & vbCrLf & _
                 "                       '" & .selectSingleNode("CO_REME_PRMR").Text & "' AND " & vbCrLf & _
                 "                       '" & .selectSingleNode("CO_REME_ULTI").Text & "' " & vbCrLf & _
                 " AND    SG_SIST      = '" & .selectSingleNode("SG_SIST_ORIG").Text & "' " & vbCrLf & _
                 " AND    CO_VEIC_LEGA = '" & .selectSingleNode("CO_VEIC_LEGA").Text & "' " & vbCrLf & _
                 " AND    CO_EMPR      =  " & .selectSingleNode("CO_EMPR").Text & vbCrLf & _
                 " AND    DT_FECH_PROC <  " & fgDtXML_To_Oracle(.selectSingleNode("DT_FECH_PROC").Text)
    End With

    fgExecuteSQL strSQL

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExcluirGrupo Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Este método é responsável pela obtenção de um código sequencial de remessa para o movimento futuro.

Private Function flObterCodigoRemessa() As String

Dim lngRetorno                              As Long

On Error GoTo ErrHandler

    lngRetorno = fgExecuteSequence("A8.SQ_A8_NU_SEQU_REME_PJ")
    flObterCodigoRemessa = fgDt_To_Xml(fgDataHoraServidor(DataHoraAux)) & _
                           "A6 " & _
                           fgCompletaString(lngRetorno, "0", 9, True)

Exit Function

ErrHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterCodigoRemessa", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pelo bloqueio de um registro da tabela de item caixa, para a inclusão de um movimento futuro.

Private Sub flBloquearItemCaixaParaInclusao(ByVal pstrItemCaixa As String)

Dim rsItemCaixa                              As ADODB.Recordset
Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT CO_ITEM_CAIX               " & vbNewLine & _
             " FROM   A6.TB_ITEM_CAIX_OPER_ATIV  " & vbNewLine & _
             " FOR UPDATE "

    '" WHERE  CO_ITEM_CAIX      =       '" & pstrItemCaixa & "' " & _

    Set rsItemCaixa = fgQuerySQL(strSQL)

    Exit Sub

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flBloquearItemCaixaParaInclusao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Sub

' Este método é responsável pela obtenção de uma data de inclusão, para que não seja incluído um movimento com chave
' primária repetida.

Private Function flObterDataGravacao(ByVal pstrDataRemessa As String, _
                                     ByVal pstrSiglaSistema As String, _
                                     ByVal pstrCodigoRemessa As String) As String

Dim strSQL                                  As String
Dim objRs                                   As ADODB.Recordset
Dim datDataGravacao                         As Date

    On Error GoTo ErrorHandler

    strSQL = " SELECT   MAX(DH_REME) AS DH_REME     " & vbNewLine & _
             " FROM     A6.TB_MOVI_FUTU_OPER_ATIV   " & vbNewLine & _
             " WHERE    DH_REME             =    " & fgDtHrXML_To_Oracle(pstrDataRemessa) & vbNewLine & _
             " AND      SG_SIST             =   '" & pstrSiglaSistema & "'" & vbNewLine & _
             " AND      CO_REME_MOVI_FUTU   =   '" & pstrCodigoRemessa & "'"

    Set objRs = fgQuerySQL(strSQL)

    If IsNull(objRs!DH_REME) Then
        datDataGravacao = fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux)
    Else
        If DateDiff("s", fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux), objRs!DH_REME) = 0 Then
            datDataGravacao = DateAdd("s", 1, fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux))
        Else
            datDataGravacao = fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux)
        End If
    End If

    flObterDataGravacao = fgDtHr_To_Xml(datDataGravacao)

    objRs.Close
    Set objRs = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objRs = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterDataGravacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
