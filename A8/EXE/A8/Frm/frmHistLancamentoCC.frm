VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHistLancamentoCC 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalhe de Lançamentos de Conta-Corrente"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wbDetalhe 
      Height          =   5535
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   8655
      ExtentX         =   15266
      ExtentY         =   9763
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin TabDlg.SSTab sstDetalheLancamentoCC 
      Height          =   8415
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   14843
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Histórico de Status"
      TabPicture(0)   =   "frmHistLancamentoCC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvwHistoricoStatus"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Composição de Net"
      TabPicture(1)   =   "frmHistLancamentoCC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwComposicaoNet"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView lvwComposicaoNet 
         Height          =   2295
         Left            =   -74910
         TabIndex        =   2
         Top             =   390
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwHistoricoStatus 
         Height          =   2295
         Left            =   90
         TabIndex        =   3
         Top             =   390
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Data / Hora"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Situação"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Usuário"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Motivo"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Código Erro"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Descrição Erro"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   30
      Top             =   8370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistLancamentoCC.frx":0038
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistLancamentoCC.frx":014A
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistLancamentoCC.frx":0464
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistLancamentoCC.frx":07B6
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistLancamentoCC.frx":08C8
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistLancamentoCC.frx":0BE2
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistLancamentoCC.frx":0EFC
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistLancamentoCC.frx":1216
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistLancamentoCC.frx":1530
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistLancamentoCC.frx":1882
            Key             =   "amarelo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistLancamentoCC.frx":1BD4
            Key             =   "laranja"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistLancamentoCC.frx":1F26
            Key             =   "vermelho"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandosForm 
      Height          =   330
      Left            =   8070
      TabIndex        =   0
      Top             =   8490
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   582
      ButtonWidth     =   1482
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "OK"
            Object.ToolTipText     =   "Fechar formulário"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHistLancamentoCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Objeto responsável pela consulta do histórico de Lançamento de Conta Corrente, através de
' interação com a camada de controle de caso de uso MIU.
Option Explicit

Public vntSequenciaOperacao                 As Variant
Public lngCodigoEmpresa                     As Long
Public lngTipoLancamentoITGR                As Long
Public intSequenciaLancamento               As Integer
Public strNetOperacoes                      As String

Private Const COL_DATA_HOTA                 As Integer = 0
Private Const COL_SITUACAO                  As Integer = 1
Private Const COL_USUARIO                   As Integer = 2
Private Const COL_MOTIVO                    As Integer = 3
Private Const COL_COD_ERRO                  As Integer = 4
Private Const COL_TX_ERRO                   As Integer = 5

Private Const strFuncionalidade             As String = "frmHistLancamentoCC"

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40

Private strConteudoBrowser                  As String

'Carregar a mensagem HTML com o detalhe do lançamento do conta corrente para exibição no browser
Private Sub flCarregaMensagemHTML(ByVal pvntSequencial As Variant)

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim strMensagemHTML                         As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Call fgCursor(True)
    Call flAtualizaConteudoBrowser

    '>>> Formata XML Filtro padrão... --------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Sequencial", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Sequencial", _
                                     "Sequencial", pvntSequencial)
    
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Empresa", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Empresa", _
                                     "Empresa", lngCodigoEmpresa)
    
    '>>> -------------------------------------------------------------------------------------------
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    strMensagemHTML = objMensagem.ObterMensagemHTMLPorOperacao(xmlDomFiltros.xml, _
                                                               vntCodErro, _
                                                               vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Call flFormataDatas(strMensagemHTML)
    
    'Força o Browser a atualizar a página com o conteúdo obtido
    Call flAtualizaConteudoBrowser(strMensagemHTML)
     
    Set objMensagem = Nothing
    Set xmlDomFiltros = Nothing
    
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    fgCursor
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaMensagemHTML", 0

End Sub

'Formatar as datas para apresentação no browser
Private Sub flFormataDatas(ByRef strMensagemHTML As String)

Dim lngPosicao                              As Long
Dim strDataRawFormat                        As String
Dim strSaida                                As String
Dim datDataFormatada                        As Date

Const TAG_DATA_HORA                         As String = "|DH|"
Const TAG_DATA_HORA_SIZE                    As Long = 19

Const TAG_DATA                              As String = "|DT|"
Const TAG_DATA_SIZE                         As Long = 13

    Do
        lngPosicao = InStr(strMensagemHTML, TAG_DATA)
        If lngPosicao > 0 Then
            strSaida = Mid$(strMensagemHTML, 1, lngPosicao - 1)
            strDataRawFormat = Mid$(strMensagemHTML, lngPosicao, TAG_DATA_SIZE)
            datDataFormatada = fgDtXML_To_Date(Mid$(strDataRawFormat, Len(TAG_DATA) + 1, 8))
            strSaida = strSaida & datDataFormatada & Mid$(strMensagemHTML, lngPosicao + TAG_DATA_SIZE)
            strMensagemHTML = strSaida
        End If
    Loop While lngPosicao <> 0

     Do
        lngPosicao = InStr(strMensagemHTML, TAG_DATA_HORA)
        If lngPosicao > 0 Then
            strSaida = Mid$(strMensagemHTML, 1, lngPosicao - 1)
            strDataRawFormat = Mid$(strMensagemHTML, lngPosicao, TAG_DATA_HORA_SIZE)
            datDataFormatada = fgDtHrStr_To_DateTime(Mid$(strDataRawFormat, Len(TAG_DATA_HORA) + 1, 14))
            strSaida = strSaida & datDataFormatada & Mid$(strMensagemHTML, lngPosicao + TAG_DATA_HORA_SIZE)
            strMensagemHTML = strSaida
        End If
    Loop While lngPosicao <> 0

End Sub

'Atualiza o conteúdo do browser
Private Sub flAtualizaConteudoBrowser(Optional pstrConteudo As String = vbNullString)

On Error GoTo ErrorHandler

    strConteudoBrowser = pstrConteudo
    wbDetalhe.Navigate "about:blank"

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flAtualizaConteudoBrowser", 0
End Sub

Private Sub lvwComposicaoNet_ItemClick(ByVal Item As MSComctlLib.ListItem)

    On Error GoTo ErrorHandler

    Call flCarregaMensagemHTML(Split(Item.Key, "|")(1))
    
    Exit Sub

ErrorHandler:
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)

End Sub

Private Sub sstDetalheLancamentoCC_Click(PreviousTab As Integer)

    On Error GoTo ErrorHandler

    If sstDetalheLancamentoCC.Tab = 0 Then
        If lvwHistoricoStatus.ListItems.Count > 0 Then
            lvwHistoricoStatus.ListItems(1).Selected = True
            Call flCarregaMensagemHTML(vntSequenciaOperacao)
        End If
    Else
        If lvwComposicaoNet.ListItems.Count > 0 Then
            lvwComposicaoNet.ListItems(1).Selected = True
            Call flCarregaMensagemHTML(Split(lvwComposicaoNet.SelectedItem.Key, "|")(1))
        End If
    End If
    
    Exit Sub

ErrorHandler:
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)

End Sub

Private Sub wbDetalhe_DocumentComplete(ByVal pDisp As Object, URL As Variant)

'Ocorre um erro, porém a sentença é executada com sucesso
    On Error Resume Next
    
    If strConteudoBrowser <> "" Then
        pDisp.Document.Body.innerHTML = strConteudoBrowser
    End If
End Sub

Private Sub Form_Load()
    
On Error GoTo ErrorHandler

    Call fgCursor(True)

    Call fgCenterMe(Me)
    Set Me.Icon = mdiLQS.Icon
    
    DoEvents
    
    fgCursor False
    
    sstDetalheLancamentoCC.TabEnabled(1) = False
    
    Call flInicializar
    Call flPreencherHistorico
    Call flCarregaMensagemHTML(vntSequenciaOperacao)
    Call flVerificarConsolidacaoLancamento
    
    Exit Sub

ErrorHandler:
    Call fgCursor(False)
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)

End Sub

' Obtém as propriedades inicias da tela, através de interação com a camada de
' controle de caso de uso, método A8MIU.clsMiu.ObterMapaNavegacao
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim strMapaNavegacao        As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    Set xmlMapaNavegacao = Nothing
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    strMapaNavegacao = objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, _
                                                 strFuncionalidade, _
                                                 vntCodErro, _
                                                 vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmHistLancamentoCC", "flInicializar")
    End If
    
Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0
    
End Sub

' Obtém as mensagens pertinenntes ao filtro e as exibe no listview do histótico.
' Utiliza a camada de controle de caso de uso, método A8MIU.clsMIU.Executar
Private Sub flPreencherHistorico()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim strLeitura                              As String

Dim objListItem                             As MSComctlLib.ListItem
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    With xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_HistSituacaoIntegracao")
        .selectSingleNode("@Operacao").Text = gstrOperLerTodos
        .selectSingleNode("NU_SEQU_OPER_ATIV").Text = vntSequenciaOperacao
        .selectSingleNode("TP_LANC_ITGR").Text = lngTipoLancamentoITGR
        .selectSingleNode("NR_SEQU_LANC").Text = intSequenciaLancamento
        
        Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
        strLeitura = objMIU.Executar(.xml, vntCodErro, vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objMIU = Nothing
        
    End With
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    If xmlLerTodos.loadXML(strLeitura) Then
        lvwHistoricoStatus.ListItems.Clear
        For Each objDomNode In xmlLerTodos.documentElement.childNodes
            Set objListItem = lvwHistoricoStatus.ListItems.Add
            objListItem.Text = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_SITU_LANC_CC").Text)
            objListItem.SubItems(COL_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
            objListItem.SubItems(COL_USUARIO) = objDomNode.selectSingleNode("CO_USUA_ATLZ").Text
            objListItem.SubItems(COL_MOTIVO) = objDomNode.selectSingleNode("TX_JUST_CANC").Text
            objListItem.SubItems(COL_COD_ERRO) = objDomNode.selectSingleNode("CO_ERRO").Text
            objListItem.SubItems(COL_TX_ERRO) = objDomNode.selectSingleNode("TX_MESG_ERRO").Text
        Next objDomNode
    End If

    Set xmlLerTodos = Nothing

Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlLerTodos = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flPreencherHistorico", 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlMapaNavegacao = Nothing
    Set frmHistLancamentoCC = Nothing
End Sub

Private Sub tlbComandosForm_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    Unload Me

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbComandosForm_ButtonClick"
End Sub

Private Sub flVerificarConsolidacaoLancamento()

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsOperacao
#End If

Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim objXMLNode                              As MSXML2.IXMLDOMNode

Dim intCondicaoNet                          As Integer

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
Dim vntPasso                                As Variant
    
    On Error GoTo ErrorHandler
    
    vntPasso = 1
    If UBound(Split(strNetOperacoes, "|")) < 2 Then Exit Sub
    
    vntPasso = 2
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    vntPasso = 3
    Call fgAppendNode(xmlLeitura, "", "Repeat_Filtros", "")
    Call fgAppendNode(xmlLeitura, "Repeat_Filtros", "Grupo_NumeroOperacao", vbNullString)
    
    vntPasso = 4
    For intCondicaoNet = 1 To UBound(Split(strNetOperacoes, "|"))
        Call fgAppendNode(xmlLeitura, "Grupo_NumeroOperacao", "NumeroOperacao", Split(strNetOperacoes, "|")(intCondicaoNet))
    Next
    
    vntPasso = 5
    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    
    vntPasso = 6
    Call xmlLeitura.loadXML(objOperacao.ObterDetalheOperacao(xmlLeitura.xml, vntCodErro, vntMensagemErro))
    
    vntPasso = 7
    sstDetalheLancamentoCC.TabEnabled(1) = True
    Call flInicializarLvwComposicaoNet
    
    vntPasso = 8
    For Each objXMLNode In xmlLeitura.selectNodes("//Repeat_DetalheOperacao/*")
        With lvwComposicaoNet.ListItems.Add(, "|" & objXMLNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
            .Text = objXMLNode.selectSingleNode("NO_VEIC_LEGA").Text
            .SubItems(1) = objXMLNode.selectSingleNode("IN_ENTR_SAID_RECU_FINC").Text
            .SubItems(2) = fgVlrXml_To_Interface(objXMLNode.selectSingleNode("VA_OPER_ATIV").Text)
            .SubItems(3) = fgDtXML_To_Interface(objXMLNode.selectSingleNode("DT_OPER_ATIV").Text)
            .SubItems(4) = objXMLNode.selectSingleNode("SG_LOCA_LIQU").Text
        End With
    Next
    
    Set xmlLeitura = Nothing
   
    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flVerificarConsolidacaoLancamento", 0, , "Erro no passo " & vntPasso
    
End Sub

Private Sub flInicializarLvwComposicaoNet()

    On Error GoTo ErrorHandler
    
    With Me.lvwComposicaoNet.ColumnHeaders
        .Add , , "Veículo Legal", 3600
        .Add , , "D/C", 840
        .Add , , "Valor", , lvwColumnRight
        .Add , , "Data", 1070
        .Add , , "Câmara", 1250
    End With
    
    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarLvwComposicaoNet", 0
    
End Sub
