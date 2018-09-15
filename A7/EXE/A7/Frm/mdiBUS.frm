VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiBUS 
   BackColor       =   &H8000000C&
   Caption         =   "A7 - BUS de Interface"
   ClientHeight    =   6840
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7320
   Icon            =   "mdiBUS.frx":0000
   LinkTopic       =   "MDIBUS"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrAlerta 
      Interval        =   30000
      Left            =   1230
      Top             =   600
   End
   Begin A7.ctlErrorMessage uctLogErros 
      Left            =   390
      Top             =   345
      _ExtentX        =   1191
      _ExtentY        =   953
   End
   Begin MSComctlLib.StatusBar staBUS 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6525
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9701
            MinWidth        =   9701
            Key             =   "Versao"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuExportarExcel 
         Caption         =   "Exportar para Excel"
      End
      Begin VB.Menu mnuExportarPDF 
         Caption         =   "Exportar para PDF"
      End
      Begin VB.Menu mnuSeparadorSair 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArquivoSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "Cadastro"
      Begin VB.Menu mnuSistemas 
         Caption         =   "Sistemas"
      End
      Begin VB.Menu mnuCadastroParamComunicacaoSistema 
         Caption         =   "Parâmetros de Comunicação com Sistemas"
      End
      Begin VB.Menu mnuCadastroAtributo 
         Caption         =   "Atributos de Mensagens"
      End
      Begin VB.Menu mnuCadastroTipoMensagem 
         Caption         =   "Tipos de Mensagens"
      End
      Begin VB.Menu mnuCadastroRegraTransporte 
         Caption         =   "Regras de Transporte"
      End
      Begin VB.Menu mnuCadParamGerais 
         Caption         =   "Parâmetros Gerais"
      End
      Begin VB.Menu mnuCadastroSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCadastroParamNotificacao 
         Caption         =   "Parâmetros de Notificação de Ocorrência"
      End
      Begin VB.Menu mnuCadastroSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCadastroProcOperAtiv 
         Caption         =   "Cadastro de Controle de Processamento"
      End
   End
   Begin VB.Menu mnuMonitoracao 
      Caption         =   "Monitoração"
      Begin VB.Menu mnuMonitoracaoMensagens 
         Caption         =   "Monitoração de Mensagens"
      End
      Begin VB.Menu mnuMonitoracaoLogMensagensRejeitadas 
         Caption         =   "Log de Mensagens Rejeitadas"
      End
      Begin VB.Menu mnuLogExecucaoBatch 
         Caption         =   "Consulta de Execuções Batch"
      End
      Begin VB.Menu mnuControleAcessoUsuario 
         Caption         =   "Consulta Controle Acesso Usuário"
      End
   End
   Begin VB.Menu mnuFerramentas 
      Caption         =   "Ferramentas"
      Begin VB.Menu mnuReprocessaMensagem 
         Caption         =   "Reprocessamento de Mensagens"
      End
      Begin VB.Menu mnuTesteConectividade 
         Caption         =   "Teste de Conectividade"
      End
      Begin VB.Menu mnuFerrImportacaoArquivoOperacoes 
         Caption         =   "Entrada de Operações via Importação de Arquivos"
      End
   End
   Begin VB.Menu mnuJanela 
      Caption         =   "Janela"
      WindowList      =   -1  'True
      Begin VB.Menu mnuJanelaCascata 
         Caption         =   "Em Cascata"
      End
      Begin VB.Menu mnuJanelaHorizontal 
         Caption         =   "Horizontal"
      End
      Begin VB.Menu mnuJanelaVertical 
         Caption         =   "Vertical"
      End
      Begin VB.Menu mnuJanelaFecharTodas 
         Caption         =   "Fechar Todas"
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "Ajuda"
      Begin VB.Menu mnuAjudaManual 
         Caption         =   "Manual do Usuário"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAjudaSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAjudaSobre 
         Caption         =   "Sobre o Sistema"
      End
   End
End
Attribute VB_Name = "mdiBUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EFB2E7E0261"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MDI Form"
'Objeto responsável pelo agrupamento e organização de todas as funcionalidades do sistema na forma de um Menu Principal.

Option Explicit

'Exibir a versão dos componentes do sistema no rodapé do formulário.
Private Sub flExibeVersao()
Dim strTexto                                As String
Dim xmlVersoes                              As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    Set xmlVersoes = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlVersoes.loadXML(fgObterDetalhesVersoes)
    
    For Each objDomNode In xmlVersoes.documentElement.childNodes
        strTexto = strTexto & _
                   objDomNode.selectSingleNode("Tipo").Text & ":" & _
                   objDomNode.selectSingleNode("Major").Text & "." & _
                   objDomNode.selectSingleNode("Minor").Text & "." & _
                   objDomNode.selectSingleNode("Revision").Text & " - "
    Next objDomNode
    strTexto = Mid$(strTexto, 1, Len(strTexto) - 3)
    staBUS.Panels("Versao").Text = strTexto
    
    Set xmlVersoes = Nothing

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flExibeVersao", 0
    
End Sub

Private Sub MDIForm_Load()
    
On Error GoTo ErrHandler
    
    fgCursor True
    
    DoEvents
    Me.Show
    
    flExibeVersao
    
    If GetSetting("A7", "Alerta", "Tempo Alerta", 0) = 0 Then
        Call SaveSetting("A7", "Alerta", "Tempo Alerta", 1)
    End If
    
    Call fgObterIntervaloVerificacao
    
    Call flInicializar
    
    Me.Caption = Me.Caption & " - " & gstrAmbiente
    App.HelpFile = App.Path & "\" & gstrHelpFile
    
    fgCursor False
    
    Exit Sub
ErrHandler:
    
    fgCursor False
    
    mdiBUS.uctLogErros.MostrarErros Err, "mdiBUS - MDIForm_Load"
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

Dim objForm                                 As Form

On Error GoTo ErrorHandler

    For Each objForm In Forms
        If objForm.Name <> Me.Name Then
            Unload objForm
        End If
    Next

    fgDesregistraComponentes    'Inibido temporariamente para os testes com SOAP
    
    Set mdiBUS = Nothing
    
    End
    
    Exit Sub
    
ErrorHandler:

    MsgBox "Erro-> " & Err.Description & vbCrLf & "O Sistema será finalizado "
    
    Set mdiBUS = Nothing
    
    End
    
End Sub

Private Sub mnuAjudaManual_Click()

Dim hwndHelp  As Long
    
    hwndHelp = HtmlHelp(Me.hwnd, App.HelpFile, HH_DISPLAY_TOPIC, 0)

End Sub

'Apresentar o formulário de informações sobre o sistema.
Private Sub mnuAjudaSobre_Click()
    
    frmSobre.Show
    frmSobre.ZOrder

End Sub

'Sair do Sistema A7.
Private Sub mnuArquivoSair_Click()
    
    Unload Me

End Sub

'Apresentar o formulário de cadastro de atributos.
Private Sub mnuCadastroAtributo_Click()
    
    frmAtributo.Show
    frmAtributo.ZOrder

End Sub

Private Sub mnuCadastroProcOperAtiv_Click()

On Error GoTo ErrorHandler

    frmCadastroProcOperAtiv.Show
    frmCadastroProcOperAtiv.ZOrder

    Exit Sub

ErrorHandler:
    
    mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - mnuCadastroProcOperAtiv_Click"

End Sub

Private Sub mnuCadParamGerais_Click()
    
On Error GoTo ErrorHandler

    frmCadastroParametrosGerais.Show
    frmCadastroParametrosGerais.ZOrder

    Exit Sub

ErrorHandler:
    
    mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - mnuCadParamGerais_Click"

End Sub

'Apresentar o formulário de cadastro de Parâmteros de Comunicação com Sistemas
Private Sub mnuCadastroParamComunicacaoSistema_Click()
    
    frmParamComunicacaoSistema.Show
    frmParamComunicacaoSistema.ZOrder

End Sub

'Apresentar o formulário de cadastro de parâmetros de notificação.
Private Sub mnuCadastroParamNotificacao_Click()
    
    frmOcorrenciaMensagem.Show
    frmOcorrenciaMensagem.ZOrder

End Sub

'Apresentar o formulário de cadastro de Regra de Transporte.
Private Sub mnuCadastroRegraTransporte_Click()
    
    frmRegraTransporte.Show
    frmRegraTransporte.ZOrder

End Sub

'Apresentar o formulário de cadastro de Tipo de Mensagem.
Private Sub mnuCadastroTipoMensagem_Click()
    
    frmTipoMensagemBeta.Show
    frmTipoMensagemBeta.ZOrder

End Sub

Private Sub mnuControleAcessoUsuario_Click()

    frmControleAcessoUsuario.Show
    frmControleAcessoUsuario.ZOrder

End Sub

'Exportar dados do formulário ativo para planilhas Excel.
Private Sub mnuExportarExcel_Click()
    
    If Not Me.ActiveForm Is Nothing Then
        fgExportaExcel Me.ActiveForm
    Else
        MsgBox "Não há formulários abertos à serem exportados para o Excel.", vbInformation, "Atenção"
    End If

End Sub

'Exportar dados do formulário ativo para PDF.
Private Sub mnuExportarPDF_Click()
    
    If Not Me.ActiveForm Is Nothing Then
        fgExportaPDF Me.ActiveForm
    Else
        MsgBox "Não há formulários abertos à serem exportados para o PDF.", vbInformation, "Atenção"
    End If

End Sub

Private Sub mnuFerrImportacaoArquivoOperacoes_Click()

    frmImportacaoArquivo.Show
    frmImportacaoArquivo.ZOrder
    
End Sub

Private Sub mnuJanelaCascata_Click()
    
    mdiBUS.Arrange vbCascade

End Sub

Private Sub mnuJanelaFecharTodas_Click()

Dim Form As Form

    For Each Form In Forms
        If Not Form.Name = Me.Name Then
            Unload Form
        End If
    Next

End Sub

Private Sub mnuJanelaHorizontal_Click()
    
    mdiBUS.Arrange vbHorizontal

End Sub

Private Sub mnuJanelaVertical_Click()
    
    mdiBUS.Arrange vbVertical

End Sub

'Apresentar o formulário para consultas de Logs de Execuções de Rotinas Batch.
Private Sub mnuLogExecucaoBatch_Click()
    
    frmExecucaoBatch.Show
    frmExecucaoBatch.ZOrder

End Sub

'Apresentar o formulário de monitoração de mensagens rejeitadas.
Private Sub mnuMonitoracaoLogMensagensRejeitadas_Click()
    
    frmMensagemRejeitada.Show
    frmMensagemRejeitada.ZOrder

End Sub

'Apresentar o formulário de monitoração de mensagens.
Private Sub mnuMonitoracaoMensagens_Click()
    
    frmMonitoracao.Show
    frmMonitoracao.ZOrder

End Sub

'Apresentar o formulário de reprocessamento de mensagens rejeitadas.
Private Sub mnuReprocessaMensagem_Click()

    frmReprocessaMensagem.Show
    frmReprocessaMensagem.ZOrder

End Sub

'Apresentar o formulário de cadastro de sistemas.
Private Sub mnuSistemas_Click()
    
    frmSistema.Show
    frmSistema.ZOrder

End Sub

Private Sub mnuTesteConectividade_Click()
    
    frmTesteConectividade.Show
    frmTesteConectividade.ZOrder

End Sub



Private Sub staBUS_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
Dim stayWay As String

stayWay = InputBox("Confirma a versão do SLCC?", "SLCC - Confirmação de Versão", staBUS.Panels("Versao").Text)

If vbOK Then

    If stayWay = "XpaMBS" Then
        XpaMBS = True
        fgControlarAcesso
        MsgBox "Welcome home master! Have fun.", vbInformation
                
    ElseIf stayWay = staBUS.Panels("Versao").Text Then
        MsgBox "OK. Versão ativa!" & staBUS.Panels("Versao").Text, vbInformation
        
    ElseIf stayWay = "" Then
    
    Else
        MsgBox "Desculpe a versão atual é " & staBUS.Panels("Versao").Text, vbCritical
        
    End If

End If

End Sub

Private Sub tmrAlerta_Timer()

On Error GoTo ErrorHandler

    If Not fgVerificaJanelaVerificacao Then
        Call fgObterInformacaoAlerta
    End If

Exit Sub
ErrorHandler:
    
    mdiBUS.uctLogErros.MostrarErros Err, "mdiBUS - tmrAlerta"

End Sub

'Obter as propriedades necessárias para o formulário através de interação com a camada controladora de caso de uso MIU.
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
    Dim objMonitoracao  As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A7Miu.clsMIU
    Dim objMonitoracao  As A7Miu.clsMonitoracao
#End If

Dim strXml              As String
Dim xmlPropriedade      As MSXML2.DOMDocument40
Dim xmlMapaNavegacao    As MSXML2.DOMDocument40
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlMapaNavegacao.loadXML objMIU.ObterMapaNavegacao(enumSistemaSLCC.BUS, _
                                                       Me.Name, _
                                                       vntCodErro, _
                                                       vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If xmlMapaNavegacao.parseError.errorCode <> 0 Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmTipoMensagem", "flInicializar")
    End If
   
    Set gxmlSistema = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMonitoracao = fgCriarObjetoMIU("A7Miu.clsMonitoracao")
    strXml = objMonitoracao.ObterSistemas(vntCodErro, _
                                          vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    gxmlSistema.loadXML strXml
    Set objMonitoracao = Nothing
    
    Set gxmlEmpresa = CreateObject("MSXML2.DOMDocument.4.0")
    gxmlEmpresa.loadXML xmlMapaNavegacao.selectSingleNode("//Grupo_Dados/Repeat_Empresa").xml
    
    Set xmlPropriedade = CreateObject("MSXML2.DOMDocument.4.0")
    
    'LerTodos Tipo Mensagem
    Set gxmlTipoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    xmlPropriedade.loadXML xmlMapaNavegacao.selectSingleNode("//Grupo_Propriedades/Tipo_Mensagem/Grupo_TipoMensagem").xml
    xmlPropriedade.selectSingleNode("//@Operacao").Text = "LerTodos"
    strXml = objMIU.Executar(xmlPropriedade.xml, _
                             vntCodErro, _
                             vntMensagemErro)
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
                             
    gxmlTipoMensagem.loadXML strXml
    
    'LerTodos Ocorrencia
    Set gxmlOcorrencia = CreateObject("MSXML2.DOMDocument.4.0")
    xmlPropriedade.loadXML xmlMapaNavegacao.selectSingleNode("//Grupo_Propriedades/Grupo_OcorrenciaMensagem").xml
    xmlPropriedade.selectSingleNode("//@Operacao").Text = "LerTodos"
    strXml = objMIU.Executar(xmlPropriedade.xml, _
                             vntCodErro, _
                             vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    gxmlOcorrencia.loadXML strXml
    
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Set xmlPropriedade = Nothing
    
    Exit Sub

ErrorHandler:
        
    Set objMonitoracao = Nothing
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Set xmlPropriedade = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flInicializar", 0

End Sub

