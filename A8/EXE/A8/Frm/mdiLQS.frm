VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.MDIForm mdiLQS 
   BackColor       =   &H8000000C&
   Caption         =   "A8 - Sistema de Liquidação e Controle das Câmaras"
   ClientHeight    =   5745
   ClientLeft      =   1935
   ClientTop       =   2670
   ClientWidth     =   10995
   Icon            =   "mdiLQS.frx":0000
   LinkTopic       =   "MDILQS"
   WindowState     =   2  'Maximized
   Begin A8.ctlSysTray ctlSysTray1 
      Left            =   480
      Top             =   3960
      _extentx        =   688
      _extenty        =   688
      intray          =   0
      trayicon        =   "mdiLQS.frx":030A
      traytip         =   "SLCC Control."
   End
   Begin MSComctlLib.ImageList imgLQS 
      Left            =   2700
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiLQS.frx":075E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staLQS 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5430
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Key             =   "BackOffice"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Key             =   "Contingencia"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9701
            MinWidth        =   9701
            Key             =   "Versao"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrIntervalo 
      Interval        =   1000
      Left            =   0
      Top             =   4980
   End
   Begin A8.ctlErrorMessage uctlogErros 
      Left            =   750
      Top             =   1140
      _extentx        =   1191
      _extenty        =   1032
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuExportarExcel 
         Caption         =   "Exportar para Excel"
      End
      Begin VB.Menu mnuExportarPDF 
         Caption         =   "Exportar para PDF"
      End
      Begin VB.Menu mnuFerramentasSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuSegregacaoAcesso 
      Caption         =   "&Segregação Acesso"
      Begin VB.Menu mnuSegrControleAcessoDados 
         Caption         =   "Controle Acesso Dados"
      End
      Begin VB.Menu mnuCadastroSeparador01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSegrGrupoUsuario 
         Caption         =   "Grupo Usuário"
      End
      Begin VB.Menu mnuSegrGrupoUsuarioMsgSPB 
         Caption         =   "Grupo Usuário X Mensagens SPB"
      End
      Begin VB.Menu mnuCadastroSeparador03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSegrGrupoVeiculoLegal 
         Caption         =   "Grupo Veículo Legal"
      End
      Begin VB.Menu mnuCadastroSeparador04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCadMargemSegurancaGradeHorario 
         Caption         =   "Parametrização de Grade de Horário"
      End
      Begin VB.Menu mnuCadHorarioLimiteIntSistLeg 
         Caption         =   "Horário Limite Integração Sistemas Legados"
      End
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnuIdPartCamVeicLega 
         Caption         =   "Identificador Particip. Câmara Veic. Legal"
      End
      Begin VB.Menu mnuCadVeicLegalXGrupoVeicLegal 
         Caption         =   "Veículo Legal X Grupo Veículo Legal"
      End
      Begin VB.Menu mnuCadVeicLegalXEmpresa 
         Caption         =   "Veículo Legal X Empresa"
      End
      Begin VB.Menu mnuCadastroSeparador02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCadParamWorkflow 
         Caption         =   "Parametrização de Workflow"
      End
      Begin VB.Menu mnuCadParamAlerta 
         Caption         =   "Parametrização de Alertas"
      End
      Begin VB.Menu mnuCadastroSeparador05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCadTipoJustificativaConciliacao 
         Caption         =   "Tipo Justificativa Conciliação"
      End
      Begin VB.Menu mnuCadastroSeparador06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParamHistCC 
         Caption         =   "Histórico Conta Corrente"
      End
      Begin VB.Menu mnuParamHistCNTB 
         Caption         =   "Conta e Histórico Contábil"
      End
      Begin VB.Menu mnuCCCorretoras 
         Caption         =   "Conta Corrente Corretoras"
      End
      Begin VB.Menu mnuParamReprocCC 
         Caption         =   "Parametrização Reprocessamento"
      End
   End
   Begin VB.Menu mnuFerramentas 
      Caption         =   "&Ferramentas"
      Begin VB.Menu mnuFerrComplOperCompromissada 
         Caption         =   "Complementação Operação Compromissada"
      End
      Begin VB.Menu mnuFerrConfirmacaoOperMsgSPB 
         Caption         =   "Confirmação Operação/Mensagem SPB"
      End
      Begin VB.Menu mnuSep094 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFerrReenvioCancEstornoMsgSPB 
         Caption         =   "Reenvio, Cancelamento e Estorno de Mensagem SPB"
      End
      Begin VB.Menu mnuFerramentasSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFerrEntradaManual 
         Caption         =   "Entrada Manual"
         Begin VB.Menu mnuFerrEntradaManualOperacao 
            Caption         =   "Operação"
         End
         Begin VB.Menu mnuFerrEntradaManualOperacaoBeta 
            Caption         =   "Operação Beta"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFerrEntradaManualMsgSPB 
            Caption         =   "Mensagem SPB"
         End
      End
      Begin VB.Menu mnuFerrCancelamentoEntradaManual 
         Caption         =   "Cancelamento Entrada Manual"
      End
      Begin VB.Menu mnuFerramentasSep05 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFerrSoliRepasseFincPagtoDespesas 
         Caption         =   "Solicitação de Repasse Financeiro / Pagamento de Dispesas"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFerramentaMensagemInconsistente 
         Caption         =   "Reprocessamento de mensagens Inconsistentes"
      End
      Begin VB.Menu mnuFerrParamAtividade 
         Caption         =   "Parametrização de Atividade"
      End
      Begin VB.Menu mnuFerramentasSep06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFerrImportacaoArquivo 
         Caption         =   "Importação Arquivo"
         Begin VB.Menu mnuFerrImportacaoArquivoCBLC 
            Caption         =   "Arquivo CBLC"
         End
         Begin VB.Menu mnuFerrImportacaoArquivoBMD 
            Caption         =   "Arquivo BMF (Derivativos)"
         End
      End
   End
   Begin VB.Menu mnuConciliacao 
      Caption         =   "C&onciliação"
      Begin VB.Menu mnuFerrConciliacaoOperMsgSPB 
         Caption         =   "Operações"
      End
      Begin VB.Menu mnuSep394389 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoRegistroOperacaoBO 
         Caption         =   "Registros de Operações (BackOffice)"
      End
      Begin VB.Menu mnuFerramentasSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoTitulosBO 
         Caption         =   "Liquidação Física (BackOffice)"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoFinanceiraBO 
         Caption         =   "Liquidação Financeira Bruta / Bilateral NET (Backoffice)"
      End
      Begin VB.Menu mnuConciliacaoFinanceiraBilateralBO 
         Caption         =   "Liquidação Financeira Bilateral LTR0001 (Backoffice)"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoFinanceiraMultilateralBO 
         Caption         =   "Liquidação Financeira Multilateral (Backoffice)"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoCorretorasBO 
         Caption         =   "Liquidação de Corretoras (Backoffice)"
      End
      Begin VB.Menu mnuFerramentasSep91 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoCompGenericaBO 
         Caption         =   "NET Compromissada Genérica (Backoffice)"
      End
      Begin VB.Menu mnuConciliacaoSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoBrutaBO 
         Caption         =   "Liquidação Bruta - CBLC (Backoffice)"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoMultilateralBO 
         Caption         =   "Liquidação Multilateral - CBLC (Backoffice)"
      End
      Begin VB.Menu mnuConciliacaoSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoMultilateralBMD_BO 
         Caption         =   "Liquidação Multilateral - BMD (Backoffice)"
      End
      Begin VB.Menu mnuConciliacaoSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConciliacaoPagamentoDespesas 
         Caption         =   "Pagamento de Despesas via STR0007 (Backoffice)"
      End
   End
   Begin VB.Menu mnuLiberacao 
      Caption         =   "&Liberação"
      Begin VB.Menu mnuFerrLiberacaoOperMsgSPB 
         Caption         =   "Liberação Operação/Mensagem SPB"
      End
      Begin VB.Menu mnuLiberacaoSep34893 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoRegistroOperacaoAA 
         Caption         =   "Registros de Operações (Administrador de Área)"
      End
      Begin VB.Menu mnuLiberacaoSep384737 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoTitulosAA 
         Caption         =   "Liquidação Física (Administrador de Área)"
      End
      Begin VB.Menu mnuLiberacaoSep38947 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoFinanceiraAA 
         Caption         =   "Liquidação Financeira Bruta / Bilateral NET (Administrador de Área)"
      End
      Begin VB.Menu mnuConciliacaoFinanceiraBilateralAA 
         Caption         =   "Liquidação Financeira Bilateral LTR0001 (Administrador de Área)"
      End
      Begin VB.Menu mnuLiberacaoLiquidacaoDespesaCETIP 
         Caption         =   "Liquidação Despesa CETIP (Administrador Geral)"
      End
      Begin VB.Menu mnuLiberacaoSep2934 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoFinanceiraMultilateralAA 
         Caption         =   "Liquidação Financeira Multilateral (Administrador de Área)"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoFinanceiraMultilateralAG 
         Caption         =   "Liquidação Financeira Multilateral (Administrador Geral)"
      End
      Begin VB.Menu mnuLiberacaoSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoCorretorasAA 
         Caption         =   "Liquidação de Corretoras (Administrador de Área)"
      End
      Begin VB.Menu mnuLiberacaoSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConciliacaoLiquidacaoCompGenericaAA 
         Caption         =   "NET Compromissada Genérica (Administrador de Área)"
      End
      Begin VB.Menu mnuLiberacaoSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiberacaoLiquidacaoBrutaAA 
         Caption         =   "Liquidação Bruta - CBLC (Administrador de Área)"
      End
      Begin VB.Menu mnuLiberacaoSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiberacaoLiquidacaoMultilateralAA 
         Caption         =   "Liquidação Multilateral - CBLC (Administrador de Área)"
      End
      Begin VB.Menu mnuLiberacaoLiquidacaoMultilateralAG 
         Caption         =   "Liquidação Multilateral - CBLC (Administrador Geral)"
      End
      Begin VB.Menu mnuLiberacaoSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiberacaoLiquidacaoEventosAA 
         Caption         =   "Liquidação Eventos - CBLC (Administrador de Área)"
      End
      Begin VB.Menu mnuLiberacaoLiquidacaoEventosAG 
         Caption         =   "Liquidação Eventos - CBLC (Administrador Geral)"
      End
      Begin VB.Menu mnuLiberacaoSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiberacaoLiquidacaoBMD_AA 
         Caption         =   "Liquidação Multilateral - BMD (Administrador de Área)"
      End
      Begin VB.Menu mnuLiberacaoLiquidacaoBMD_AGP 
         Caption         =   "Liquidação Multilateral - BMD (Administrador Geral - Prévia)"
      End
      Begin VB.Menu mnuLiberacaoLiquidacaoBMD_AGD 
         Caption         =   "Liquidação Multilateral - BMD (Administrador Geral - Definitiva)"
      End
      Begin VB.Menu mnuLiberacaoSep05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiberacaoRegistroRodaDolarBMC 
         Caption         =   "Registro de Operações Roda de Dólar Pronto - BMC"
      End
      Begin VB.Menu mnuLiberacaoSep07 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiberacaoPagamentoDespesas 
         Caption         =   "Pagamento de Despesas via STR0007 (Administrador de Área)"
      End
   End
   Begin VB.Menu mnuAlcada 
      Caption         =   "Con&trole Alçada"
      Begin VB.Menu mnuAlcadaAA 
         Caption         =   "Liberação Alçada Administrador de Área"
      End
      Begin VB.Menu mnuAlcadaAG 
         Caption         =   "Liberação Alçada Administrador Geral"
      End
   End
   Begin VB.Menu mnuContingencia 
      Caption         =   "Conti&ngência"
      Begin VB.Menu mnuContingParamSituacao 
         Caption         =   "Parametrização Situação"
      End
      Begin VB.Menu mnuContingBaixarLiquidarOper 
         Caption         =   "Baixar/Liquidar Operação"
      End
   End
   Begin VB.Menu mnuContaCorrente 
      Caption         =   "Conta Co&rrente"
      Begin VB.Menu mnuCCSuspenderDisponibilizarCancelarLanc 
         Caption         =   "Suspender, Disponibilizar ou Cancelar Lançamento"
      End
      Begin VB.Menu mnumnuContaCorrenteSep01 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCCIntegracaoOnLine 
         Caption         =   "Integração Conta Corrente"
      End
      Begin VB.Menu mnuCCIntegracaoOnLineEstorno 
         Caption         =   "Integração Conta Corrente - Estorno"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCCIntegracaoBatch 
         Caption         =   "Integração Contabilidade"
      End
   End
   Begin VB.Menu mnuConsulta 
      Caption         =   "Consul&ta"
      Begin VB.Menu mnuConsOperacao 
         Caption         =   "Operações"
      End
      Begin VB.Menu mnuConsMovimentacao 
         Caption         =   "Movimentação"
         Begin VB.Menu mnuConsMovimentacaoOper 
            Caption         =   "Operações"
         End
         Begin VB.Menu mnuConsMovimentacaoRejeitada 
            Caption         =   "Remessa Rejeitada"
         End
         Begin VB.Menu mnuConsMovimentacaoCC_HA 
            Caption         =   "Conta Corrente e Contabilidade"
         End
      End
      Begin VB.Menu mnuConsMensagemSPB 
         Caption         =   "Mensagens SPB"
      End
      Begin VB.Menu mnuConsLancamentoCC 
         Caption         =   "Lançamentos Conta Corrente"
      End
      Begin VB.Menu mnuConsRemessaRejeitada 
         Caption         =   "Remessas Rejeitadas"
      End
      Begin VB.Menu mnuConsultaVeiculoLegal 
         Caption         =   "Veiculo Legal "
      End
      Begin VB.Menu mnuConsultaLiquidacaoMultilateral 
         Caption         =   "Liquidação Multilateral"
      End
      Begin VB.Menu mnuConciliacaoCCR 
         Caption         =   "Resumo Diário CCR"
      End
      Begin VB.Menu mnuConsultaOperCCR 
         Caption         =   "Consulta Operações CCR"
      End
      Begin VB.Menu mnuConsultaMesgSisbacen 
         Caption         =   "Consulta Mensagens Sisbacen"
      End
   End
   Begin VB.Menu mnuJanela 
      Caption         =   "&Janela"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascata 
         Caption         =   "Cascata"
      End
      Begin VB.Menu mnuHorizontal 
         Caption         =   "Horizontal"
      End
      Begin VB.Menu mnuVertical 
         Caption         =   "Vertical"
      End
      Begin VB.Menu mnuseparatorJan 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFecharTodas 
         Caption         =   "Fechar Todas"
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "Aj&uda"
      Begin VB.Menu mnuAjudaManual 
         Caption         =   "Manual do Usuário"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAjudaSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSobre 
         Caption         =   "Sobre o Sistema"
      End
   End
End
Attribute VB_Name = "mdiLQS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objFrmCadastroVeiculoLegal As frmConsultaVeiculoLegal
Dim objFrmConsultaVeiculoLegal As frmConsultaVeiculoLegal

Private Sub ctlSysTray1_MouseDblClick(Button As Integer, Id As Long)

On Error GoTo ErrorHandler

    frmAlerta.Show
    frmAlerta.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - ctlSysTray1_MouseDblClick"
End Sub

'' Encaminhar a solicitação (Obtenção do Tipo de Back Office do usuário) à camada
'' controladora de caso de uso (componente / classe / metodo ) :
''
'' A8MIU.clsControleAcessDado.ObterTipoBackOfficeUsuario
''
'' O método retornará uma String para a camada de interface
''
Private Sub flConfigurarBackOffice()

#If EnableSoap = 1 Then
    Dim objControleAcesso                   As MSSOAPLib30.SoapClient30
#Else
    Dim objControleAcesso                   As A8MIU.clsControleAcessDado
#End If
    
Dim strTipoBackOffice                       As String
Dim strXMLErro                              As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
    
On Error GoTo ErrorHandler

    Set objControleAcesso = fgCriarObjetoMIU("A8MIU.clsControleAcessDado")
    strTipoBackOffice = objControleAcesso.ObterTipoBackOfficeUsuario(vntCodErro, _
                                                                     vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    staLQS.Panels("BackOffice").Text = "Back Office : " & objControleAcesso.ObterDescricaoTipoBackoffice(strTipoBackOffice)
    Set objControleAcesso = Nothing
    
    gintTipoBackoffice = Val(strTipoBackOffice)
    
Exit Sub
ErrorHandler:
    Set objControleAcesso = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    Else
        strXMLErro = Err.Description
    End If
    
    mdiLQS.uctlogErros.MostrarErros Err, "mdiLQS - flConfigurarBackOffice", Me.Caption
    
    If fgObterCodigoDeErroDeNegocioXMLErro(strXMLErro) = "40" Then
       '40 - Tipo BackOffice não cadastrado para o usuário
       MsgBox "O Tipo BackOffice é obrigatório" & vbNewLine & _
              "O Sistema será finalizado.", vbCritical, App.Title
       End
    ElseIf fgObterCodigoDeErroDeNegocioXMLErro(strXMLErro) = "38" Then
       '38 - Usuário associado a mais de um Tipo Back Office
       MsgBox "Usuário associado a mais de um Tipo Back Office" & vbNewLine & _
              "O Sistema será finalizado.", vbCritical, App.Title
       End
    End If
    
End Sub

Private Sub MDIForm_Load()
    
On Error GoTo ErrorHandler
    
    fgCursor True
    
    DoEvents
    Me.Show
            
    Call fgObterIntervaloVerificacao
    Call flConfigurarBackOffice
    Call fgCarregarXMLGeralTelaFiltro
   
    Me.Caption = Me.Caption & " - " & gstrAmbiente
    App.HelpFile = App.Path & "\" & gstrHelpFile
    
    fgCursor False
    
Exit Sub
ErrorHandler:

    MsgBox "Erro-> " & Err.Description & vbCrLf & "O Sistema será finalizado "
    
    Set mdiLQS = Nothing
    End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

Dim objForm                                 As Form

On Error GoTo ErrorHandler

    For Each objForm In Forms
        If objForm.Name <> Me.Name Then
            Unload objForm
        End If
    Next

    fgDesregistraComponentes
    
    Set mdiLQS = Nothing
    
    End
    
Exit Sub
ErrorHandler:

    MsgBox "Erro-> " & Err.Description & vbCrLf & "O Sistema será finalizado "
    
    Set mdiLQS = Nothing
    
    End

End Sub

Private Sub mnuAjudaManual_Click()
    Dim hwndHelp                            As Long
    hwndHelp = HtmlHelp(Me.hwnd, App.HelpFile, HH_DISPLAY_TOPIC, 0)
End Sub

Private Sub mnuAlcadaAA_Click()

On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmLiberacaoOperacaoMensagem.PerfilAcesso = AdmArea
    DoEvents
    
    frmLiberacaoOperacaoMensagem.Show
    frmLiberacaoOperacaoMensagem.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuAlcadaAA_Click"
End Sub

Private Sub mnuAlcadaAG_Click()

On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmLiberacaoOperacaoMensagem.PerfilAcesso = AdmGeral
    DoEvents
    
    frmLiberacaoOperacaoMensagem.Show
    frmLiberacaoOperacaoMensagem.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuAlcadaAA_Click"
End Sub

Private Sub mnuCadVeicLegalXEmpresa_Click()

    On Error GoTo ErrorHandler
    
    frmAlteracaoEmpresaVeiculoLegal.Show
    frmAlteracaoEmpresaVeiculoLegal.ZOrder vbBringToFront

    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuCadVeicLegalXEmpresa_Click"

End Sub

Private Sub mnuCCCorretoras_Click()

On Error GoTo ErrorHandler

    frmCadastroContaCOTR.Show
    frmCadastroContaCOTR.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuParamHistCC_Click"
End Sub




Private Sub mnuConciliacaoCCR_Click()

    On Error GoTo ErrorHandler

    DoEvents
    
    frmConciliacaoCCR.Show
    frmConciliacaoCCR.ZOrder

    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoCCR_Click"


End Sub

Private Sub mnuConciliacaoFinanceiraBilateralAA_Click()

    On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmConciliacaoFinanceiraBilateral.PerfilAcesso = AdmArea
    DoEvents
    
    frmConciliacaoFinanceiraBilateral.Show
    frmConciliacaoFinanceiraBilateral.ZOrder

    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoFinanceiraBilateralAA_Click"

End Sub

Private Sub mnuConciliacaoFinanceiraBilateralBO_Click()

    On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmConciliacaoFinanceiraBilateral.PerfilAcesso = BackOffice
    DoEvents
    
    frmConciliacaoFinanceiraBilateral.Show
    frmConciliacaoFinanceiraBilateral.ZOrder

    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoFinanceiraBilateralBO_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoBrutaBO_Click()

On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmLiquidacaoBruta.PerfilAcesso = BackOffice
    DoEvents
    
    frmLiquidacaoBruta.Show
    frmLiquidacaoBruta.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoBrutaBO_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoCompGenericaAA_Click()
    
On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmCompromissadaGenerica.PerfilAcesso = AdmArea
    DoEvents
    
    frmCompromissadaGenerica.Show
    frmCompromissadaGenerica.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoCompGenericaAA_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoCompGenericaBO_Click()
    
On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmCompromissadaGenerica.PerfilAcesso = BackOffice
    DoEvents
    
    frmCompromissadaGenerica.Show
    frmCompromissadaGenerica.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoCompGenericaBO_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoCorretorasAA_Click()
    
On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmLiquidacaoCorretoras.PerfilAcesso = AdmArea
    DoEvents
    
    frmLiquidacaoCorretoras.Show
    frmLiquidacaoCorretoras.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoCorretorasAA_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoCorretorasBO_Click()
    
On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmLiquidacaoCorretoras.PerfilAcesso = BackOffice
    DoEvents
    
    frmLiquidacaoCorretoras.Show
    frmLiquidacaoCorretoras.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoCorretorasBO_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoMultilateralBMD_BO_Click()

On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmLiquidacaoMultilateralBMF.PerfilAcesso = BackOffice
    DoEvents
    
    frmLiquidacaoMultilateralBMF.Show
    frmLiquidacaoMultilateralBMF.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoMultilateralBMD_BO_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoMultilateralBO_Click()

On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmLiquidacaoMultilateralCBLC.PerfilAcesso = BackOffice
    DoEvents
    
    frmLiquidacaoMultilateralCBLC.Show
    frmLiquidacaoMultilateralCBLC.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoMultilateralBO_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoRegistroOperacaoBO_Click()

On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmConciliacaoRegistroOperacao.PerfilAcesso = BackOffice
    
    frmConciliacaoRegistroOperacao.Show
    frmConciliacaoRegistroOperacao.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoRegistroOperacaoBO_Click"

End Sub

Private Sub mnuCadHorarioLimiteIntSistLeg_Click()

On Error GoTo ErrorHandler

    frmHorarioLimiteIntegracao.Show
    frmHorarioLimiteIntegracao.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuCadHorarioLimiteIntSistLeg_Click"
End Sub

Private Sub mnuCadParamAlerta_Click()

On Error GoTo ErrorHandler

    frmCadastroAlerta.Show
    frmCadastroAlerta.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuCadParamAlerta_Click"
End Sub

Private Sub mnuCadParamWorkflow_Click()

On Error GoTo ErrorHandler

    frmCadastroWorkflow.Show
    frmCadastroWorkflow.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuCadParamWorkflow_Click"
End Sub

Private Sub mnuCadVeicLegalXGrupoVeicLegal_Click()

On Error GoTo ErrorHandler
    
    'identifica que a tela se comportará como cadastro
    If objFrmCadastroVeiculoLegal Is Nothing Then Set objFrmCadastroVeiculoLegal = New frmConsultaVeiculoLegal
    objFrmCadastroVeiculoLegal.blnConsulta = False
    objFrmCadastroVeiculoLegal.Show
    objFrmCadastroVeiculoLegal.ZOrder
    
Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuCadVeicLegalXGrupoVeicLegal_Click"
End Sub

Private Sub mnuCCIntegracaoBatch_Click()

On Error GoTo ErrorHandler

    frmIntegracaoCCBatch.Show
    frmIntegracaoCCBatch.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - frmIntegracaoCCBatch_Click"

End Sub

Private Sub mnuCCIntegracaoOnLine_Click()

Dim objForm                                 As frmIntegrarCCOnLine

On Error GoTo ErrorHandler
    
    Set objForm = flEncontraFormPorTag("IntegrarCCOnLine")
    
    If objForm Is Nothing Then
        Set objForm = New frmIntegrarCCOnLine
        With objForm
            .Tag = "IntegrarCCOnLine"
            .Caption = "Conta Corrente - Integrar Conta Corrente"
        End With
    End If
    
    With objForm
        .Show
        .ZOrder
    End With

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - frmIntegrarCCOnLine_Click"

End Sub

Private Sub mnuCCSuspenderDisponibilizarCancelarLanc_Click()
    
On Error GoTo ErrorHandler

    frmSuspenderDisponibilizarLancamentoCC.Show
    frmSuspenderDisponibilizarLancamentoCC.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuCCSuspenderDisponibilizarCancelarLanc_Click"

End Sub



Private Sub mnuConciliacaoPagamentoDespesas_Click()
On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmLiquidacaoPagDespesas.PerfilAcesso = BackOffice
    DoEvents
    
    frmLiquidacaoPagDespesas.Show
    frmLiquidacaoPagDespesas.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoPagamentoDespesas_Click"

End Sub

Private Sub mnuConsLancamentoCC_Click()

On Error GoTo ErrorHandler

   frmConsultaContaCorrente.Show
   frmConsultaContaCorrente.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - frmConsultaContaCorrente_Click"


End Sub

Private Sub mnuConsMensagemSPB_Click()

On Error GoTo ErrorHandler

   frmConsultaMensagem.Show
   frmConsultaMensagem.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConsMensagemSPB_Click"
End Sub

Private Sub mnuConsultaLiquidacaoMultilateral_Click()

On Error GoTo ErrorHandler

    frmConsultaLiquidacaoMultilateral.Show
    frmConsultaLiquidacaoMultilateral.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConsultaLiquidacaoMultilateral_Click"

End Sub

Private Sub mnuConsultaMesgSisbacen_Click()

On Error GoTo ErrorHandler

    frmConsultaMensagemSisbacen.Show
    frmConsultaMensagemSisbacen.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - frmConsultaMensagemSisbacen_Click"

End Sub

Private Sub mnuConsultaOperCCR_Click()


On Error GoTo ErrorHandler

    frmConsultaOperacaoCCR.Show
    frmConsultaOperacaoCCR.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - frmConsultaOperacaoCCR_Click"

End Sub

Private Sub mnuConsultaVeiculoLegal_Click()

On Error GoTo ErrorHandler
    
    'identifica que a tela se comportará como consulta
    If objFrmConsultaVeiculoLegal Is Nothing Then Set objFrmConsultaVeiculoLegal = New frmConsultaVeiculoLegal
    objFrmConsultaVeiculoLegal.blnConsulta = True
    objFrmConsultaVeiculoLegal.Show
    objFrmConsultaVeiculoLegal.ZOrder
  
Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConsultaVeiculoLegal_Click"

End Sub

Private Sub mnuContingBaixarLiquidarOper_Click()

On Error GoTo ErrorHandler

    frmAlteracaoStatusOperacao.Show
    frmAlteracaoStatusOperacao.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuContingBaixarLiquidarOper_Click"
End Sub

Private Sub mnuContingParamSituacao_Click()

On Error GoTo ErrorHandler

    frmParametrizacaoSituacaoContingencia.Show
    frmParametrizacaoSituacaoContingencia.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuContingParamSituacao_Click"
End Sub

Private Sub mnuFerramentaMensagemInconsistente_Click()

On Error GoTo ErrorHandler

    frmMensagemIncosistente.Show
    frmMensagemIncosistente.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFerramentaMensagemInconsistente_Click"

End Sub

Private Sub mnuFerrCancelamentoEntradaManual_Click()

On Error GoTo ErrorHandler

    frmCancelamentoEntradaManual.Show
    frmCancelamentoEntradaManual.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFerrCancelamentoEntradaManual_Click"
End Sub

Private Sub mnuConciliacaoLiquidacaoFinanceiraAA_Click()
    
On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmConciliacaoFinanceira.PerfilAcesso = AdmArea
    DoEvents
    
    frmConciliacaoFinanceira.Show
    frmConciliacaoFinanceira.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoFinanceiraAA_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoFinanceiraBO_Click()
    
On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmConciliacaoFinanceira.PerfilAcesso = BackOffice
    DoEvents
    
    frmConciliacaoFinanceira.Show
    frmConciliacaoFinanceira.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoFinanceiraBO_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoFinanceiraMultilateralBO_Click()

On Error GoTo ErrorHandler

    frmConciliacaoFinanceiraMultilateral.PerfilAcesso = BackOffice

    frmConciliacaoFinanceiraMultilateral.Show
    frmConciliacaoFinanceiraMultilateral.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoTitulosBO_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoFinanceiraMultilateralAA_Click()

On Error GoTo ErrorHandler

    frmConciliacaoFinanceiraMultilateral.PerfilAcesso = AdmArea

    frmConciliacaoFinanceiraMultilateral.Show
    frmConciliacaoFinanceiraMultilateral.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoTitulos_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoFinanceiraMultilateralAG_Click()

On Error GoTo ErrorHandler

    frmConciliacaoFinanceiraMultilateral.PerfilAcesso = AdmGeral

    frmConciliacaoFinanceiraMultilateral.Show
    frmConciliacaoFinanceiraMultilateral.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoTitulos_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoRegistroOperacaoAA_Click()
On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmConciliacaoRegistroOperacao.PerfilAcesso = AdmArea
    
    frmConciliacaoRegistroOperacao.Show
    frmConciliacaoRegistroOperacao.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoRegistroOperacaoAA_Click"

End Sub

Private Function flEncontraFormPorTag(ByVal strTag As String) As Object

Dim objForm                                 As Form

On Error GoTo ErrorHandler

    For Each objForm In Forms
        If objForm.Tag = strTag Then
            Set flEncontraFormPorTag = objForm
            Exit Function
        End If
    Next objForm

    Set flEncontraFormPorTag = Nothing

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flEncontraFormPorTag", 0
End Function

Private Sub mnuConciliacaoLiquidacaoTitulosAA_Click()

Dim objForm                                 As frmConciliacaoTitulos

On Error GoTo ErrorHandler

    Set objForm = flEncontraFormPorTag("TituloAdminArea")
    
    If objForm Is Nothing Then
        Set objForm = New frmConciliacaoTitulos
        'Configura o Perfil de acesso
        objForm.PerfilAcesso = enumPerfilAcesso.AdmArea
        objForm.Tag = "TituloAdminArea"
        objForm.Caption = "Ferramentas - Conciliação e Liquidação de Títulos (BMA) - Administrador de Área"
    End If

    DoEvents
    
    objForm.Show
    objForm.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoTitulosAA_Click"

End Sub

Private Sub mnuConciliacaoLiquidacaoTitulosBO_Click()
    
Dim objForm                                 As frmConciliacaoTitulos

On Error GoTo ErrorHandler

    Set objForm = flEncontraFormPorTag("TituloBackOffice")
    
    If objForm Is Nothing Then
        Set objForm = New frmConciliacaoTitulos
        'Configura o Perfil de acesso
        objForm.PerfilAcesso = enumPerfilAcesso.BackOffice
        objForm.Tag = "TituloBackOffice"
        objForm.Caption = "Ferramentas - Conciliação e Liquidação de Títulos (BMA) - BackOffice"
    End If

    DoEvents
    
    objForm.Show
    objForm.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConciliacaoLiquidacaoTitulosBO_Click"

End Sub

Private Sub mnuFerrEntradaManualMsgSPB_Click()

On Error GoTo ErrorHandler

    frmEntradaManualSPB.Show
    frmEntradaManualSPB.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFerrEntradaManualMsgSPB_Click"
End Sub

Private Sub mnuFerrEntradaManualOperacaoBeta_Click()

On Error GoTo ErrorHandler

    frmEntradaManualSLCCBeta.Show
    frmEntradaManualSLCCBeta.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFerrSoliRepasseFincPagtoDespesas_Click"

End Sub


Private Sub mnuFerrImportacaoArquivoBMD_Click()

On Error GoTo ErrorHandler

    frmImportacaoArquivoBMF.Show
    frmImportacaoArquivoBMF.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFerrImportacaoArquivoBMD_Click"

End Sub

Private Sub mnuFerrImportacaoArquivoCBLC_Click()

On Error GoTo ErrorHandler

    frmImportacaoArquivoCBLC.Show
    frmImportacaoArquivoCBLC.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFerrImportacaoArquivoCBLC_Click"

End Sub

Private Sub mnuFerrParamAtividade_Click()

    On Error GoTo ErrorHandler

    frmControleFluxoAtividade.Show
    frmControleFluxoAtividade.ZOrder

    Exit Sub

ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuCadParamAtividade_Click"

End Sub

Private Sub mnuFerrSoliRepasseFincPagtoDespesas_Click()

On Error GoTo ErrorHandler

    frmSoliRepasseFincPagtoDespesas.Show
    frmSoliRepasseFincPagtoDespesas.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFerrSoliRepasseFincPagtoDespesas_Click"
End Sub

Private Sub mnuIdPartCamVeicLega_Click()

On Error GoTo ErrorHandler

    frmCadastroIdentificadorPartCamara.Show
    frmCadastroIdentificadorPartCamara.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConsultaVeiculoLegal_Click"
End Sub

Private Sub mnuLiberacaoLiquidacaoBMD_AA_Click()

On Error GoTo ErrorHandler
    
    'Configura o Perfil de acesso
    frmLiquidacaoMultilateralBMF.PerfilAcesso = AdmArea
    DoEvents
    
    frmLiquidacaoMultilateralBMF.Show
    frmLiquidacaoMultilateralBMF.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuLiberacaoLiquidacaoBMD_AA_Click"

End Sub

Private Sub mnuLiberacaoLiquidacaoBMD_AGD_Click()

On Error GoTo ErrorHandler
    
    'Configura o Perfil de acesso
    frmLiquidacaoMultilateralBMF.PerfilAcesso = AdmGeral
    DoEvents
    
    frmLiquidacaoMultilateralBMF.Show
    frmLiquidacaoMultilateralBMF.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuLiberacaoLiquidacaoBMD_AGD_Click"

End Sub

Private Sub mnuLiberacaoLiquidacaoBMD_AGP_Click()

On Error GoTo ErrorHandler
    
    'Configura o Perfil de acesso
    frmLiquidacaoMultilateralBMF.PerfilAcesso = AdmGeralPrevia
    DoEvents
    
    frmLiquidacaoMultilateralBMF.Show
    frmLiquidacaoMultilateralBMF.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuLiberacaoLiquidacaoBMD_AGP_Click"

End Sub

Private Sub mnuLiberacaoLiquidacaoBrutaAA_Click()
    
On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmLiquidacaoBruta.PerfilAcesso = AdmArea
    DoEvents
    
    frmLiquidacaoBruta.Show
    frmLiquidacaoBruta.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuLiberacaoLiquidacaoBrutaAA_Click"


End Sub

Private Sub mnuLiberacaoLiquidacaoDespesaCETIP_Click()

On Error GoTo ErrorHandler

    frmLiquidacaoDespesaCETIP.Show
    frmLiquidacaoDespesaCETIP.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuLiberacaoLiquidacaoDespesaCETIP_Click"

End Sub

Private Sub mnuLiberacaoLiquidacaoEventosAA_Click()

On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmConciliacaoEventos.PerfilAcesso = AdmArea
    DoEvents
    
    frmConciliacaoEventos.Show
    frmConciliacaoEventos.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuLiberacaoLiquidacaoEventosAA_Click"

End Sub

Private Sub mnuLiberacaoLiquidacaoEventosAG_Click()

On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmConciliacaoEventos.PerfilAcesso = AdmGeral
    DoEvents
    
    frmConciliacaoEventos.Show
    frmConciliacaoEventos.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuLiberacaoLiquidacaoEventosAA_Click"

End Sub

Private Sub mnuLiberacaoLiquidacaoMultilateralAA_Click()

On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmLiquidacaoMultilateralCBLC.PerfilAcesso = AdmArea
    DoEvents
    
    frmLiquidacaoMultilateralCBLC.Show
    frmLiquidacaoMultilateralCBLC.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuLiberacaoLiquidacaoMultilateralAA_Click"

End Sub

Private Sub mnuLiberacaoLiquidacaoMultilateralAG_Click()

On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmLiquidacaoMultilateralCBLC.PerfilAcesso = AdmGeral
    DoEvents
    
    frmLiquidacaoMultilateralCBLC.Show
    frmLiquidacaoMultilateralCBLC.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuLiberacaoLiquidacaoMultilateralAG_Click"

End Sub

Private Sub mnuLiberacaoPagamentoDespesas_Click()
On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmLiquidacaoPagDespesas.PerfilAcesso = AdmArea
    DoEvents
    
    frmLiquidacaoPagDespesas.Show
    frmLiquidacaoPagDespesas.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuLiberacaoPagamentoDespesas_Click"

End Sub

Private Sub mnuLiberacaoRegistroRodaDolarBMC_Click()

On Error GoTo ErrorHandler
    
    frmRodaDolarPronto.Show
    frmRodaDolarPronto.ZOrder

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuLiberacaoRegistroRodaDolarBMC_Click"

End Sub

Private Sub mnuParamHistCC_Click()

On Error GoTo ErrorHandler

    frmCadastroConta.Show
    frmCadastroConta.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuParamHistCC_Click"
End Sub

Private Sub mnuParamHistCNTB_Click()

On Error GoTo ErrorHandler

    frmParamHistCntaCntb.Show
    frmParamHistCntaCntb.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuParamHistCNTB_Click"
End Sub

Private Sub mnuParamReprocCC_Click()

On Error GoTo ErrorHandler
    
    frmParamReprocCC.Show
    frmParamReprocCC.ZOrder
    
    Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuParamReprocCC_Click"
End Sub



Private Sub mnuSegrGrupoUsuarioMsgSPB_Click()

On Error GoTo ErrorHandler

    frmAssocGrupoUsuarioMensSPB.Show
    frmAssocGrupoUsuarioMensSPB.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuSegrGrupoUsuarioMsgSPB_Click"
End Sub

Private Sub mnuSegrControleAcessoDados_Click()

On Error GoTo ErrorHandler

    frmCadastroConversaoMBS.Show
    frmCadastroConversaoMBS.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuSegrControleAcessoDados_Click"
End Sub

Private Sub mnuSegrGrupoUsuario_Click()

On Error GoTo ErrorHandler

    frmGrupoUsuario.Show
    frmGrupoUsuario.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuSegrGrupoUsuario_Click"
End Sub

Private Sub mnuSegrGrupoVeiculoLegal_Click()

On Error GoTo ErrorHandler

    frmGrupoVeiculoLegal.Show
    frmGrupoVeiculoLegal.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuSegrGrupoVeiculoLegal_Click"
End Sub

Private Sub mnuCadMargemSegurancaGradeHorario_Click()

On Error GoTo ErrorHandler

    frmGradeHorario.Show
    frmGradeHorario.ZOrder vbBringToFront

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuCadMargemSegurancaGradeHorario_Click"
End Sub

Private Sub mnuCadTipoJustificativaConciliacao_Click()

On Error GoTo ErrorHandler

    frmTipoJustificativaConciliacao.Show
    frmTipoJustificativaConciliacao.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuCadTipoJustificativaConciliacao_Click"
End Sub

Private Sub mnuCascata_Click()
    mdiLQS.Arrange vbCascade
End Sub

Private Sub mnuFerrComplOperCompromissada_Click()

On Error GoTo ErrorHandler

    frmComplementacaoOperacao.Show
    frmComplementacaoOperacao.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFerrComplOperCompromissada_Click"
End Sub

Private Sub mnuFerrConciliacaoOperMsgSPB_Click()

On Error GoTo ErrorHandler

    frmConciliacaoOperacao.Show
    frmConciliacaoOperacao.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFerrConciliacaoOperMsgSPB_Click"
End Sub

Private Sub mnuFerrConfirmacaoOperMsgSPB_Click()

On Error GoTo ErrorHandler

    frmConfirmacaoOperacao.Show
    frmConfirmacaoOperacao.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFerrConfirmacaoOperMsgSPB_Click"
End Sub

Private Sub mnuConsOperacao_Click()

On Error GoTo ErrorHandler

    frmConsultaOperacao.Show
    frmConsultaOperacao.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConsOperacao_Click"
End Sub

Private Sub mnuFerrEntradaManualOperacao_Click()

On Error GoTo ErrorHandler

    frmEntradaManualSLCCBeta.Show
    frmEntradaManualSLCCBeta.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFerrEntradaManualOperacao_Click"
End Sub

Private Sub mnuExportarExcel_Click()

On Error GoTo ErrorHandler

    fgCursor True
    
    If Not Me.ActiveForm Is Nothing Then
        fgExportaExcel Me.ActiveForm
    Else
        MsgBox "Não há formulários abertos à serem exportados para o Excel.", vbInformation, "Atenção"
    End If
    
    fgCursor False
    
Exit Sub
ErrorHandler:

    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, "mdiLQS - mnuExportarExcel_Click"
    
End Sub

Private Sub mnuExportarPDF_Click()

On Error GoTo ErrorHandler

    fgCursor True
    
    If Not Me.ActiveForm Is Nothing Then
        fgExportaPDF Me.ActiveForm
    Else
        MsgBox "Não há formulários abertos à serem exportados para o PDF.", vbInformation, "Atenção"
    End If
    
    fgCursor False
    
Exit Sub
ErrorHandler:

    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, "mdiLQS - mnuExportarPDF_Click"
    
End Sub

Private Sub mnuFecharTodas_Click()

Dim Form                                    As Form

On Error GoTo ErrorHandler

    For Each Form In Forms
        If Not Form.Name = Me.Name Then
            Unload Form
        End If
    Next

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFecharTodas_Click"
End Sub

Private Sub mnuHorizontal_Click()
    mdiLQS.Arrange vbHorizontal
End Sub

Private Sub mnuFerrLIberacaoOperMsgSPB_Click()

On Error GoTo ErrorHandler

    'Configura o Perfil de acesso
    frmLiberacaoOperacaoMensagem.PerfilAcesso = enumPerfilAcesso.Nenhum
    DoEvents
    
    frmLiberacaoOperacaoMensagem.Show
    frmLiberacaoOperacaoMensagem.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFerrLIberacaoOperMsgSPB_Click"
End Sub

Private Sub mnuFerrReenvioCancEstornoMsgSPB_Click()

On Error GoTo ErrorHandler

    frmReenvioCancelamentoEstornoMsg.Show
    frmReenvioCancelamentoEstornoMsg.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuFerrReenvioCancEstornoMsgSPB_Click"
End Sub

Private Sub mnuConsRemessaRejeitada_Click()

On Error GoTo ErrorHandler

    frmRemessaRejeitada.Show
    frmRemessaRejeitada.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConsRemessaRejeitada_Click"
End Sub

Private Sub mnuSair_Click()
    Unload Me
End Sub

Private Sub mnuSobre_Click()

On Error GoTo ErrorHandler

    frmSobre.Show
    frmSobre.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuSobre_Click"
End Sub

Private Sub mnuVertical_Click()
    mdiLQS.Arrange vbVertical
End Sub

Private Sub tmrIntervalo_Timer()

Dim strTip                                  As String

On Error GoTo ErrorHandler

    tmrIntervalo.Interval = 60000

    'Incrementos
    glngContaMinutosAlerta = glngContaMinutosAlerta + 1
    glngContaMinutosContingencia = glngContaMinutosContingencia + 1
    
    If glngTempoAlerta <= 0 Or glngTempoContingencia <= 0 Then
        fgCarregarIntervalos
    End If
    
    'Alerta
    If Not fgDesenv Then
        'se estiver desenvolvimento nao precisa
        If glngTempoAlerta <= glngContaMinutosAlerta Then
            glngContaMinutosAlerta = 0
            Load frmAlerta
            If frmAlerta.VerificarAlertas(strTip) Then
                ctlSysTray1.TrayTip = strTip
                ctlSysTray1.InTray = True
            Else
                ctlSysTray1.InTray = False
            End If
        End If
    End If
    
    'Contingência
    If glngTempoContingencia <= glngContaMinutosContingencia Then
        glngContaMinutosContingencia = 0
        If frmParametrizacaoSituacaoContingencia.ExiteSistemaContingencia Then
            flExibeBotaoContingencia True
        Else
            flExibeBotaoContingencia False
        End If
    End If
    
Exit Sub
ErrorHandler:

    mdiLQS.uctlogErros.MostrarErros Err, "mdiLQS - tmrIntervalo_Timer"
    
End Sub

'' Controla a exibição do aviso de contingência
Private Sub flExibeBotaoContingencia(ByVal pblnExibe As Boolean)

On Error GoTo ErrorHandler

    If pblnExibe Then
        staLQS.Panels("Contingencia").Text = "Sistema em Contingência"
        Set staLQS.Panels("Contingencia").Picture = imgLQS.ListImages(1).ExtractIcon
        staLQS.Panels("Contingencia").Enabled = True
    Else
        staLQS.Panels("Contingencia").Text = Space$(23)
        Set staLQS.Panels("Contingencia").Picture = Nothing
        staLQS.Panels("Contingencia").Enabled = False
    End If

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flExibeBotaoContingencia", 0

End Sub

Private Sub uctLogErros_ErroGerado(ErrNumber As Long, ErrDescription As String, Cancel As Boolean)
    lngErrNumber = ErrNumber
End Sub

Private Sub flDumpNomesMenus()

'Faz um dump no debugwindow dos nomes de todos os menus
'para ajudar na configuracao do ambiente de testes em desenvolvimento (insercao de registros na a8.mbs_funcao)

Dim ctl As Control
    
    For Each ctl In Me.Controls
        If TypeOf ctl Is Menu Then
            Debug.Print ctl.Name
        End If
    Next

End Sub

Private Sub mnuConsMovimentacaoCC_HA_Click()

On Error GoTo ErrorHandler

    frmConsultaMovimentacaoCC_HA.Show
    frmConsultaMovimentacaoCC_HA.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConsMovimentacaoCC_HA_Click"

End Sub

Private Sub mnuConsMovimentacaoOper_Click()

On Error GoTo ErrorHandler

    frmConsultaMovimentacao.Show
    frmConsultaMovimentacao.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConsMovimentacaoOper_Click"

End Sub

Private Sub mnuConsMovimentacaoRejeitada_Click()

On Error GoTo ErrorHandler

    frmConsultaMovimentacaoRejeitada.Show
    frmConsultaMovimentacaoRejeitada.ZOrder

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - mnuConsMovimentacaoRejeitada_Click"

End Sub

