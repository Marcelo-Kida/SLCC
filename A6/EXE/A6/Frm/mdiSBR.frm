VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiSBR 
   BackColor       =   &H8000000C&
   Caption         =   "A6 - Sub-reserva"
   ClientHeight    =   5325
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8130
   Icon            =   "mdiSBR.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin A6.ctlErrorMessage uctLogErros 
      Left            =   2970
      Top             =   2700
      _extentx        =   1191
      _extenty        =   1296
   End
   Begin A6.ctlMenu ctlMenu1 
      Left            =   1320
      Top             =   1140
      _ExtentX        =   2778
      _ExtentY        =   1667
   End
   Begin MSComctlLib.StatusBar staSBR 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5010
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "BackOffice"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9701
            MinWidth        =   9701
            Key             =   "Versao"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Key             =   "UltimaModificacao"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuExportarExcel 
         Caption         =   "Exportar para Excel"
      End
      Begin VB.Menu mnuExportarPDF 
         Caption         =   "Exportar para PDF"
      End
      Begin VB.Menu mnuSeparatorSair 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArquivoSair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnuItemCaixa 
         Caption         =   "&Itens de Caixa"
      End
      Begin VB.Menu mnuItemCaixaTipoOperacao 
         Caption         =   "Itens de Caixa X &Tipo de Operação"
      End
      Begin VB.Menu mnuItemCaixaGrupoVeicLegal 
         Caption         =   "Itens de Caixa X &Grupo de Veículo Legal"
      End
      Begin VB.Menu mnuProdutoTipoOperacao 
         Caption         =   "&Produto X Tipo de Operação"
      End
   End
   Begin VB.Menu mnuSubReserva 
      Caption         =   "Sub &Reserva"
      Begin VB.Menu mnuAjuste 
         Caption         =   "Ajuste de Movimento"
      End
      Begin VB.Menu mnuDivisor1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCaixaFuturo 
         Caption         =   "&Futuro"
      End
      Begin VB.Menu mnuDivisor2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubReservaAbertura 
         Caption         =   "&Abertura"
      End
      Begin VB.Menu mnuSubReservaFechamento 
         Caption         =   "&Fechamento"
      End
      Begin VB.Menu mnuSubReservaMonitoracao 
         Caption         =   "&Monitoração de Movimentação"
      End
      Begin VB.Menu mnuSubReservaResumo 
         Caption         =   "&Resumo"
      End
      Begin VB.Menu mnuD0 
         Caption         =   "&D-Zero"
      End
      Begin VB.Menu mnuSubReservaConsultaAberturaFechamento 
         Caption         =   "Histórico Abertura/Fechamento"
      End
      Begin VB.Menu mnuDivisor3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuControleRemessa 
         Caption         =   "&Controle de Remessa"
      End
      Begin VB.Menu mnuRemessaRejeitada 
         Caption         =   "&Remessa Rejeitada"
      End
   End
   Begin VB.Menu mnuJanelas 
      Caption         =   "&Janelas"
      WindowList      =   -1  'True
      Begin VB.Menu mnuJanelaCascata 
         Caption         =   "&Cascata"
      End
      Begin VB.Menu mnuJanelaHorizontal 
         Caption         =   "&Horizontal"
      End
      Begin VB.Menu mnuJanelaVertical 
         Caption         =   "&Vertical"
      End
      Begin VB.Menu mnuJanelaTraco01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJanelaFecharTodas 
         Caption         =   "&Fechar Todas"
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "A&juda"
      Begin VB.Menu mnuAjudaManual 
         Caption         =   "&Manual do Usuário"
         HelpContextID   =   1
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAjudaSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAjudaSobre 
         Caption         =   "&Sobre o Sistema"
      End
   End
End
Attribute VB_Name = "mdiSBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, agrupar os itens de menu de funcionalidades do sistema
' apresentadas ao usuário.

Option Explicit

' Exibe versão dos componentes do sistema ao usuário.

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
    staSBR.Panels("Versao").Text = strTexto
    
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
    
    Call flConfigurarBackOffice
    Call flExibeVersao
    Call fgCarregarXMLGeralTelaFiltro
    
    Me.Caption = Me.Caption & " - " & gstrAmbiente
    App.HelpFile = App.Path & "\" & gstrHelpFile
    
    fgCursor
    
    Exit Sub
ErrHandler:
    fgCursor
    
    mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - MDIForm_Load"

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

Dim objForm                                 As Form

On Error GoTo ErrorHandler

    For Each objForm In Forms
        If objForm.Name <> Me.Name Then
            Unload objForm
        End If
    Next

    fgDesregistraComponentes '       <-- Inibido temporariamente para os testes com SOAP
    
    Set mdiSBR = Nothing
    
    End
    
    Exit Sub
ErrorHandler:

    MsgBox "Erro-> " & Err.Description & vbCrLf & "O Sistema será finalizado "
    
    Set mdiSBR = Nothing
    
    End
End Sub

Private Sub mnuAjudaManual_Click()
    
    Dim hwndHelp                            As Long
    hwndHelp = HtmlHelp(Me.hwnd, App.HelpFile, HH_DISPLAY_TOPIC, 0)

End Sub

Private Sub mnuAjudaSobre_Click()

On Error GoTo ErrorHandler

    frmSobre.Show
    frmSobre.ZOrder 0

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuAjudaSobre_Click"
   
End Sub

Private Sub mnuAjuste_Click()

On Error GoTo ErrorHandler

    frmAjuste.Show
    frmAjuste.ZOrder 0

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuAjuste_Click"
   
End Sub

Private Sub mnuArquivoSair_Click()

On Error GoTo ErrorHandler

   Unload Me

    Exit Sub

ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuArquivoSair_Click"

End Sub

Private Sub mnuCaixaFuturo_Click()
    
On Error GoTo ErrorHandler

    frmCaixaFuturo.Show
    frmCaixaFuturo.ZOrder 0

    Exit Sub
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuCaixaFuturo_Click"
   
End Sub

Private Sub mnuControleRemessa_Click()

On Error GoTo ErrorHandler

    frmControleRemessa.Show
    frmControleRemessa.ZOrder

    Exit Sub
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuControleRemessa_Click"
   
End Sub

Private Sub mnuD0_Click()

On Error GoTo ErrorHandler

    frmSubReservaD0.Show
    frmSubReservaD0.ZOrder 0

    Exit Sub
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuD0_Click"
    
End Sub

Private Sub mnuExportarExcel_Click()
    
On Error GoTo ErrorHandler

    If Not Me.ActiveForm Is Nothing Then
        With frmGridsExportExcel
            Set .objMyForm = Me.ActiveForm
            .Show vbModal
        End With
    Else
        MsgBox "Não há formulários abertos à serem exportados para o Excel.", vbInformation, "Atenção"
    End If

    Exit Sub

ErrorHandler:
    
    fgCursor
    mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuExportarExcel_Click"
    
End Sub

Private Sub mnuExportarPDF_Click()
    
On Error GoTo ErrorHandler

    If Not Me.ActiveForm Is Nothing Then
        fgExportaPDF Me.ActiveForm
    Else
        MsgBox "Não há formulários abertos à serem exportados para o PDF.", vbInformation, "Atenção"
    End If

    Exit Sub

ErrorHandler:
    
    fgCursor
    mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuExportarPDF_Click"
    
End Sub

Private Sub mnuItemCaixa_Click()

On Error GoTo ErrorHandler

    frmItemCaixa.Show
    frmItemCaixa.ZOrder 0

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuItemCaixa_Click"
   
End Sub

Private Sub mnuItemCaixaGrupoVeicLegal_Click()

On Error GoTo ErrorHandler

    frmItemCaixaGrupoVeicLegal.Show
    frmItemCaixaGrupoVeicLegal.ZOrder 0

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuItemCaixaGrupoVeicLegal_Click"
   
End Sub

Private Sub mnuItemCaixaTipoOperacao_Click()

On Error GoTo ErrorHandler

    frmItemCaixaTipoOperacao.Show
    frmItemCaixaTipoOperacao.ZOrder 0

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuItemCaixaTipoOperacao_Click"
   
End Sub

Private Sub mnuJanelaCascata_Click()

On Error GoTo ErrorHandler

    mdiSBR.Arrange vbCascade

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuJanelaCascata_Click"
   
End Sub

Private Sub mnuJanelaFecharTodas_Click()

Dim Form                                    As Form

On Error GoTo ErrorHandler

    For Each Form In Forms
        If Not Form.Name = Me.Name Then
            Unload Form
        End If
    Next

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuJanelaFecharTodas_Click"

End Sub

Private Sub mnuJanelaHorizontal_Click()
    
    mdiSBR.Arrange vbHorizontal

End Sub

Private Sub mnuJanelaVertical_Click()
    
    mdiSBR.Arrange vbVertical

End Sub

Private Sub mnuProdutoTipoOperacao_Click()

On Error GoTo ErrorHandler

    frmProdutoPJTipoOperacao.Show
    frmProdutoPJTipoOperacao.ZOrder 0

    Exit Sub
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuProdutoTipoOperacao_Click"
   
End Sub

Private Sub mnuRemessaRejeitada_Click()

On Error GoTo ErrorHandler

    frmRemessaRejeitada.Show
    frmRemessaRejeitada.ZOrder 0

    Exit Sub
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuRemessaRejeitada_Click"
   
End Sub

Private Sub mnuSubReservaAbertura_Click()

On Error GoTo ErrorHandler

    frmSubReservaAbertura.Show
    frmSubReservaAbertura.ZOrder 0

    Exit Sub
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuSubReservaAbertura_Click"
    
End Sub

Private Sub mnuSubReservaConsultaAberturaFechamento_Click()

On Error GoTo ErrorHandler

    frmConsultaAberturaFechamento.Show
    frmConsultaAberturaFechamento.ZOrder 0

    Exit Sub
ErrorHandler:
   
   mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuSubReservaConsultaAberturaFechamento_Click"

End Sub

Private Sub mnuSubReservaFechamento_Click()

On Error GoTo ErrorHandler

    frmSubReservaFechamento.Show
    frmSubReservaFechamento.ZOrder 0

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuSubReservaFechamento_Click"
   
End Sub

Private Sub mnuSubReservaMonitoracao_Click()

On Error GoTo ErrorHandler

    frmSubReservaMonitoracao.Show
    frmSubReservaMonitoracao.ZOrder 0

    Exit Sub
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuSubReservaMonitoracao_Click"
    
End Sub

Private Sub mnuSubReservaResumo_Click()

On Error GoTo ErrorHandler

    frmSubReservaResumo.Show
    frmSubReservaResumo.ZOrder 0

    Exit Sub
ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - mnuSubReservaResumo_Click"
    
End Sub

' Configura e exibe tipo de backoffice do usuário logado.

Private Sub flConfigurarBackOffice()

#If EnableSoap = 1 Then
    Dim objControleAcesso   As MSSOAPLib30.SoapClient30
#Else
    Dim objControleAcesso   As A6MIU.clsControleAcesso
#End If

Dim strTipoBackOffice       As String
Dim strXMLErro              As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    intNumeroSequencialErro = 1
    Set objControleAcesso = fgCriarObjetoMIU("A6MIU.clsControleAcesso")
    intNumeroSequencialErro = 2
    strTipoBackOffice = objControleAcesso.ObterTipoBackOfficeUsuario(vntCodErro, _
                                                                     vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    staSBR.Panels("BackOffice").Text = "Back Office : " & objControleAcesso.ObterDescricaoTipoBackoffice(strTipoBackOffice)
    intNumeroSequencialErro = 3
    Set objControleAcesso = Nothing

    Exit Sub

ErrorHandler:
    Set objControleAcesso = Nothing

    strXMLErro = Err.Description
    mdiSBR.uctLogErros.MostrarErros Err, "mdiSBR - flConfigurarBackOffice"

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

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
