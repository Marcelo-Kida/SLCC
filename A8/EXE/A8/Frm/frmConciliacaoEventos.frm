VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConciliacaoEventos 
   Caption         =   "Conciliação de Eventos"
   ClientHeight    =   8760
   ClientLeft      =   990
   ClientTop       =   2115
   ClientWidth     =   13350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   13350
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboMensagem 
      Height          =   315
      ItemData        =   "frmConciliacaoEventos.frx":0000
      Left            =   5010
      List            =   "frmConciliacaoEventos.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   300
      Width           =   2460
   End
   Begin VB.OptionButton optNaturezaMovimento 
      Caption         =   "&Recebimento"
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   8910
      TabIndex        =   4
      Top             =   360
      Width           =   195
   End
   Begin VB.ComboBox cboEmpresa 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   4350
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   12720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoEventos.frx":0022
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoEventos.frx":0134
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoEventos.frx":0246
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoEventos.frx":0598
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoEventos.frx":08EA
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoEventos.frx":0C3C
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoEventos.frx":0F8E
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoEventos.frx":12A8
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoEventos.frx":16FA
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   8430
      Width           =   13350
      _ExtentX        =   23548
      _ExtentY        =   582
      ButtonWidth     =   2328
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Concordar   "
            Key             =   "concordancia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Pagamento "
            Key             =   "pagamento"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Pg. Conting."
            Key             =   "pagamentocontingencia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Regularizar  "
            Key             =   "regularizacao"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Mostrar Operações"
            Key             =   "MostrarOperacao"
            Object.ToolTipText     =   "Mostrar Operações"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Mostrar Mensagens"
            Key             =   "MostrarMensagem"
            Object.ToolTipText     =   "Mostrar Mensagens"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair            "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton optNaturezaMovimento 
      Caption         =   "Pa&gamento"
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   7650
      TabIndex        =   5
      Top             =   360
      Value           =   -1  'True
      Width           =   195
   End
   Begin MSComctlLib.ListView lvwOperacao 
      Height          =   2295
      Left            =   60
      TabIndex        =   1
      Top             =   4140
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Conciliar Operação"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Número Comando"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Número Comando Original"
         Object.Width           =   3792
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Data"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ID Ativo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Qtde Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Qtde a Conciliar"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Qtde Conciliada"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "PU"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMensagem 
      Height          =   3255
      Left            =   60
      TabIndex        =   2
      Top             =   780
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblRecebimento 
      AutoSize        =   -1  'True
      Caption         =   "Recebimento"
      Height          =   195
      Left            =   9150
      TabIndex        =   10
      Top             =   360
      Width           =   945
   End
   Begin VB.Label lblPagamento 
      AutoSize        =   -1  'True
      Caption         =   "Pagamento"
      Height          =   195
      Left            =   7890
      TabIndex        =   9
      Top             =   360
      Width           =   810
   End
   Begin VB.Label lblMensagem 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem"
      Height          =   195
      Left            =   5010
      TabIndex        =   7
      Top             =   60
      Width           =   780
   End
   Begin VB.Label lblConciliacao 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   615
   End
   Begin VB.Image imgDummyH 
      Height          =   60
      Left            =   60
      MousePointer    =   7  'Size N S
      Top             =   4080
      Width           =   14040
   End
End
Attribute VB_Name = "frmConciliacaoEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Conciliação / Liquidação de Eventos CBLC

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlRetornoErro                      As MSXML2.DOMDocument40
Private xmlOperacoes                        As MSXML2.DOMDocument40

'Constantes utilizadas no formulário
Private Const COL_MSG_CO_PARP_CAMR          As Integer = 0
Private Const COL_MSG_TP_PGTO_LDL           As Integer = 1
Private Const COL_MSG_TP_BKOF               As Integer = 1
Private Const COL_MSG_VA_OPER_ATIV          As Integer = 2
Private Const COL_MSG_VA_FINC               As Integer = 3
Private Const COL_MSG_DIFERENCA             As Integer = 4
Private Const COL_MSG_CO_ULTI_SITU_PROC     As Integer = 5

Private Const KEY_MSG_CO_PARP_CAMR          As Integer = 1
Private Const KEY_MSG_TP_PGTO_LDL           As Integer = 2
Private Const KEY_MSG_TP_BKOF               As Integer = 2

Private Const TAG_MSG_NU_CTRL_IF            As Integer = 1
Private Const TAG_MSG_DH_REGT_MESG_SPB      As Integer = 2
Private Const TAG_MSG_NU_SEQU_CNTR_REPE     As Integer = 3
Private Const TAG_MSG_DH_ULTI_ATLZ          As Integer = 4
Private Const TAG_MSG_CO_ULTI_SITU_PROC     As Integer = 5

Private Const COL_OP_CO_VEIC_LEGA           As Integer = 0
Private Const COL_OP_CO_OPER_ATIV           As Integer = 1
Private Const COL_OP_IN_OPER_DEBT_CRED      As Integer = 2
Private Const COL_OP_VA_OPER_ATIV           As Integer = 3
Private Const COL_OP_CO_ULTI_SITU_PROC      As Integer = 4

Private Const KEY_OP_NU_SEQU_OPER_ATIV      As Integer = 1

Private Const strFuncionalidade             As String = "frmConciliacaoEventos"
'------------------------------------------------------------------------------------------
'Fim declaração constantes

Private Enum enumNaturezaMovimento
    Pagamento = 0
    Recebimento = 1
End Enum

Private Enum enumTipoPesquisa
    Operacao = 0
    MENSAGEM = 1
End Enum

Private intAcaoProcessamento                As enumAcaoConciliacao

Private lngPerfil                           As Long
Private blnDummyH                           As Boolean

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

'Calcula as diferenças entre os valores de operação e mensagem
Private Sub flCalcularDiferencasListView()

Dim objListItem                             As ListItem
Dim dblValorOperacao                        As Double
Dim dblValorMensagem                        As Double

On Error GoTo ErrorHandler

    For Each objListItem In lvwMensagem.ListItems
        With objListItem
            dblValorOperacao = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_MSG_VA_OPER_ATIV)))
            dblValorMensagem = fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_MSG_VA_FINC)))

            .SubItems(COL_MSG_DIFERENCA) = fgVlrXml_To_Interface(dblValorMensagem - dblValorOperacao)

            If dblValorMensagem - dblValorOperacao <> 0 Then
                .ListSubItems(COL_MSG_DIFERENCA).ForeColor = vbRed
            End If
        End With
    Next

Exit Sub
ErrorHandler:
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCalcularDiferencasListView", 0)

End Sub

'Mostra os campos de detalhes das operações
Private Sub flCarregarListaDetalheOperacoes()

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strItemKey                              As String

On Error GoTo ErrorHandler

    If lvwMensagem.SelectedItem Is Nothing Then Exit Sub

    strItemKey = lvwMensagem.SelectedItem.Key
    lvwOperacao.ListItems.Clear

    For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(strItemKey))

        With lvwOperacao.ListItems.Add(, "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)

            .Text = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
            .SubItems(COL_OP_CO_OPER_ATIV) = objDomNode.selectSingleNode("CO_OPER_ATIV").Text
            .SubItems(COL_OP_VA_OPER_ATIV) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
            .SubItems(COL_OP_CO_ULTI_SITU_PROC) = objDomNode.selectSingleNode("DE_SITU_PROC").Text

            If objDomNode.selectSingleNode("CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito Then
                .SubItems(COL_OP_IN_OPER_DEBT_CRED) = "Débito"
            Else
                .SubItems(COL_OP_IN_OPER_DEBT_CRED) = "Crédito"
            End If
        
        End With

    Next

    Call fgClassificarListview(Me.lvwOperacao, lngIndexClassifListOper, True)
    
Exit Sub
ErrorHandler:
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaDetalheOperacoes", 0)

End Sub

'Exibe os detalhes de lista de mensagens
Private Sub flCarregarListaNetMensagens(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objMensagem         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem         As A8MIU.clsMensagem
#End If

Dim strRetLeitura           As String
Dim xmlRetLeitura           As MSXML2.DOMDocument40
Dim objDomNode              As MSXML2.IXMLDOMNode
Dim strListItemKey          As String
Dim strListItemTag          As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    strRetLeitura = objMensagem.ObterDetalheMensagem(pstrFiltro, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set objMensagem = Nothing

    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")

        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregarListaNetMensagens")
        End If

        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheMensagem/*")
            strListItemKey = flMontarChaveItemListview(objDomNode)

            strListItemTag = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                             "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                             "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & _
                             "|" & objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & _
                             "|" & objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text

            If Not fgExisteItemLvw(Me.lvwMensagem, strListItemKey) Then
                With lvwMensagem.ListItems.Add(, strListItemKey)

                    .Text = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                    If objDomNode.selectSingleNode("DS_PARP_CAMR").Text <> vbNullString Then
                        .Text = .Text & " - " & objDomNode.selectSingleNode("DS_PARP_CAMR").Text
                    End If

                    If PerfilAcesso = AdmArea Then
                        .SubItems(COL_MSG_TP_PGTO_LDL) = flObterDescricaoTipoPagamento(objDomNode.selectSingleNode("TP_PGTO_LDL").Text)
                    Else
                        .SubItems(COL_MSG_TP_BKOF) = objDomNode.selectSingleNode("DE_BKOF").Text
                    End If

                    .SubItems(COL_MSG_CO_ULTI_SITU_PROC) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                    .SubItems(COL_MSG_VA_FINC) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                    .SubItems(COL_MSG_VA_OPER_ATIV) = " "
                    
                    If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
                        .SubItems(COL_MSG_VA_FINC) = "-" & .SubItems(COL_MSG_VA_FINC)
                    End If

                End With

            Else
                With lvwMensagem.ListItems(strListItemKey)

                    .SubItems(COL_MSG_CO_ULTI_SITU_PROC) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                    .SubItems(COL_MSG_VA_FINC) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                    If optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value Then
                        .SubItems(COL_MSG_VA_FINC) = "-" & .SubItems(COL_MSG_VA_FINC)
                    End If

                End With
            End If

            lvwMensagem.ListItems(strListItemKey).Tag = strListItemTag

        Next
    End If

    Call fgClassificarListview(Me.lvwMensagem, lngIndexClassifListMesg, True)
    
    Set xmlRetLeitura = Nothing

Exit Sub
ErrorHandler:
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetMensagens", 0)

End Sub

'Carregar dados com NET de operações
Private Sub flCarregarListaNetOperacoes(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objOperacao         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao         As A8MIU.clsOperacao
#End If

Dim strRetLeitura           As String
Dim xmlRetLeitura           As MSXML2.DOMDocument40
Dim objDomNode              As MSXML2.IXMLDOMNode
Dim strListItemKey          As String
Dim dblValorOperacao        As Double
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strRetLeitura = objOperacao.ObterDetalheOperacao(pstrFiltro, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing

    Call xmlOperacoes.loadXML(strRetLeitura)

    If strRetLeitura <> vbNullString Then
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")

        If Not xmlRetLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlRetLeitura, App.EXEName, TypeName(Me), "flCarregarListaNetOperacoes")
        End If

        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheOperacao/*")

            strListItemKey = flMontarChaveItemListview(objDomNode)

            If Not fgExisteItemLvw(Me.lvwMensagem, strListItemKey) Then
                dblValorOperacao = flValorOperacoes(strListItemKey)

                If (dblValorOperacao < 0 And optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value) Or _
                   (dblValorOperacao >= 0 And optNaturezaMovimento(enumNaturezaMovimento.Recebimento).value) Then

                    With lvwMensagem.ListItems.Add(, strListItemKey)

                        .Text = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                        If objDomNode.selectSingleNode("DS_PARP_CAMR").Text <> vbNullString Then
                            .Text = .Text & " - " & objDomNode.selectSingleNode("DS_PARP_CAMR").Text
                        End If

                        If PerfilAcesso = AdmArea Then
                            .SubItems(COL_MSG_TP_PGTO_LDL) = flObterDescricaoTipoPagamento(objDomNode.selectSingleNode("TP_PGTO_LDL").Text)
                        Else
                            .SubItems(COL_MSG_TP_BKOF) = objDomNode.selectSingleNode("DE_BKOF").Text
                        End If

                        .SubItems(COL_MSG_VA_OPER_ATIV) = fgVlrXml_To_Interface(dblValorOperacao)
                        .SubItems(COL_MSG_VA_FINC) = " "

                    End With
                End If
            End If
        Next
    End If

    Call fgClassificarListview(Me.lvwMensagem, lngIndexClassifListMesg, True)
    
    Set xmlRetLeitura = Nothing
    Exit Sub

ErrorHandler:
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaNetOperacoes", 0)

End Sub

'Configurar os botões da tela conforme o perfil do usuário liberando ou não utilização

Private Sub flConfigurarBotoesPorPerfil(pstrPerfil As String)

On Error GoTo ErrorHandler

    With tlbComandos

        .Buttons("concordancia").Visible = (PerfilAcesso = enumPerfilAcesso.AdmArea)
        .Buttons("pagamento").Visible = (PerfilAcesso = enumPerfilAcesso.AdmGeral)
        .Buttons("pagamentocontingencia").Visible = .Buttons("pagamento").Visible
        .Buttons("regularizacao").Visible = .Buttons("pagamento").Visible
        .Refresh

    End With

Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flConfigurarBotoesPorPerfil", 0

End Sub

'Inicializa controles de tela e variáveis
Private Sub flInicializarFormulario()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant
    
On Error GoTo ErrorHandler

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")

    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializarFormulario")
    End If

    Call fgCarregarCombos(Me.cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR")

    Set objMIU = Nothing
    
Exit Sub
ErrorHandler:
    
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flInicializarFormulario", 0

End Sub

'Formata as colunas da lista de mensagens
Private Sub flInicializarLvwMensagem()

On Error GoTo ErrorHandler

    With Me.lvwMensagem.ColumnHeaders
        .Clear
        .Add , , "Participante Negociação", 3000
        
        If PerfilAcesso = AdmArea Then
            .Add , , "Tipo Evento", 2800
        Else
            .Add , , "Área", 2800
        End If
        
        .Add , , "Valor Sistema Origem", 2000, lvwColumnRight
        .Add , , "Valor Câmara", 2000, lvwColumnRight
        .Add , , "Diferença", 2000, lvwColumnRight
        .Add , , "Status Mensagem", 2000
    End With

Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarLvwMensagem", 0

End Sub

'Formata as colunas da lista de operações
Private Sub flInicializarLvwOperacao()

On Error GoTo ErrorHandler

    With Me.lvwOperacao.ColumnHeaders
        .Clear
        .Add , , "Veículo Legal", 3000
        .Add , , "Código Operação", 2000
        .Add , , "D/C", 800
        .Add , , "Valor Sistema Origem", 2000, lvwColumnRight
        .Add , , "Status Operações", 4000
    End With

Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarLvwOperacao", 0

End Sub

'Apaga o conteúdo das listas de mensagens e operações
Private Sub flLimparListas()
    
    Me.lvwMensagem.ListItems.Clear
    Me.lvwOperacao.ListItems.Clear

End Sub

'Monta o conteúdo que será utilizado com a propriedade 'Key' dos itens do ListView
Private Function flMontarChaveItemListview(ByVal objDomNode As MSXML2.IXMLDOMNode)

Dim strListItemKey                          As String

On Error GoTo ErrorHandler

    strListItemKey = "|" & objDomNode.selectSingleNode("CO_PARP_CAMR").Text

    If PerfilAcesso = AdmArea Then
        strListItemKey = strListItemKey & _
                         "|" & objDomNode.selectSingleNode("TP_PGTO_LDL").Text
    Else
        strListItemKey = strListItemKey & _
                         "|" & objDomNode.selectSingleNode("TP_BKOF").Text
    End If

    flMontarChaveItemListview = strListItemKey

Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMontarChaveItemListview", 0

End Function

'Monta uma expressão XPath para seleção do conteúdo de um documento XML
Private Function flMontarCondicaoNavegacaoXMLOperacoes(ByVal strItemKey As String)

Dim strCondicao                             As String

On Error GoTo ErrorHandler

    strCondicao = "Repeat_DetalheOperacao/Grupo_DetalheOperacao[CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_MSG_CO_PARP_CAMR) & "' "

    If PerfilAcesso = AdmArea Then
        strCondicao = strCondicao & _
                    " and TP_PGTO_LDL='" & Split(strItemKey, "|")(KEY_MSG_TP_PGTO_LDL) & "' "
    Else
        strCondicao = strCondicao & _
                    " and TP_BKOF='" & Split(strItemKey, "|")(KEY_MSG_TP_BKOF) & "' "
    End If

    strCondicao = strCondicao & "]"
    flMontarCondicaoNavegacaoXMLOperacoes = strCondicao

Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMontarCondicaoNavegacaoXMLOperacoes", 0

End Function

'Monta uma expressão XPath para a somatória dos valores de operações
Private Function flMontarExpressaoCalculoNetOperacoes(ByVal strItemKey As String)

Dim strDebito                               As String
Dim strCredito                              As String

On Error GoTo ErrorHandler
    
    strDebito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                     " and ../CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_MSG_CO_PARP_CAMR) & "' "

    strCredito = "sum(//VA_OPER_ATIV_VLRXML[../CAMPO_IN_OPER_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                     " and ../CO_PARP_CAMR='" & Split(strItemKey, "|")(KEY_MSG_CO_PARP_CAMR) & "' "

    If PerfilAcesso = AdmArea Then
        strDebito = strDebito & _
                    " and ../TP_PGTO_LDL='" & Split(strItemKey, "|")(KEY_MSG_TP_PGTO_LDL) & "' "

        strCredito = strCredito & _
                    " and ../TP_PGTO_LDL='" & Split(strItemKey, "|")(KEY_MSG_TP_PGTO_LDL) & "' "
    Else
        strDebito = strDebito & _
                    " and ../TP_BKOF='" & Split(strItemKey, "|")(KEY_MSG_TP_BKOF) & "' "

        strCredito = strCredito & _
                    " and ../TP_BKOF='" & Split(strItemKey, "|")(KEY_MSG_TP_BKOF) & "' "
    End If

    strDebito = strDebito & "]) - "
    strCredito = strCredito & "]) "

    flMontarExpressaoCalculoNetOperacoes = strDebito & strCredito

Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMontarExpressaoCalculoNetOperacoes", 0

End Function

'Monta o XML com os dados de filtro para seleção de operações
Private Function flMontarXMLFiltroPesquisa(ByVal intTipoPesquisa As enumTipoPesquisa) As String

Dim xmlFiltros                              As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")

    Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlFiltros, "Grupo_BancoLiquidante", "BancoLiquidante", fgObterCodigoCombo(Me.cboEmpresa.Text))

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.CLBCAcoes)
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(fgDataHoraServidor(DataAux))))

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_SegregaBackOffice", "")
    Call fgAppendNode(xmlFiltros, "Grupo_SegregaBackOffice", "Segrega", "False")
    
    If intTipoPesquisa = Operacao Then
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
        If PerfilAcesso = AdmArea Then
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackofficeAutomatico)
        Else
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackofficeAutomatico)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaAdmArea)
        End If

        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LayoutEntrada", "")
        Call fgAppendNode(xmlFiltros, "Grupo_LayoutEntrada", "LayoutEntrada", "124")
    Else
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
        Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", cboMensagem.Text)
    
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_ControleRepeticao", "")
        Call fgAppendNode(xmlFiltros, "Grupo_ControleRepeticao", "ControleRepeticao", "> 1")
    
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
        If PerfilAcesso = AdmArea Then
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
        Else
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.ConcordanciaAdmArea)
        End If
    End If

    flMontarXMLFiltroPesquisa = xmlFiltros.xml

Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLFiltroPesquisa", 0

End Function

'Monta XML com as chaves das operações que serão processadas
Private Function flMontarXMLProcessamento() As String

Dim objListItem                             As ListItem
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlProcessamento                        As MSXML2.DOMDocument40
Dim xmlItemEnvioMsg                         As MSXML2.DOMDocument40
Dim xmlItemOperacao                         As MSXML2.DOMDocument40
Dim intIgnoraGradeHorario                   As Integer
Dim strItemKey                              As String

On Error GoTo ErrorHandler

    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlProcessamento, "", "Repeat_Processamento", "")

    For Each objListItem In Me.lvwMensagem.ListItems
        With objListItem
            If .Checked Then

                Set xmlItemEnvioMsg = CreateObject("MSXML2.DOMDocument.4.0")

                Call fgAppendNode(xmlItemEnvioMsg, "", "Grupo_EnvioMensagem", "")
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "CO_EMPR", _
                                                   fgObterCodigoCombo(cboEmpresa.Text))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "CO_MESG_SPB", _
                                                   cboMensagem.Text)
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "CO_PARP_CAMR", _
                                                   Split(.Key, "|")(KEY_MSG_CO_PARP_CAMR))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "NU_CTRL_IF", _
                                                   Split(.Tag, "|")(TAG_MSG_NU_CTRL_IF))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "DH_REGT_MESG_SPB", _
                                                   Split(.Tag, "|")(TAG_MSG_DH_REGT_MESG_SPB))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "NU_SEQU_CNTR_REPE", _
                                                   Split(.Tag, "|")(TAG_MSG_NU_SEQU_CNTR_REPE))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "DH_ULTI_ATLZ", _
                                                   Split(.Tag, "|")(TAG_MSG_DH_ULTI_ATLZ))
                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "CO_ULTI_SITU_PROC", _
                                                   Split(.Tag, "|")(TAG_MSG_CO_ULTI_SITU_PROC))

                'intIgnoraGradeHorario = IIf(.ForeColor = vbRed, 1, 0)
                intIgnoraGradeHorario = 1

                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "IgnoraGradeHorario", _
                                                   intIgnoraGradeHorario)

                Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                                   "Repeat_Operacao", _
                                                   "")

                strItemKey = objListItem.Key

                For Each objDomNode In xmlOperacoes.selectNodes(flMontarCondicaoNavegacaoXMLOperacoes(strItemKey))
                    Set xmlItemOperacao = CreateObject("MSXML2.DOMDocument.4.0")

                    Call fgAppendNode(xmlItemOperacao, "", "Grupo_Operacao", "")
                    Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                                       "NU_SEQU_OPER_ATIV", _
                                                       objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                    Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                                       "DH_ULTI_ATLZ", _
                                                       objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)

                    Call fgAppendXML(xmlItemEnvioMsg, "Repeat_Operacao", xmlItemOperacao.xml)

                    Set xmlItemOperacao = Nothing
                Next

                Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemEnvioMsg.xml)
                Set xmlItemEnvioMsg = Nothing
            End If
        End With
    Next

    If xmlProcessamento.selectNodes("Repeat_Processamento/*").length = 0 Then
        flMontarXMLProcessamento = vbNullString
    Else
        flMontarXMLProcessamento = xmlProcessamento.xml
    End If

    Set xmlProcessamento = Nothing

Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLProcessamento", 0

End Function

'Exibe o resultado da última operação executada
Private Sub flMostrarResultado(ByVal pstrResultadoOperacao As String)

On Error GoTo ErrorHandler

    With frmResultOperacaoLote
        .strDescricaoOperacao = " liquidados "
        .Resultado = pstrResultadoOperacao
        .Show vbModal
    End With

Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flMostrarResultado", 0

End Sub

'Obter no Mapa de Navegação a descrição de um TpPagtoLDL

Function flObterDescricaoTipoPagamento(ByVal strTipoPagto As String)

Dim strRet                                  As String

On Error GoTo ErrorHandler
    
    strRet = xmlMapaNavegacao.selectSingleNode("//Grupo_TipoPagamento/DE_DOMI[../CO_DOMI=" & strTipoPagto & "]").Text
    flObterDescricaoTipoPagamento = strRet

Exit Function
ErrorHandler:
    
    If Err.Number = 91 Then
        flObterDescricaoTipoPagamento = "TpPgtoLDL inesperado (" & strTipoPagto & ")"
    Else
        fgRaiseError App.EXEName, TypeName(Me), "flObterDescricaoTipoPagamento", 0
    End If

End Function

'Monta a tela com os dados do filtro selecionado
Private Sub flPesquisar()

Dim strDocFiltros                           As String

On Error GoTo ErrorHandler

    Call flLimparListas

    If Me.cboEmpresa.ListIndex = -1 Or Me.cboEmpresa.Text = vbNullString Then
        frmMural.Display = "Selecione a Empresa."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        cboEmpresa.SetFocus
        Exit Sub
    End If

    If Me.cboMensagem.ListIndex = -1 Or Me.cboMensagem.Text = vbNullString Then
        frmMural.Display = "Selecione o Tipo de Mensagem."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        cboEmpresa.SetFocus
        Exit Sub
    End If

    fgCursor True

    strDocFiltros = flMontarXMLFiltroPesquisa(Operacao)
    Call flCarregarListaNetOperacoes(strDocFiltros)

    If lvwMensagem.ListItems.Count > 0 Then
        lvwMensagem.ListItems(1).Selected = True
        Call lvwMensagem_ItemClick(lvwMensagem.ListItems(1))
    End If

    strDocFiltros = flMontarXMLFiltroPesquisa(MENSAGEM)
    Call flCarregarListaNetMensagens(strDocFiltros)
        
    Call flCalcularDiferencasListView

    fgCursor

Exit Sub
ErrorHandler:
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flPesquisar", 0)

End Sub

'Enviar itens de mensagem e operações para liquidação
Private Function flProcessar() As String

#If EnableSoap = 1 Then
    Dim objOperacaoMensagem     As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacaoMensagem     As A8MIU.clsOperacaoMensagem
#End If

Dim strXMLRetorno               As String
Dim strXMLProc                  As String
Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant

On Error GoTo ErrorHandler

    strXMLProc = flMontarXMLProcessamento

    If strXMLProc <> vbNullString Then
        fgCursor True
        Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
        strXMLRetorno = objOperacaoMensagem.ProcessarLoteLiquidacaoEventosCBLC(intAcaoProcessamento, strXMLProc, vntCodErro, vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If

        Set objOperacaoMensagem = Nothing
        fgCursor
    End If

    If strXMLRetorno <> vbNullString Then
        Set xmlRetornoErro = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlRetornoErro.loadXML(strXMLRetorno)
    Else
        Set xmlRetornoErro = Nothing
    End If

    flProcessar = strXMLRetorno

Exit Function
ErrorHandler:
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flProcessar", Me.Caption

End Function

'Valida a seleção dos itens na tela, para posterior processamento
Private Function flValidarItensProcessamento(ByVal intAcao As enumAcaoConciliacao) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem
    
    For Each objListItem In Me.lvwMensagem.ListItems
        With objListItem
            
            If Trim$(.SubItems(COL_MSG_VA_FINC)) = vbNullString Then
                flValidarItensProcessamento = "Valor da Mensagem não encontrado em um ou mais itens. Solicitação não permitida."
                Exit Function
            End If
            
            Select Case intAcaoProcessamento
                Case enumAcaoConciliacao.AdmGeralPagamentoContingencia
                    If Val(Split(.Tag, "|")(TAG_MSG_CO_ULTI_SITU_PROC)) <> enumStatusMensagem.AConciliar Then
                        flValidarItensProcessamento = "Solicitação não permitida para Mensagens com Status " & .SubItems(COL_MSG_CO_ULTI_SITU_PROC) & "."
                        Exit Function
                    End If
        
                    If .ListSubItems(COL_MSG_DIFERENCA).ForeColor <> vbRed Then
                        flValidarItensProcessamento = "Valores de um ou mais itens batidos. Pagamento em Contingência não permitida."
                        Exit Function
                    End If
            
                Case enumAcaoConciliacao.AdmGeralPagamento
                    If Val(Split(.Tag, "|")(TAG_MSG_CO_ULTI_SITU_PROC)) <> enumStatusMensagem.ConcordanciaAdmArea Then
                        flValidarItensProcessamento = "Solicitação não permitida para Mensagens com Status " & .SubItems(COL_MSG_CO_ULTI_SITU_PROC) & "."
                        Exit Function
                    End If
        
                    If .ListSubItems(COL_MSG_DIFERENCA).ForeColor = vbRed Then
                        flValidarItensProcessamento = "Valores de um ou mais itens não batidos. Solicitação não permitida."
                        Exit Function
                    End If
                
                Case Else
                    If .ListSubItems(COL_MSG_DIFERENCA).ForeColor = vbRed Then
                        flValidarItensProcessamento = "Valores de um ou mais itens não batidos. Solicitação não permitida."
                        Exit Function
                    End If
                
            End Select
        End With
    Next

End Function

'Calcula o valor da operações
Private Function flValorOperacoes(ByVal strItemKey As String)

Dim strExpression                   As String
Dim vntValor                        As Variant

    vntValor = 0

    strExpression = flMontarExpressaoCalculoNetOperacoes(strItemKey)
    vntValor = vntValor + Val(fgFuncaoXPath(xmlOperacoes, strExpression))

    flValorOperacoes = vntValor

End Function

'Configura o perfil de acesso do usuário
Property Get PerfilAcesso() As enumPerfilAcesso
    PerfilAcesso = lngPerfil
End Property

'Configura o perfil de acesso do usuário
Property Let PerfilAcesso(pPerfil As enumPerfilAcesso)
    lngPerfil = pPerfil

    Select Case pPerfil
        Case enumPerfilAcesso.AdmArea
            With Me
                .Caption = "Liberação - Liquidação Eventos CBLC (Administrador de Área)"
                
                .cboMensagem.Visible = True
                .lblMensagem.Visible = True
                .lblPagamento.Visible = True
                .lblRecebimento.Visible = True
                .optNaturezaMovimento(enumNaturezaMovimento.Pagamento).Visible = True
                .optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value = False
                .optNaturezaMovimento(enumNaturezaMovimento.Recebimento).Visible = True
                .optNaturezaMovimento(enumNaturezaMovimento.Recebimento).value = False
                .lvwOperacao.Visible = True
                .imgDummyH.Visible = True
                .cboMensagem.ListIndex = -1
            
            End With
            
        Case Else
            With Me
                .Caption = "Liberação - Liquidação Eventos CBLC (Administrador Geral)"
                
                .cboMensagem.Visible = False
                .lblMensagem.Visible = False
                .lblPagamento.Visible = False
                .lblRecebimento.Visible = False
                .optNaturezaMovimento(enumNaturezaMovimento.Pagamento).Visible = False
                .optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value = True
                .optNaturezaMovimento(enumNaturezaMovimento.Recebimento).Visible = False
                .lvwOperacao.Visible = False
                .imgDummyH.Visible = False
                .cboMensagem.ListIndex = 0
            
            End With
            
    End Select

    Call Form_Resize
    DoEvents
    
    Call flConfigurarBotoesPorPerfil(PerfilAcesso)
    Call flLimparListas
    Call flInicializarLvwMensagem

    If cboEmpresa.Text <> vbNullString And cboMensagem.Text <> vbNullString Then
        Call flPesquisar
    End If

End Property

Private Sub cboEmpresa_Click()

On Error GoTo ErrorHandler
    
    If cboEmpresa.Text <> vbNullString And cboMensagem.Text <> vbNullString Then
        Call flPesquisar
    End If
    
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - cboEmpresa_Click", Me.Caption

End Sub

Private Sub cboMensagem_Click()

On Error GoTo ErrorHandler
    
    If cboMensagem.ListIndex = 0 Then
        optNaturezaMovimento(enumNaturezaMovimento.Pagamento).value = True
    ElseIf cboMensagem.ListIndex = 1 Then
        optNaturezaMovimento(enumNaturezaMovimento.Recebimento).value = True
    End If
    
    If cboEmpresa.Text <> vbNullString And cboMensagem.Text <> vbNullString Then
        Call flPesquisar
    End If
    
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - cboMensagem_Click", Me.Caption

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler

    If KeyCode = vbKeyF5 Then
        fgCursor True
        Call tlbComandos_ButtonClick(tlbComandos.Buttons("refresh"))
        fgCursor
    End If

Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_KeyDown", Me.Caption

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents

    fgCursor True
    Call flInicializarFormulario
    Call flInicializarLvwMensagem
    Call flInicializarLvwOperacao
    fgCursor

    Set xmlOperacoes = CreateObject("MSXML2.DOMDocument.4.0")

Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_Load", Me.Caption

End Sub

Private Sub Form_Resize()

On Error Resume Next

    With Me
        If .imgDummyH.Top < 2000 Then
            .imgDummyH.Top = 2000
        End If
        
        If .imgDummyH.Top > (.Height - 1500) And (.Height - 1500) > 0 Then
            .imgDummyH.Top = .Height - 1500
        End If

        .imgDummyH.Left = 0
        .imgDummyH.Width = .Width

        .lvwMensagem.Top = .cboEmpresa.Top + .cboEmpresa.Height + 120
        .lvwMensagem.Left = .cboEmpresa.Left
        .lvwMensagem.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top - .lvwMensagem.Top, .Height - .lvwMensagem.Top - 720)
        .lvwMensagem.Width = .Width - 240

        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Left = .cboEmpresa.Left
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 720
        .lvwOperacao.Width = .Width - 240
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlRetornoErro = Nothing
    Set xmlOperacoes = Nothing
    Set frmConciliacaoEventos = Nothing
End Sub

Private Sub imgDummyH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnDummyH = True
End Sub

Private Sub imgDummyH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not blnDummyH Or Button = vbRightButton Then
        Exit Sub
    End If

    Me.imgDummyH.Top = Y + imgDummyH.Top

    On Error Resume Next

    With Me
        If .imgDummyH.Top < 2000 Then
            .imgDummyH.Top = 2000
        End If
        
        If .imgDummyH.Top > (.Height - 3000) And (.Height - 3000) > 0 Then
            .imgDummyH.Top = .Height - 3000
        End If

        .lvwMensagem.Height = .imgDummyH.Top - .lvwMensagem.Top
        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 720
    End With

    On Error GoTo 0

End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnDummyH = False
End Sub

Private Sub lvwMensagem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lvwMensagem, ColumnHeader.Index)
    lngIndexClassifListMesg = ColumnHeader.Index
    
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_ColumnClick", Me.Caption

End Sub

Private Sub lvwMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    Call flCarregarListaDetalheOperacoes
    
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwMensagem_ItemClick", Me.Caption

End Sub

Private Sub lvwOperacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrorHandler

    Call fgClassificarListview(Me.lvwOperacao, ColumnHeader.Index)
    lngIndexClassifListOper = ColumnHeader.Index

Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwOperacao_ColumnClick", Me.Caption

End Sub

Private Sub lvwOperacao_DblClick()

On Error GoTo ErrorHandler

    If Not lvwOperacao.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            .SequenciaOperacao = Split(lvwOperacao.SelectedItem.Key, "|")(KEY_OP_NU_SEQU_OPER_ATIV)
            .Show vbModal
        End With
    End If

Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwOperacao_DblClick", Me.Caption

End Sub

Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultadoOperacao                    As String
Dim strValidaProcessamento                  As String

On Error GoTo ErrorHandler

    Button.Enabled = False: DoEvents
    intAcaoProcessamento = 0

    Select Case Button.Key
        Case "refresh"
            Call flPesquisar
            
        Case "concordancia"
            intAcaoProcessamento = enumAcaoConciliacao.AdmAreaLiberar
            
        Case "pagamento"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamento
            
        Case "pagamentocontingencia"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralPagamentoContingencia
            
        Case "regularizacao"
            intAcaoProcessamento = enumAcaoConciliacao.AdmGeralRegularizar
            
        Case gstrSair
            Unload Me
        
    End Select

    If intAcaoProcessamento <> 0 Then
        strValidaProcessamento = flValidarItensProcessamento(intAcaoProcessamento)
        If strValidaProcessamento <> vbNullString Then
            frmMural.Display = strValidaProcessamento
            frmMural.IconeExibicao = IconExclamation
            frmMural.Show vbModal
            GoTo ExitSub
        End If

        If MsgBox("Confirma o processamento do(s) item(s) ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            GoTo ExitSub
        End If

        strResultadoOperacao = flProcessar
        
        If strResultadoOperacao <> vbNullString Then
            Call flMostrarResultado(strResultadoOperacao)
            Call flPesquisar
        End If
    End If

ExitSub:
    fgCursor
    Button.Enabled = True
    Exit Sub

ErrorHandler:
    Button.Enabled = True
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick", Me.Caption

End Sub
