VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConsultaOperacaoCCR 
   Caption         =   "Consulta Operação CCR"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   12870
   WindowState     =   2  'Maximized
   Begin VB.Frame fraFiltro 
      Height          =   945
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   6465
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   4350
      End
      Begin VB.TextBox txtCodReembolso 
         Height          =   315
         Left            =   6480
         MaxLength       =   20
         TabIndex        =   7
         Top             =   480
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpData 
         Height          =   315
         Left            =   4680
         TabIndex        =   9
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   48758785
         CurrentDate     =   38455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Reembolso"
         Height          =   195
         Left            =   6480
         TabIndex        =   6
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Operação"
         Height          =   195
         Left            =   4680
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblConciliacao 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   0
      Top             =   6990
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
            Picture         =   "frmConsultaOperacaoCCR.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacaoCCR.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacaoCCR.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacaoCCR.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacaoCCR.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacaoCCR.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacaoCCR.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacaoCCR.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacaoCCR.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwOperacao 
      Height          =   3255
      Left            =   60
      TabIndex        =   3
      Top             =   4020
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tipo Operação"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo Comércio"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Código Reembolso"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tipo Instrumento"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Data Operação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Valor Operação"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMensagem 
      Height          =   2685
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   4736
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Numero Controle IF"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Data Hora Mensagem"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   7125
      Width           =   12870
      _ExtentX        =   22701
      _ExtentY        =   582
      ButtonWidth     =   2328
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair            "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Image imgDummyH 
      Height          =   60
      Left            =   60
      MousePointer    =   7  'Size N S
      Top             =   3930
      Width           =   14040
   End
End
Attribute VB_Name = "frmConsultaOperacaoCCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Objeto responsavel pela conciliação financeira de operações bruta e bilateral,
'' através da camada de controle de caso de uso MIU.
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlRetornoErro                      As MSXML2.DOMDocument40

Private xmlTpInstntoCCR                     As MSXML2.DOMDocument40
Private xmlTpOpComercExtr                   As MSXML2.DOMDocument40

Private xmlRetOPER                          As MSXML2.DOMDocument40

'Constantes de Configuração de Colunas de Mensagem
Private Const COL_NU_CTRL_IF                 As Integer = 0
Private Const COL_DH_REGT_MESG_SPB           As Integer = 1

'Constantes de Configuração de Colunas de Operação
Private Const COL_TpComerc                    As Integer = 0
Private Const COL_TpOpComercExtr              As Integer = 1
Private Const COL_CodReemb                    As Integer = 2
Private Const COL_CodSICAP                    As Integer = 3
Private Const COL_CodSICAPCtrapart            As Integer = 4
Private Const COL_PaisCtrapart                As Integer = 5
Private Const COL_PaisOrigemMercdria          As Integer = 6
Private Const COL_TpInstntoCCR                As Integer = 7
Private Const COL_DtOp                        As Integer = 8
Private Const COL_DtVenc_Exprc                As Integer = 9
Private Const COL_VlrOp                       As Integer = 10
Private Const COL_DtHrManut                   As Integer = 11
Private Const COL_CNPJBaseEntManut            As Integer = 12
Private Const COL_TpPessoaImptdr_Exptdr       As Integer = 13
Private Const COL_CNPJ_CPFImptdr_Exptdr       As Integer = 14
Private Const COL_IndrOpSup360Dia             As Integer = 15
Private Const COL_TxtRef                      As Integer = 16
Private Const COL_TpRecolht_DevCCR            As Integer = 17
Private Const COL_VlrSldOpCCR                 As Integer = 18
Private Const COL_VlrJuros                    As Integer = 19
Private Const COL_VlrTaxAdm                   As Integer = 20
Private Const COL_NumCtrlCCROr                As Integer = 21
Private Const COL_SitOpCCR                    As Integer = 22


'Constantes de posicionamento de campos na propriedade Key do item do ListView de Mensagens
Private Const POS_NU_CTRL_IF                As Integer = 1
Private Const POS_DH_REGT_MESG_SPB          As Integer = 2
Private Const POS_NU_SEQU_CNTR_REPE         As Integer = 3
Private Const POS_DH_ULTI_ATLZ              As Integer = 4
Private Const POS_NU_SEQU_OPER_ATIV         As Integer = 5
Private Const POS_CO_TEXT_XML               As Integer = 6


'Constante para o Mapa de Navegação
Private Const strFuncionalidade             As String = "frmConciliacaoFinanceira"

Private blnDummyH                           As Boolean

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

'Formata as colunas da lista de operações
Private Sub flInicializarlvwDetalhe()

On Error GoTo ErrorHandler

    With Me.lvwOperacao.ColumnHeaders
        
        .Clear
        
        .Add , , "Tipo Comércio", 2000
        .Add , , "Tipo Operação Comércio Exterior", 2000
        .Add , , "Código Reembolso", 2000
        .Add , , "Código SICAP", 2000, lvwColumnRight
        .Add , , "Código SICAP Contraparte", 2000, lvwColumnRight
        .Add , , "País Contraparte", 2000, lvwColumnRight
        .Add , , "País Origem Mercadoria", 2000, lvwColumnRight
        .Add , , "Tipo Instrumento CCR", 2000
        .Add , , "Data Operação", 2000, lvwColumnCenter
        .Add , , "Data Vencimento ou Expiração", 2000, lvwColumnCenter
        .Add , , "Valor Operação", 2000, lvwColumnRight
        .Add , , "Data Hora Manutenção", 2000, lvwColumnCenter
        .Add , , "CNPJ Base Entidade Manutenção", 2000, lvwColumnRight
        .Add , , "Tipo Pessoa Importador ou Exportador", 2000
        .Add , , "CNPJ ou CPF Importador ou Exportador", 2000
        .Add , , "Indicador Operação Superior 360 Dias", 2000
        .Add , , "Texto Referência", 2000
        .Add , , "Tipo Recolhimento ou Devolução CCR", 2000
        .Add , , "Valor Saldo Operação CCR", 2000, lvwColumnRight
        .Add , , "Valor Juros", 2000, lvwColumnRight
        .Add , , "Valor Taxa Administração", 2000, lvwColumnRight
        .Add , , "Número Controle CCR Original", 2000
        .Add , , "Situação Operação CCR", 2000, lvwColumnRight
                
        
        
    End With

Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarlvwDetalhe", 0

End Sub

'Controlar a chamada das funcionalidades que irão preencher as informações da Interface
Private Sub flCarregarLista()

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
    
    fgCursor True
    
    Call flCarregarMensagens(strDocFiltros)

    fgCursor
    
Exit Sub
ErrorHandler:
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarLista", 0)

End Sub


Private Sub flCarregarMensagens(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim strRetCCR0001                           As String
Dim xmlRetCCR0001                           As MSXML2.DOMDocument40

Dim strRetOPER                              As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objNodeAux                              As MSXML2.IXMLDOMNode
Dim strListItemKey                          As String
Dim strListItemTag                          As String
Dim strValorMensagem                        As String

Dim strfiltroCCR0001                        As String

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    vntCodErro = 0
    
    strfiltroCCR0001 = flMontarXMLFiltroPesquisaCCR0001()
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    strRetCCR0001 = objMensagem.ObterDetalheMensagemCamara(strfiltroCCR0001, vntCodErro, vntMensagemErro)
    Set objMensagem = Nothing
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    
    If strRetCCR0001 <> vbNullString Then
        Set xmlRetCCR0001 = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlRetCCR0001.loadXML(strRetCCR0001) Then
            Call xmlRetCCR0001(strRetCCR0001, App.EXEName, TypeName(Me), "flCarregaLista")
        End If
        
        For Each objDomNode In xmlRetCCR0001.selectNodes("Repeat_DetalheMensagemCamara/*")
        
            strListItemKey = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                             "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                             "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & _
                             "|" & objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text & _
                             "|" & objDomNode.selectSingleNode("CO_TEXT_XML").Text

            With lvwMensagem.ListItems.Add(, strListItemKey, objDomNode.selectSingleNode("NU_CTRL_IF").Text)
                .SubItems(1) = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_RECB_ENVI_MESG_SPB").Text)
            End With
        
        Next
    End If
    
    Set xmlRetCCR0001 = Nothing

    
Exit Sub
ErrorHandler:
        
    Set xmlRetCCR0001 = Nothing
        
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarMensagens", 0)

End Sub


Private Sub flCarregarOperacoesPorMensagem()


#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
    Dim objMensagem        As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
    Dim objMensagem        As A8MIU.clsMensagem
#End If

Dim objListItem             As ListItem
Dim objDomNode              As MSXML2.IXMLDOMNode
Dim objNodeAux              As MSXML2.IXMLDOMNode
Dim objDomNodeList          As MSXML2.IXMLDOMNodeList
Dim xmlCCR0001R1            As MSXML2.DOMDocument40

Dim strXPaxth               As String

Dim strListItemKey          As String
Dim lngCount                As Long
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    
    lvwOperacao.ListItems.Clear
    lngCount = 0
   
    Set xmlCCR0001R1 = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    xmlCCR0001R1.loadXML objMensagem.ObterXMLMensagem(CLng(Split(lvwMensagem.SelectedItem.Key, "|")(5)), vntCodErro, vntMensagemErro)
    Set objMensagem = Nothing
    
    If Trim(txtCodReembolso) = vbNullString Then
        strXPaxth = "//Repet_CCR0001R1_OpComercExtr/*"
    Else
        strXPaxth = "//Repet_CCR0001R1_OpComercExtr/*[CodReemb='" & Trim(txtCodReembolso) & "']"
    End If
    
    
    For Each objDomNode In xmlCCR0001R1.selectNodes(strXPaxth)
        
        lngCount = lngCount + 1
        
        strListItemKey = "|" & objDomNode.selectSingleNode("TpComerc").Text & "|" & lngCount

        With lvwOperacao.ListItems.Add(, strListItemKey)
            .Text = ObterValorTAG(objDomNode, "TpComerc")
            .SubItems(COL_TpOpComercExtr) = ObterValorTAG(objDomNode, "TpOpComercExtr")
            .SubItems(COL_CodReemb) = ObterValorTAG(objDomNode, "CodReemb")
            .SubItems(COL_CodSICAP) = ObterValorTAG(objDomNode, "CodSICAP")
            .SubItems(COL_CodSICAPCtrapart) = ObterValorTAG(objDomNode, "CodSICAPCtrapart")
            .SubItems(COL_PaisCtrapart) = ObterValorTAG(objDomNode, "PaisCtrapart")
            .SubItems(COL_PaisOrigemMercdria) = ObterValorTAG(objDomNode, "PaisOrigemMercdria")
            .SubItems(COL_TpInstntoCCR) = ObterValorTAG(objDomNode, "TpInstntoCCR")
            .SubItems(COL_DtOp) = fgDtXML_To_Interface(ObterValorTAG(objDomNode, "DtOp"))
            .SubItems(COL_DtVenc_Exprc) = fgDtXML_To_Interface(ObterValorTAG(objDomNode, "DtVenc_Exprc"))
            .SubItems(COL_VlrOp) = fgVlrXml_To_Interface(ObterValorTAG(objDomNode, "VlrOp"))
            .SubItems(COL_DtHrManut) = fgDtHrXML_To_Interface(ObterValorTAG(objDomNode, "DtHrManut"))
            .SubItems(COL_CNPJBaseEntManut) = ObterValorTAG(objDomNode, "CNPJBaseEntManut")
            .SubItems(COL_TpPessoaImptdr_Exptdr) = ObterValorTAG(objDomNode, "TpPessoaImptdr_Exptdr")
            .SubItems(COL_CNPJ_CPFImptdr_Exptdr) = fgFormataCnpj(ObterValorTAG(objDomNode, "CNPJ_CPFImptdr_Exptdr"))
            .SubItems(COL_IndrOpSup360Dia) = ObterValorTAG(objDomNode, "IndrOpSup360Dia")
            .SubItems(COL_TxtRef) = ObterValorTAG(objDomNode, "TxtRef")
            .SubItems(COL_TpRecolht_DevCCR) = ObterValorTAG(objDomNode, "TpRecolht_DevCCR")
            .SubItems(COL_VlrSldOpCCR) = fgVlrXml_To_Interface(ObterValorTAG(objDomNode, "VlrSldOpCCR"))
            .SubItems(COL_VlrJuros) = fgVlrXml_To_Interface(ObterValorTAG(objDomNode, "VlrJuros"))
            .SubItems(COL_VlrTaxAdm) = fgVlrXml_To_Interface(ObterValorTAG(objDomNode, "VlrTaxAdm"))
            .SubItems(COL_NumCtrlCCROr) = ObterValorTAG(objDomNode, "NumCtrlCCROr")
            .SubItems(COL_SitOpCCR) = ObterValorTAG(objDomNode, "SitOpCCR")
        End With
    Next
    
    Set xmlCCR0001R1 = Nothing
    
Exit Sub
ErrorHandler:
    Set xmlCCR0001R1 = Nothing
    Set objMensagem = Nothing
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarOperacoesPorMensagem", 0)

End Sub

Private Function flObterDescricaoDominio(ByVal pstrNomeTAG, pstrValorTAG As String)
    
Dim strRetorno                              As String
Dim objNodeAux                              As MSXML2.IXMLDOMNode
    
On Error GoTo ErrorHandler
            
    Select Case Trim(pstrNomeTAG)
        
        Case "TpComerc"
            
            If pstrValorTAG = "Ex" Then
                strRetorno = "Exportação"
            Else
                strRetorno = "Importação"
            End If
        
        Case "TpOpComercExtr"
                Set objNodeAux = xmlTpOpComercExtr.selectSingleNode("Repeat_DominioAtributo/*[CO_DOMI='" & Trim(pstrValorTAG) & "']")
                
                If objNodeAux Is Nothing Then
                    strRetorno = ""
                Else
                    strRetorno = objNodeAux.selectSingleNode("DE_DOMI").Text
                End If
            
        Case "PaisCtrapart", "PaisOrigemMercdria"
            
            If pstrValorTAG = "1" Then
                strRetorno = "Argentina"
            ElseIf pstrValorTAG = "2" Then
                strRetorno = "Bolívia"
            ElseIf pstrValorTAG = "3" Then
                strRetorno = "Brasil"
            ElseIf pstrValorTAG = "4" Then
                strRetorno = "Colômbia"
            ElseIf pstrValorTAG = "5" Then
                strRetorno = "Chile"
            ElseIf pstrValorTAG = "6" Then
                strRetorno = "Equador"
            ElseIf pstrValorTAG = "7" Then
                strRetorno = "México"
            ElseIf pstrValorTAG = "8" Then
                strRetorno = "Paraguai"
            ElseIf pstrValorTAG = "9" Then
                strRetorno = "Peru"
            ElseIf pstrValorTAG = "10" Then
                strRetorno = "Uruguai"
            ElseIf pstrValorTAG = "11" Then
                strRetorno = "Venezuela"
            ElseIf pstrValorTAG = "12" Then
                strRetorno = "República Dominicana"
            End If
                
        Case "TpInstntoCCR"
                Set objNodeAux = xmlTpInstntoCCR.selectSingleNode("Repeat_DominioAtributo/*[CO_DOMI='" & Trim(pstrValorTAG) & "']")
                
                If objNodeAux Is Nothing Then
                    strRetorno = ""
                Else
                    strRetorno = objNodeAux.selectSingleNode("DE_DOMI").Text
                End If
            
        Case "TpRecolht_DevCCR"
            If pstrValorTAG = "IF" Then
                strRetorno = "Gerado por requisição da IF SISTEMA"
            ElseIf pstrValorTAG = "BC" Then
                strRetorno = "Gerado por iniciativa do BACEN quando receber débito do exterior ou estorno de débito do exterior"
            End If
        Case "SitOpCCR"
            If pstrValorTAG = "1" Then
                strRetorno = "Pendente de registro"
            ElseIf pstrValorTAG = "2" Then
                strRetorno = "Registrada"
            ElseIf pstrValorTAG = "3" Then
                strRetorno = "Pendente de aceite"
            ElseIf pstrValorTAG = "4" Then
                strRetorno = "Rejeitada"
            ElseIf pstrValorTAG = "5" Then
                strRetorno = "Excluída"
            ElseIf pstrValorTAG = "6" Then
                strRetorno = "Reembolsada"
            ElseIf pstrValorTAG = "7" Then
                strRetorno = "Recolhida"
            ElseIf pstrValorTAG = "13" Then
                strRetorno = "Rejeitada com erro"
            End If
        Case Else
            strRetorno = pstrValorTAG
        End Select
        
        flObterDescricaoDominio = strRetorno
        
Exit Function
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - flObterDescricaoDominio", Me.Caption

End Function


Private Function ObterValorTAG(ByVal pxmlNode As MSXML2.IXMLDOMNode, ByVal pstrNomeTAG As String) As String
    
    If pxmlNode Is Nothing Then
        ObterValorTAG = ""
    ElseIf pxmlNode.selectSingleNode(pstrNomeTAG) Is Nothing Then
        ObterValorTAG = ""
    Else
        ObterValorTAG = flObterDescricaoDominio(ByVal pstrNomeTAG, Trim(pxmlNode.selectSingleNode(pstrNomeTAG).Text))
    End If
    
End Function




' Carrega as propriedades necessárias ao formulário, através da
' camada de controle de caso de uso, método A8MIU.clsMIU.ObterMapaNavegacao

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
    Dim objMensagem        As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
    Dim objMensagem        As A8MIU.clsMensagem
#End If

Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler
    
    vntCodErro = 0
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializar")
    End If
    
    Call fgCarregarCombos(Me.cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR")
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    Set xmlTpInstntoCCR = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlTpOpComercExtr = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlTpInstntoCCR.loadXML(objMensagem.ObterDominioSPB("TpInstntoCCR", vntCodErro, vntMensagemErro))
    Call xmlTpOpComercExtr.loadXML(objMensagem.ObterDominioSPB("TpOpComercExtr", vntCodErro, vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMensagem = Nothing
    
    Set objMIU = Nothing
    
Exit Sub
ErrorHandler:
    
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'Limpar os list view de mensagem e de operação
Private Sub flLimparListas()
    
    Me.lvwMensagem.ListItems.Clear
    Me.lvwOperacao.ListItems.Clear

End Sub

Public Sub RedimensionarForm()
    
    Call Form_Resize

End Sub

Private Sub cboEmpresa_Click()
    
On Error GoTo ErrorHandler

    If cboEmpresa.Text <> vbNullString Then
        Call flCarregarLista
    End If
    
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - cboEmpresa_Click", Me.Caption

End Sub




Private Sub dtpData_Change()

On Error GoTo ErrorHandler

    Call flCarregarLista
    
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - dtpData_Change", Me.Caption

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
    Call flInicializar
    Call flInicializarlvwDetalhe
    
    
    dtpData.value = fgDataHoraServidor(DataAux)
    
    fgCursor
    
                
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

        .lvwMensagem.Top = .fraFiltro.Height + 100
        
        .lvwMensagem.Left = .fraFiltro.Left
        .lvwMensagem.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top - .lvwMensagem.Top, .Height - .lvwMensagem.Top - 720)
        .lvwMensagem.Width = .Width - 240

        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Left = .fraFiltro.Left
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 900
        .lvwOperacao.Width = .Width - 240
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlRetOPER = Nothing
    Set xmlTpInstntoCCR = Nothing
    Set xmlTpOpComercExtr = Nothing
    Set xmlRetornoErro = Nothing
    Set frmConciliacaoCCR = Nothing

End Sub

Private Sub imgDummyH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    blnDummyH = True

End Sub

Private Sub imgDummyH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Not blnDummyH Or Button = vbRightButton Then
        Exit Sub
    End If

    Me.imgDummyH.Top = y + imgDummyH.Top

    On Error Resume Next

    With Me
        If .imgDummyH.Top < 2000 Then
            .imgDummyH.Top = 2000
        End If
        If .imgDummyH.Top > (.Height - 1500) And (.Height - 1500) > 0 Then
            .imgDummyH.Top = .Height - 1500
        End If

        .lvwMensagem.Height = .imgDummyH.Top - .lvwMensagem.Top
        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 900
    End With

    On Error GoTo 0

End Sub

Private Sub imgDummyH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
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


Private Sub lvwMensagem_DblClick()
On Error GoTo ErrorHandler

    If Not lvwMensagem.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            .NumeroControleIF = Split(lvwMensagem.SelectedItem.Key, "|")(POS_NU_CTRL_IF)
            .DataRegistroMensagem = fgDtHrStr_To_DateTime(Split(lvwMensagem.SelectedItem.Key, "|")(POS_DH_REGT_MESG_SPB))
            .NumeroSequenciaRepeticao = 1
            .Show vbModal
        End With
    End If

Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwOperacao_DblClick", Me.Caption

End Sub

Private Sub lvwMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    fgCursor True
    Call flCarregarOperacoesPorMensagem
    fgCursor

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


Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResultadoOperacao                    As String
Dim strErro                                 As String

On Error GoTo ErrorHandler
    
    Button.Enabled = False: DoEvents
    
    Select Case Button.Key
        Case "refresh"
            Call flCarregarLista
            
            
        Case gstrSair
            Unload Me
            
    End Select
    
    
ExitSub:
    fgCursor
    Button.Enabled = True
    Exit Sub

ErrorHandler:
    Button.Enabled = True
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick", Me.Caption

End Sub




Private Function flMontarXMLFiltroPesquisaCCR0001() As String

Dim xmlFiltros                              As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")

    Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlFiltros, "Grupo_BancoLiquidante", "BancoLiquidante", fgObterCodigoCombo(Me.cboEmpresa.Text))

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.CCR)
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_DataOperacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_DataOperacao", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(dtpData.value)))
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_ControleRepeticao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_ControleRepeticao", "ControleRepeticao", "> 1")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Mensagem", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Mensagem", "CodigoMensagem", "CCR0001")
    Call fgAppendNode(xmlFiltros, "Grupo_Mensagem", "CodigoMensagem", "CCR0001R1")
    
    
    flMontarXMLFiltroPesquisaCCR0001 = xmlFiltros.xml

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLFiltroPesquisaCCR0001", 0

End Function



