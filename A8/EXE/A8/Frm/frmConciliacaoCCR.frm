VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConciliacaoCCR 
   Caption         =   "Resumo Diário - CCR"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame fraTotal 
      Height          =   795
      Left            =   120
      TabIndex        =   8
      Top             =   930
      Width           =   9855
      Begin VB.TextBox txtQtdeOper 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   435
         Width           =   1965
      End
      Begin VB.TextBox txtVlrLiquido 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   435
         Width           =   1965
      End
      Begin VB.TextBox txtVlrLimiteDispImport 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4725
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0,00"
         Top             =   435
         Width           =   1965
      End
      Begin VB.TextBox txtVlrLimiteTotalImport 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0,00"
         Top             =   435
         Width           =   1965
      End
      Begin VB.Label Label 
         Caption         =   "Quantidade Operações"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   165
         Width           =   1815
      End
      Begin VB.Label Label 
         Caption         =   "Valor Líquido"
         Height          =   255
         Index           =   1
         Left            =   7200
         TabIndex        =   11
         Top             =   165
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Limite Disponível Importação"
         Height          =   255
         Left            =   4680
         TabIndex        =   10
         Top             =   165
         Width           =   2175
      End
      Begin VB.Label Label 
         Caption         =   "Limite Total Importação"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   9
         Top             =   165
         Width           =   2055
      End
   End
   Begin VB.Frame fraFiltro 
      Height          =   945
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   9825
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   3030
      End
      Begin MSComCtl2.DTPicker dtpData 
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60751873
         CurrentDate     =   38455
      End
      Begin VB.Label lblConciliacao 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Operação"
         Height          =   195
         Left            =   3270
         TabIndex        =   5
         Top             =   240
         Width           =   1095
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
            Picture         =   "frmConciliacaoCCR.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoCCR.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoCCR.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoCCR.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoCCR.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoCCR.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoCCR.frx":0F6C
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoCCR.frx":1286
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConciliacaoCCR.frx":16D8
            Key             =   "agendar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwOperacao 
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   6420
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   1508
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
      Left            =   105
      TabIndex        =   1
      Top             =   3315
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tipo Comércio"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo Operação CCR"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Código Reembolso"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tipo Instrumento CCR"
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
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Existe Operação"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   8145
      Width           =   11880
      _ExtentX        =   20955
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
      Top             =   6120
      Width           =   14040
   End
End
Attribute VB_Name = "frmConciliacaoCCR"
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
Private xmlTipoOperacaoCCR                  As MSXML2.DOMDocument40
Private xmlTipoInstrumentoCCR               As MSXML2.DOMDocument40
Private xmlRetOPER                          As MSXML2.DOMDocument40

'Constantes de Configuração de Colunas de Mensagem
Private Const COL_TpComerc                    As Integer = 0
Private Const COL_TpOpComercExtr              As Integer = 1
Private Const COL_CodReemb                    As Integer = 2
Private Const COL_TpInstntoCCR                As Integer = 3
Private Const COL_DtOp                        As Integer = 4
Private Const COL_VlrOp                       As Integer = 5
Private Const COL_TpRecolht_DevCCR            As Integer = 6
Private Const COL_VlrJuros                    As Integer = 7
Private Const COL_VlrTaxAdm                   As Integer = 8
Private Const COL_NumCtrlCCROr                As Integer = 9
'Private Const COL_VlrLimTotImptc              As Integer = 10
'Private Const COL_VlrLimDispImptc             As Integer = 11
'Private Const COL_VlrLiqd                     As Integer = 12
Private Const COL_ExisteOperacao              As Integer = 10


'Constantes de Configuração de Colunas de Operação
Private Const COL_O_TP_OPER                   As Integer = 0
Private Const COL_O_TP_COMERC                 As Integer = 1
Private Const COL_O_COD_REEB                  As Integer = 2
Private Const COL_O_TP_INST_CCR               As Integer = 3
Private Const COL_O_DT_OPER                   As Integer = 4
Private Const COL_O_VALOR_OPER                As Integer = 5

'Constantes de posicionamento de campos na propriedade Key do item do ListView de Mensagens
Private Const POS_NU_CTRL_IF                As Integer = 1
Private Const POS_DH_REGT_MESG_SPB          As Integer = 2
Private Const POS_NU_SEQU_CNTR_REPE         As Integer = 3
Private Const POS_DH_ULTI_ATLZ              As Integer = 4
Private Const POS_NU_SEQU_OPER_ATIV         As Integer = 5

'Constante para o Mapa de Navegação
Private Const strFuncionalidade             As String = "frmConciliacaoFinanceira"

Private blnDummyH                           As Boolean

Private lngIndexClassifListOper             As Long
Private lngIndexClassifListMesg             As Long

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
    fgCursor
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarLista", 0)

End Sub


Private Sub flCarregarMensagens(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim strRetCCR0007                           As String
Dim xmlRetCCR0007                           As MSXML2.DOMDocument40
Dim xmlCCR0007                              As MSXML2.DOMDocument40

Dim strRetOPER                              As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objNodeAux                              As MSXML2.IXMLDOMNode
Dim objNodeCCR0007                          As MSXML2.IXMLDOMNode

Dim strListItemKey                          As String
Dim strListItemTag                          As String
Dim strValorMensagem                        As String

Dim strFiltroOPER                           As String
Dim strfiltroCCR0007                        As String

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Dim lngCont                                 As Long

On Error GoTo ErrorHandler
    
    vntCodErro = 0
    
    txtVlrLimiteDispImport.Text = "0,00"
    txtVlrLimiteTotalImport.Text = "0,00"
    txtVlrLiquido.Text = "0,00"
    txtQtdeOper.Text = 0
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    strfiltroCCR0007 = flMontarXMLFiltroPesquisaCCR0007()
    strFiltroOPER = flMontarXMLFiltroPesquisaOPER()
    
    strRetCCR0007 = objMensagem.ObterDetalheMensagemCamara(strfiltroCCR0007, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    strRetOPER = objMensagem.ObterDetalheMensagemCamara(strFiltroOPER, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If xmlRetOPER Is Nothing Then
        Set xmlRetOPER = CreateObject("MSXML2.DOMDocument.4.0")
    End If
    
    If strRetOPER <> vbNullString Then
        If Not xmlRetOPER.loadXML(strRetOPER) Then
            'Call xmlRetCCR0007(strRetCCR0007, App.EXEName, TypeName(Me), "flCarregaLista")
        End If
    End If
    
    If strRetCCR0007 <> vbNullString Then
    
        
        Set xmlRetCCR0007 = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlRetCCR0007.loadXML(strRetCCR0007) Then
            Call xmlRetCCR0007(strRetCCR0007, App.EXEName, TypeName(Me), "flCarregaLista")
        End If
        
        lngCont = 0
        
        
        
        For Each objDomNode In xmlRetCCR0007.selectNodes("Repeat_DetalheMensagemCamara/*")
            
            strListItemKey = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                             "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                             "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text & _
                             "|" & objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text
                             

            Set xmlCCR0007 = CreateObject("MSXML2.DOMDocument.4.0")
                
            xmlCCR0007.loadXML objMensagem.ObterXMLMensagem(objDomNode.selectSingleNode("CO_TEXT_XML").Text, vntCodErro, vntMensagemErro)
            
            txtVlrLimiteDispImport.Text = fgVlrXml_To_Interface(xmlCCR0007.selectSingleNode("//VlrLimDispImptc").Text)
            txtVlrLimiteTotalImport.Text = fgVlrXml_To_Interface(xmlCCR0007.selectSingleNode("//VlrLimTotImptc").Text)
            txtVlrLiquido.Text = fgVlrXml_To_Interface(xmlCCR0007.selectSingleNode("//VlrLiqd").Text)
            
            txtQtdeOper = xmlCCR0007.selectNodes("//Repet_CCR0007_ResumDiario/*").length
            
            For Each objNodeCCR0007 In xmlCCR0007.selectNodes("//Repet_CCR0007_ResumDiario/*")
                
                lngCont = lngCont + 1
                
                With lvwMensagem.ListItems.Add(, strListItemKey & "|" & lngCont)
                    .Text = ObterValorTAG(objNodeCCR0007, "TpComerc")
                    .SubItems(COL_TpOpComercExtr) = ObterValorTAG(objNodeCCR0007, "TpOpComercExtr")
                    .SubItems(COL_CodReemb) = ObterValorTAG(objNodeCCR0007, "CodReemb")
                    .SubItems(COL_TpInstntoCCR) = ObterValorTAG(objNodeCCR0007, "TpInstntoCCR")
                    .SubItems(COL_DtOp) = fgDtXML_To_Interface(Replace(ObterValorTAG(objNodeCCR0007, "DtOp"), "-", ""))
                    .SubItems(COL_VlrOp) = fgVlrXml_To_Interface(ObterValorTAG(objNodeCCR0007, "VlrOp"))
                    .SubItems(COL_TpRecolht_DevCCR) = ObterValorTAG(objNodeCCR0007, "TpRecolht_DevCCR")
                    .SubItems(COL_VlrJuros) = fgVlrXml_To_Interface(ObterValorTAG(objNodeCCR0007, "VlrJuros"))
                    .SubItems(COL_VlrTaxAdm) = fgVlrXml_To_Interface(ObterValorTAG(objNodeCCR0007, "VlrTaxAdm"))
                    .SubItems(COL_NumCtrlCCROr) = ObterValorTAG(objNodeCCR0007, "NumCtrlCCROr")
                    
                    '.SubItems(COL_VlrLimTotImptc) = fgVlrXml_To_Interface(ObterValorTAG(xmlCCR0007, "//VlrLimTotImptc"))
                    '.SubItems(COL_VlrLimDispImptc) = fgVlrXml_To_Interface(ObterValorTAG(xmlCCR0007, "//VlrLimDispImptc"))
                    '.SubItems(COL_VlrLiqd) = fgVlrXml_To_Interface(ObterValorTAG(xmlCCR0007, "//VlrLiqd"))
                    
                    
                    Set objNodeAux = xmlRetOPER.selectSingleNode("//*[NU_CTRL_CAMR='" & Trim(objDomNode.selectSingleNode("NU_CTRL_CAMR").Text) & "']")
    
                    If objNodeAux Is Nothing Then
                        .SubItems(COL_ExisteOperacao) = "Não"
                    Else
                        .SubItems(COL_ExisteOperacao) = "Sim"
                    End If
                    
                End With
            Next
            
            Set xmlCCR0007 = Nothing
        
        Next
    End If
    
    Set xmlRetCCR0007 = Nothing
    Set objMensagem = Nothing
    
Exit Sub
ErrorHandler:
        
    Set xmlRetCCR0007 = Nothing
    Set xmlCCR0007 = Nothing
    Set objMensagem = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarMensagens", 0)

End Sub


Private Sub flCarregarOperacoesPorMensagem()

Dim objListItem             As ListItem

Dim objDomNode              As MSXML2.IXMLDOMNode
Dim objNodeAux              As MSXML2.IXMLDOMNode
Dim objDomNodeList          As MSXML2.IXMLDOMNodeList

Dim strListItemKey          As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    
    lvwOperacao.ListItems.Clear
   
    Set objDomNodeList = xmlRetOPER.selectNodes("//*[NU_CTRL_CAMR='" & Trim(lvwMensagem.SelectedItem.ListSubItems(COL_CodReemb)) & "']")
    
    For Each objDomNode In objDomNodeList
        
        strListItemKey = "K" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text

        With lvwOperacao.ListItems.Add(, strListItemKey)
            
            Select Case Trim(objDomNode.selectSingleNode("CO_MESG_SPB").Text)
                Case "CCR0002"
                    .Text = "Emissão Operação - CCR"
                Case "CCR0003"
                    .Text = "Negociação Operação - CCR"
                Case "CCR0004"
                    .Text = "Devolução Recolhimento/Estorno Reembolso-CCR"
            End Select
            
            If Mid(objDomNode.selectSingleNode("NU_ATIV_MERC").Text, 1, 2) = "Ex" Then
                .SubItems(COL_O_TP_COMERC) = "Exportação"
            Else
                .SubItems(COL_O_TP_COMERC) = "Importação"
            End If
            
            Set objNodeAux = xmlTipoInstrumentoCCR.selectSingleNode("Repeat_DominioAtributo/*[CO_DOMI='" & Trim(Mid(objDomNode.selectSingleNode("NU_ATIV_MERC").Text, 3)) & "']")
            
            If objNodeAux Is Nothing Then
                .SubItems(COL_O_TP_INST_CCR) = "Tipo Instrumento Inválido"
            Else
                .SubItems(COL_O_TP_INST_CCR) = objNodeAux.selectSingleNode("DE_DOMI").Text
            End If
            
            .SubItems(COL_O_COD_REEB) = objDomNode.selectSingleNode("NU_CTRL_CAMR").Text
            .SubItems(COL_O_DT_OPER) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER").Text)
            .SubItems(COL_O_VALOR_OPER) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)

        End With
    Next
   

    

Exit Sub
ErrorHandler:
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarOperacoesPorMensagem", 0)

End Sub




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
    Set xmlTipoInstrumentoCCR = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlTipoOperacaoCCR = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlTipoInstrumentoCCR.loadXML(objMensagem.ObterDominioSPB("TpInstntoCCR", vntCodErro, vntMensagemErro))
    Call xmlTipoOperacaoCCR.loadXML(objMensagem.ObterDominioSPB("TpOpComercExtr", vntCodErro, vntMensagemErro))
    
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

        .lvwMensagem.Top = .fraTotal.Height + 1000
        
        .lvwMensagem.Left = .fraTotal.Left
        .lvwMensagem.Height = IIf(.imgDummyH.Visible, .imgDummyH.Top - .lvwMensagem.Top, .Height - .lvwMensagem.Top - 720)
        .lvwMensagem.Width = .Width - 240

        .lvwOperacao.Top = .imgDummyH.Top + .imgDummyH.Height
        .lvwOperacao.Left = .fraTotal.Left
        .lvwOperacao.Height = .Height - .lvwOperacao.Top - 900
        .lvwOperacao.Width = .Width - 240
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlRetOPER = Nothing
    Set xmlTipoInstrumentoCCR = Nothing
    Set xmlTipoOperacaoCCR = Nothing
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

Private Sub lvwOperacao_DblClick()
    
On Error GoTo ErrorHandler

    If Not lvwOperacao.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            .SequenciaOperacao = Mid(lvwOperacao.SelectedItem.Key, 2)
            .Show vbModal
        End With
    End If

Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lvwOperacao_DblClick", Me.Caption

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


Private Function flMontarXMLFiltroPesquisaOPER() As String

Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim strAux                                  As String

On Error GoTo ErrorHandler

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")

    Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlFiltros, "Grupo_BancoLiquidante", "BancoLiquidante", fgObterCodigoCombo(Me.cboEmpresa.Text))

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.CCR)
        
    'Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_DataOperacao", "")
    'Call fgAppendNode(xmlFiltros, "Grupo_DataOperacao", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(dtpData.value)))
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
    
    strAux = fgDate_To_DtXML(dtpData.value) & "000000"
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtHrXML_To_Oracle(strAux))
    
    strAux = fgDate_To_DtXML(dtpData.value) & "235959"
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle(strAux))
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_ControleRepeticao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_ControleRepeticao", "ControleRepeticao", "= 1")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Mensagem", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Mensagem", "CodigoMensagem", "CCR0002")
    Call fgAppendNode(xmlFiltros, "Grupo_Mensagem", "CodigoMensagem", "CCR0003")
    Call fgAppendNode(xmlFiltros, "Grupo_Mensagem", "CodigoMensagem", "CCR0004")
    
    
    flMontarXMLFiltroPesquisaOPER = xmlFiltros.xml

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLFiltroPesquisaOPER", 0

End Function

Private Function flMontarXMLFiltroPesquisaCCR0007() As String

Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim strAux                                  As String

On Error GoTo ErrorHandler

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")

    Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlFiltros, "Grupo_BancoLiquidante", "BancoLiquidante", fgObterCodigoCombo(Me.cboEmpresa.Text))

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.CCR)
        
    'Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_DataOperacao", "")
    'Call fgAppendNode(xmlFiltros, "Grupo_DataOperacao", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(dtpData.value)))
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
    
    strAux = fgDate_To_DtXML(dtpData.value) & "000000"
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtHrXML_To_Oracle(strAux))
    
    strAux = fgDate_To_DtXML(dtpData.value) & "235959"
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle(strAux))
    
    
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_SequenciaControleRepeticao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_SequenciaControleRepeticao", "ControleRepeticao", "= 1")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Mensagem", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Mensagem", "CodigoMensagem", "CCR0007")
    
    
    flMontarXMLFiltroPesquisaCCR0007 = xmlFiltros.xml

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLFiltroPesquisaCCR0007", 0

End Function

'Formata as colunas da lista de operações
Private Sub flInicializarlvwDetalhe()

On Error GoTo ErrorHandler
    
    Me.lvwMensagem.ColumnHeaders.Clear
    
    With Me.lvwMensagem.ColumnHeaders
        
        .Clear
        
        .Add , , "Tipo Comércio", 1500
        .Add , , "Tipo Operação Comércio Exterior", 2600
        .Add , , "Código Reembolso", 1500
        .Add , , "Tipo Instrumento CCR", 2500
        .Add , , "Data Operação", 1200, lvwColumnCenter
        .Add , , "Valor Operação", 2000, lvwColumnRight
        .Add , , "Tipo Recolhimento ou Devolução CCR", 2000
        .Add , , "Valor Juros", 2000, lvwColumnRight
        .Add , , "Valor Taxa Administração", 2000, lvwColumnRight
        .Add , , "Número Controle CCR Original", 2500
        '.Add , , "Valor Limite Total Importação", 2000, lvwColumnRight
        '.Add , , "Valor Limite Disponível Importação", 2000, lvwColumnRight
        '.Add , , "Valor Líquido", 2000, lvwColumnRight
        .Add , , "Existe Operação", 1000
        
        
    End With

Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarlvwDetalhe", 0

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
                Set objNodeAux = xmlTipoOperacaoCCR.selectSingleNode("Repeat_DominioAtributo/*[CO_DOMI='" & Trim(pstrValorTAG) & "']")
                
                If objNodeAux Is Nothing Then
                    strRetorno = ""
                Else
                    strRetorno = objNodeAux.selectSingleNode("DE_DOMI").Text
                End If
            
        Case "TpInstntoCCR"
                Set objNodeAux = xmlTipoInstrumentoCCR.selectSingleNode("Repeat_DominioAtributo/*[CO_DOMI='" & Trim(pstrValorTAG) & "']")
                
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

