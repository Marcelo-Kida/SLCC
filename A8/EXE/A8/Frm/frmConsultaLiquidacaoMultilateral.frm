VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmConsultaLiquidacaoMultilateral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Liquidação Multilateral"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   14265
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   0
      Top             =   -60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLiquidacaoMultilateral.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLiquidacaoMultilateral.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLiquidacaoMultilateral.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLiquidacaoMultilateral.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLiquidacaoMultilateral.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLiquidacaoMultilateral.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaLiquidacaoMultilateral.frx":0F6C
            Key             =   "sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMensagem 
      Height          =   4725
      Left            =   0
      TabIndex        =   5
      Top             =   1020
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   8334
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8985
      Begin MSComCtl2.DTPicker dtpData 
         Height          =   315
         Left            =   5400
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57278465
         CurrentDate     =   38455
      End
      Begin VB.ComboBox cboLocalLiquidacao 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   4845
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   5430
         TabIndex        =   4
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Local de Liquidação"
         Height          =   195
         Left            =   210
         TabIndex        =   3
         Top             =   240
         Width           =   1440
      End
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   5745
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   635
      ButtonWidth     =   2381
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
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
            Caption         =   "Sair               "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultaLiquidacaoMultilateral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Consulta Liquidação Multilateral

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40

'Constantes utilizadas no formulário
Private Const COL_EMPRESA                   As Integer = 0
Private Const COL_PARP_CAMR                 As Integer = 1
Private Const COL_VALOR                     As Integer = 2

Private Const strFuncionalidade             As String = "frmConsultaLiquidacaoMultilateral"
'------------------------------------------------------------------------------------------
'Fim declaração constantes

'Carregar dados com valores de Mensagens
Private Sub flCarregarListaMensagens(ByVal pstrFiltro As String)

#If EnableSoap = 1 Then
    Dim objMensagem         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem         As A8MIU.clsMensagem
#End If

Dim strRetLeitura           As String
Dim xmlRetLeitura           As MSXML2.DOMDocument40
Dim objDomNode              As MSXML2.IXMLDOMNode
Dim strListItemKey          As String
Dim dblTotal                As Double
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
        Call xmlRetLeitura.loadXML(strRetLeitura)

        For Each objDomNode In xmlRetLeitura.selectNodes("Repeat_DetalheMensagem/*")
            strListItemKey = "|" & objDomNode.selectSingleNode("NU_CTRL_IF").Text & _
                             "|" & objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text & _
                             "|" & objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text

            With lvwMensagem.ListItems.Add(, strListItemKey)

                .Text = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & _
                            objDomNode.selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
                    
                .SubItems(COL_PARP_CAMR) = objDomNode.selectSingleNode("CO_PARP_CAMR").Text
                If objDomNode.selectSingleNode("DS_PARP_CAMR").Text <> vbNullString Then
                    .SubItems(COL_PARP_CAMR) = .SubItems(COL_PARP_CAMR) & " - " & objDomNode.selectSingleNode("DS_PARP_CAMR").Text
                End If

                If objDomNode.selectSingleNode("CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito Then
                    .SubItems(COL_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text)
                Else
                    .SubItems(COL_VALOR) = "-" & Trim$(fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_FINC").Text))
                End If
                            
                dblTotal = dblTotal + fgVlrXml_To_Decimal(fgVlr_To_Xml(.SubItems(COL_VALOR)))
                
            End With
        
        Next
    
        With lvwMensagem.ListItems.Add
            .Text = "Total"
            .Bold = True
            
            .SubItems(COL_VALOR) = fgVlrXml_To_Interface(fgVlr_To_Xml(dblTotal))
            .ListSubItems(COL_VALOR).Bold = True
        End With
                            
    End If

    Set xmlRetLeitura = Nothing
    
Exit Sub
ErrorHandler:
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCarregarListaMensagens", 0)

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

    Call xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Call fgCarregarCombos(Me.cboLocalLiquidacao, xmlMapaNavegacao, "LocalLiquidacao", "CO_LOCA_LIQU", "DE_LOCA_LIQU")

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
        .Add , , "Empresa", 3500
        .Add , , "Participante Câmara", 3000
        .Add , , "Valor", 2100, lvwColumnRight
    End With

Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarLvwMensagem", 0

End Sub

'Monta o XML com os dados de filtro para seleção de operações
Private Function flMontarXMLFiltroPesquisa() As String

Dim xmlFiltros                              As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")

    Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", fgObterCodigoCombo(Me.cboLocalLiquidacao.Text))
        
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDate_To_DtXML(dtpData.value)))

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_SegregaBackOffice", "")
    Call fgAppendNode(xmlFiltros, "Grupo_SegregaBackOffice", "Segrega", "False")
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
    Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LDL0001")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_ControleRepeticao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_ControleRepeticao", "ControleRepeticao", "> 1")

    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_TipoLDL", "")
    Call fgAppendNode(xmlFiltros, "Grupo_TipoLDL", "TipoLDL", "D")

    flMontarXMLFiltroPesquisa = xmlFiltros.xml

Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLFiltroPesquisa", 0

End Function

'Monta a tela com os dados do filtro selecionado
Private Sub flPesquisar()

Dim strDocFiltros                           As String

On Error GoTo ErrorHandler

    lvwMensagem.ListItems.Clear

    If Me.cboLocalLiquidacao.Text = vbNullString Then
        frmMural.Display = "Selecione o Local de Liquidação."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        cboLocalLiquidacao.SetFocus
        Exit Sub
    End If

    fgCursor True

    strDocFiltros = flMontarXMLFiltroPesquisa
    Call flCarregarListaMensagens(strDocFiltros)
        
    fgCursor

Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flPesquisar", 0)

End Sub

Private Sub cboLocalLiquidacao_Click()

On Error GoTo ErrorHandler
    
    Call flPesquisar
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - cboLocalLiquidacao_Click", Me.Caption

End Sub

Private Sub dtpData_Change()

On Error GoTo ErrorHandler
    
    If cboLocalLiquidacao.Text = vbNullString Then Exit Sub
    Call flPesquisar
    
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
    Call flInicializarLvwMensagem
    Call flInicializarFormulario
    fgCursor

    dtpData.value = fgDataHoraServidor(DataAux)
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmConsultaLiquidacaoMultilateral = Nothing
End Sub

Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    Select Case Button.Key
        Case "refresh"
            Call flPesquisar
        Case gstrSair
            Unload Me
    End Select

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbComandos_ButtonClick", Me.Caption

End Sub
