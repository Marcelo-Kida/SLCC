VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmParametrizacaoSituacaoContingencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contingência - Parametrização Situação"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6660
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Detalhe"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   6375
      Begin VB.ComboBox cboSistemaOrigem 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
      Begin VB.Frame fraMoldura 
         Caption         =   "Interromper Confirmação Automática"
         Height          =   615
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   2895
         Begin VB.OptionButton optConfirmacao 
            Caption         =   "Não"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optConfirmacao 
            Caption         =   "Sim"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sistema Origem"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstPSContigencia 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sistema Origem"
         Object.Width           =   5645
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Confirmação Automática Interrompida"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   120
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":0000
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":09EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":12C6
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":1BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":247A
            Key             =   "Sistema"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":2D54
            Key             =   "AlterarAgendamento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":362E
            Key             =   "Sistema1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":3F08
            Key             =   "SistemaDestino"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":4222
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":453C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":4856
            Key             =   "Regra"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":4B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":4E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":51A4
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":54BE
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":5810
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":5922
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":5C3C
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":5F56
            Key             =   "Evento"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContingencia.frx":6270
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   2640
      TabIndex        =   4
      Top             =   4080
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   582
      ButtonWidth     =   1720
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            ImageKey        =   "Limpar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir"
            ImageKey        =   "Excluir"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            ImageKey        =   "Salvar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmParametrizacaoSituacaoContingencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:35:07
'-------------------------------------------------
'' Objeto responsável pelo envio da solicitação (Administração geral da
'' Parametrização da Situação de Contingência de Sistemas) à camada controladora
'' de caso de uso A8MIU.
''
'' São consideradas especificamente as classes destino:
''      A8MIU.clsMiu
''

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40
Private xmlComboSistemas                    As MSXML2.DOMDocument40

Private strKeyItemSelected                  As String

Private strOperacao                         As String
Private blnPrimeiroActivate                 As Boolean

Private Const strFuncionalidade             As String = "frmParametrizacaoSituacaoContingencia"
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

'Posiciona os items no listview de acordo com a seleção atual.
Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean
    
On Error GoTo ErrorHandler

    If lstPSContigencia.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub
    
    blnEncontrou = False
    For Each objListItem In lstPSContigencia.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lstPSContigencia_ItemClick objListItem
           lstPSContigencia.ListItems(strKeyItemSelected).EnsureVisible
           blnEncontrou = True
           Exit For
        End If
    Next
    Set objListItem = Nothing
    
    If Not blnEncontrou Then
       flLimpaCampos
    End If

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flPosicionaItemListView", 0

End Sub

'' É acionado através no botão 'Salvar' da barra de ferramentas.
''
'' Tem como função, encaminhar a solicitação (Atualização dos dados na tabela) à
'' camada controladora de caso de uso (componente / classe / metodo ) :
''
'' A8MIU.clsMiu.Executar
''
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim strRetorno                              As String
Dim strPropriedades                         As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    strRetorno = flValidarCampos()
    
    If strRetorno <> "" Then
        frmMural.Caption = Me.Caption
        frmMural.Display = strRetorno
        frmMural.Show vbModal
        Exit Sub
    End If
    
    Call fgCursor(True)
    
    Call flInterfaceToXml
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Call objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strOperacao <> gstrOperExcluir Then
        xmlLer.documentElement.selectSingleNode("//@Operacao").Text = "Ler"
        xmlLer.loadXML objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        If strOperacao = "Incluir" Then
           strKeyItemSelected = fgObterCodigoCombo(cboSistemaOrigem.List(cboSistemaOrigem.ListIndex))
        End If
        strOperacao = gstrOperAlterar
        cboSistemaOrigem.Enabled = False
    Else
        flLimpaCampos
    End If
    Set objMIU = Nothing
    
    Call flCarregaListView
    
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
        
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    Set objMIU = Nothing
    
    If strOperacao <> gstrOperExcluir Then
       With xmlLer.documentElement
            strKeyItemSelected = .selectSingleNode("SG_SIST").Text
       End With
    End If
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flSalvar", 0
    

End Sub

'Valida o conteúdo dos campos selecionados na tela
Private Function flValidarCampos() As String
    
On Error GoTo ErrorHandler

    If cboSistemaOrigem.ListIndex < 0 Then
       flValidarCampos = "Informe a Sigla do Sistema."
       cboSistemaOrigem.SetFocus
       Exit Function
    End If
            
    
    If optConfirmacao(0).value = False And optConfirmacao(1).value = False Then
        flValidarCampos = "Informe o Tipo da Confirmação Automática Interrompida."
        optConfirmacao(0).SetFocus
        Exit Function
    End If
    
    flValidarCampos = ""

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flValidarCampos", 0

End Function
'Limpa o conteúdo dos campos
Private Sub flLimpaCampos()
        
On Error GoTo ErrorHandler

    strOperacao = "Incluir"
    
    cboSistemaOrigem.ListIndex = -1
    cboSistemaOrigem.Enabled = True
    
    optConfirmacao(0).value = False
    optConfirmacao(1).value = False
    
    fraMoldura.Enabled = True
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimpaCampos", 0

End Sub
'' Encaminhar a solicitação (Leitura de detalhes da parametrização da sistuação de
'' contingência selecionada) à camada controladora de caso de uso (componente /
'' classe / metodo ) :
''
'' A8MIU.clsMiu.Executar
''
'' O método retornará uma String XML para a camada de interface.
''
Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
        
    cboSistemaOrigem.Enabled = False
        
    With xmlLer.documentElement
        .selectSingleNode("//@Operacao").Text = "Ler"
        .selectSingleNode("//SG_SIST").Text = Split(lstPSContigencia.SelectedItem.Key, "|")(2)
    End With
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    xmlLer.loadXML objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
   
    With xmlLer.documentElement
         fgSearchItemCombo cboSistemaOrigem, , .selectSingleNode("SG_SIST").Text
         If .selectSingleNode("IN_SIST_SITU_CNTG").Text = enumIndicadorSimNao.sim Then
            optConfirmacao(0).value = True
         ElseIf .selectSingleNode("IN_SIST_SITU_CNTG").Text = enumIndicadorSimNao.nao Then
            optConfirmacao(1).value = True
         End If
    End With
        
    Exit Sub
    
ErrorHandler:
    
    Set xmlLer = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
End Sub
'Preenche o documento XML com o conteúdo dos campos da tela
Private Function flInterfaceToXml() As String
    
On Error GoTo ErrorHandler

    With xmlLer.documentElement
    
         .selectSingleNode("//@Operacao").Text = strOperacao
    
         If strOperacao = "Incluir" Then
            .selectSingleNode("//SG_SIST").Text = fgObterCodigoCombo(cboSistemaOrigem.List(cboSistemaOrigem.ListIndex))
            
         ElseIf strOperacao = gstrOperExcluir Then
            Exit Function
         End If
          
         If optConfirmacao(0).value = True Then
            .selectSingleNode("//IN_SIST_SITU_CNTG").Text = enumIndicadorSimNao.sim
         ElseIf optConfirmacao(1).value = True Then
            .selectSingleNode("//IN_SIST_SITU_CNTG").Text = enumIndicadorSimNao.nao
         End If
        
    End With

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0
    
End Function
'' Encaminhar a solicitação (Obtenção das propriedades da tabela do cadastro em
'' questão) à camada controladora de caso de uso (componente / classe / metodo ) :
''
''
'' A8MIU.clsMiu.ObterMapaNavegacao
''
'' O método retornará uma String XML para a camada de interface.
''
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim strMapaNavegacao                        As String
Dim strComboEmpresas                        As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    If fgVerificaJanelaVerificacao Then Exit Sub
    
    Set xmlMapaNavegacao = Nothing
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    strMapaNavegacao = objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, _
                                                 strFuncionalidade, _
                                                 vntCodErro, _
                                                 vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmParametrizacaoSituacaoContingencia", "flInicializar")
    End If
    
    If xmlLer Is Nothing Then
        Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
        xmlLer.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_ParamSituContigencia").xml
    End If
    
    If xmlComboSistemas Is Nothing Then
        Set xmlComboSistemas = CreateObject("MSXML2.DOMDocument.4.0")
        xmlComboSistemas.loadXML xmlMapaNavegacao.xml
    
        xmlComboSistemas.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ParamSituContigencia/@Objeto").Text = "A6A7A8.clsSistema"
        xmlComboSistemas.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ParamSituContigencia/@Operacao").Text = "LerTodos"
        Call fgAppendAttribute(xmlComboSistemas, "Grupo_Propriedades/Grupo_ParamSituContigencia", "IgnoraEmpresa", "S")
        
        strComboEmpresas = objMIU.Executar(xmlComboSistemas.xml, _
                                           vntCodErro, _
                                           vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call xmlComboSistemas.loadXML(strComboEmpresas)
    End If
    
    Set objMIU = Nothing
    
Exit Sub
ErrorHandler:
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'Verifica condição de contingencia do sistema
Public Function ExiteSistemaContingencia() As Boolean

On Error GoTo ErrorHandler

    If xmlMapaNavegacao Is Nothing Then
        flInicializar
    End If

    ExiteSistemaContingencia = flCarregarParametrosContingencia(True) <> vbNullString

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "ExiteSistemaContingencia", 0
End Function

'' Encaminhar a solicitação (Leitura de todas parametrizações de situação de
'' contingência, para o preenchimento do listview) à camada controladora de caso
'' de uso (componente / classe / metodo ) :
''
'' A8MIU.clsMiu.Executar
''
'' O método retornará uma String XML para a camada de interface.
''
Private Function flCarregarParametrosContingencia(ByVal pblnApenasContigencia As Boolean) As String

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim strPropriedades                         As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    If fgVerificaJanelaVerificacao() Then Exit Function
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ParamSituContigencia/@Operacao").Text = "LerTodos"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ParamSituContigencia/IN_SIST_SITU_CNTG").Text = IIf(pblnApenasContigencia, enumIndicadorSimNao.sim, vbNullString)
    strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ParamSituContigencia").xml
    flCarregarParametrosContingencia = objMIU.Executar(strPropriedades, _
                                                       vntCodErro, _
                                                       vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing

Exit Function
ErrorHandler:

    Set objMIU = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

End Function

'Mostra os dados desta tela
Private Sub flCarregaListView()

Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem
Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim strLerTodos                             As String
Dim intCont                                 As Integer

On Error GoTo ErrorHandler

    lstPSContigencia.ListItems.Clear
    
    strLerTodos = flCarregarParametrosContingencia(False)
    
    If strLerTodos = "" Then Exit Sub
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLerTodos.loadXML(strLerTodos)

    For Each xmlDomNode In xmlLerTodos.documentElement.selectNodes("//Repeat_ParamSituContigencia/*")
        intCont = intCont + 1
        
        With xmlDomNode
            'Colocado o contador para impedir o erro << Key is not unique in this collection >>
            'ocorrido no Santander em 03/12/2003, porém não reproduzido no ambiente Regerbanc
            Set objListItem = lstPSContigencia.ListItems.Add(, "|" & intCont & "|" & _
                                                               .selectSingleNode("SG_SIST").Text, _
                                                               .selectSingleNode("SG_SIST").Text & " - " & _
                                                               .selectSingleNode("NO_SIST").Text)
            If .selectSingleNode("IN_SIST_SITU_CNTG").Text = enumIndicadorSimNao.sim Then
               objListItem.SubItems(1) = "Sim"
            ElseIf .selectSingleNode("IN_SIST_SITU_CNTG").Text = enumIndicadorSimNao.nao Then
               objListItem.SubItems(1) = "Não"
            End If
        End With
    Next
    
    Set xmlLerTodos = Nothing

Exit Sub
ErrorHandler:

    Set xmlLerTodos = Nothing
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaListView", 0

End Sub

Private Sub Form_Activate()
    
On Error GoTo ErrorHandler

    If blnPrimeiroActivate Then
    
        DoEvents
        
        Call flLimpaCampos
        
        Call fgCursor(True)
        Call flCarregaListView
        Call fgCarregarCombos(cboSistemaOrigem, xmlComboSistemas, "Sistema", "SG_SIST", "NO_SIST")
        Call fgCursor(False)
        blnPrimeiroActivate = False
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_Activate"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    Set Me.Icon = mdiLQS.Icon
    fgCursor True
    Call flInicializar
    fgCenterMe Me
    blnPrimeiroActivate = True
    fgCursor

Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmParametrizacaoSituacaoContingencia - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmParametrizacaoSituacaoContingencia = Nothing
End Sub

Private Sub lstPSContigencia_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

On Error GoTo ErrorHandler

    lstPSContigencia.Sorted = True
    lstPSContigencia.SortKey = ColumnHeader.Index - 1

    If lstPSContigencia.SortOrder = lvwAscending Then
        lstPSContigencia.SortOrder = lvwDescending
    Else
        lstPSContigencia.SortOrder = lvwAscending
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstPSContigencia_ColumnClick"

End Sub

Private Sub lstPSContigencia_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
            
    Call fgCursor(True)
    
    Call flLimpaCampos
    strOperacao = gstrOperAlterar
    Call flXmlToInterface
    
    strKeyItemSelected = Item.Key
    
    cboSistemaOrigem.Enabled = False
    
    If cboSistemaOrigem.ListIndex >= 0 Then
        tlbCadastro.Buttons(gstrOperExcluir).Enabled = True
    End If
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmParametrizacaoSituacaoContingencia - lstPSContigencia_ItemClick", Me.Caption

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
        Case "Limpar"
        
            Call flLimpaCampos
            If Frame1.Enabled = False And cboSistemaOrigem.Enabled = False Then Exit Sub
            
            cboSistemaOrigem.SetFocus
            
        Case gstrSalvar
            Call flSalvar
            If strOperacao = gstrOperAlterar Then
                flPosicionaItemListView
            End If
        Case gstrOperExcluir
            If MsgBox("Confirma a exclusão do registro ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
               strOperacao = gstrOperExcluir
               Call flSalvar
            End If
        Case gstrSair
            fgCursor False
            Unload Me
            Exit Sub
    End Select
    
    'Atualiza a barra de Status forçando o timer a executar
    glngContaMinutosContingencia = glngTempoContingencia
    mdiLQS.tmrIntervalo.Enabled = False
    mdiLQS.tmrIntervalo.Interval = 1
    mdiLQS.tmrIntervalo.Enabled = True
    
    fgCursor False
    
Exit Sub
ErrorHandler:

    fgCursor False

    mdiLQS.uctlogErros.MostrarErros Err, "frmParametrizacaoSituacaoContingencia - tlbCadastro_ButtonClick", Me.Caption

    Call flCarregaListView
    
    If strOperacao = gstrOperExcluir Then
        flLimpaCampos
    ElseIf strOperacao <> gstrOperExcluir Then
        flPosicionaItemListView
    End If

End Sub

