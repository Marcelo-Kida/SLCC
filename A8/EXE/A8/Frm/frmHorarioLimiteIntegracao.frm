VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHorarioLimiteIntegracao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Segregação Acesso - Horário Limite Integração Sistemas Legados"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCadastro 
      Height          =   3975
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   6885
      Begin VB.Frame fraDetalhes 
         Caption         =   "Detalhe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   210
         TabIndex        =   6
         Top             =   2070
         Width           =   6435
         Begin VB.ComboBox cboSistema 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1200
            Width           =   4275
         End
         Begin VB.ComboBox cboEmpresa 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   480
            Width           =   4275
         End
         Begin MSComCtl2.DTPicker dtpHora 
            Height          =   315
            Left            =   4800
            TabIndex        =   3
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   22806530
            CurrentDate     =   37886
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Horário Limite"
            Height          =   195
            Index           =   2
            Left            =   4800
            TabIndex        =   9
            Top             =   960
            Width           =   960
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Sistema"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   555
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
      End
      Begin MSComctlLib.ListView lstSistemasLegados 
         Height          =   1605
         Left            =   210
         TabIndex        =   0
         Top             =   330
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   2831
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   120
      Top             =   3990
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
            Picture         =   "frmHorarioLimiteIntegracao.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorarioLimiteIntegracao.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorarioLimiteIntegracao.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorarioLimiteIntegracao.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorarioLimiteIntegracao.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorarioLimiteIntegracao.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorarioLimiteIntegracao.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   3060
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
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHorarioLimiteIntegracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:34:24
'-------------------------------------------------
'' Objeto responsável pelo envio da solicitação (Administração geral de horários
'' limite de integração de sistemas legados) à camada controladora de caso de uso
'' A8MIU.
''
'' São consideradas especificamente as classes destino:
''      A8MIU.clsMiu
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40

Private strOperacao                         As String
Private strKeyItemSelected                  As String

Private arrKey()                            As String

Private Const strFuncionalidade             As String = "frmHorarioLimiteIntegracao"
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

'Formata as colunas do ListView
Private Sub flFormataListView()

On Error GoTo ErrorHandler

    lstSistemasLegados.ColumnHeaders.Clear
    lstSistemasLegados.ColumnHeaders.Add , , "Empresa ", 2450, lvwColumnLeft
    lstSistemasLegados.ColumnHeaders.Add , , "Sistema ", 2450, lvwColumnLeft
    lstSistemasLegados.ColumnHeaders.Add , , "Horário Limite", 1200, lvwColumnRight

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flFormataListView", 0

End Sub

'Posiciona os itens do ListView
Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean
 
On Error GoTo ErrorHandler

    If lstSistemasLegados.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub
    
    blnEncontrou = False
    For Each objListItem In lstSistemasLegados.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lstSistemasLegados_ItemClick objListItem
           lstSistemasLegados.ListItems(strKeyItemSelected).EnsureVisible
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

'' É acionado através no botão 'Salvar' da barra de ferramentas.Tem como função,
'' encaminhar a solicitação (Atualização dos dados na tabela) à camada
'' controladora de caso de uso (componente / classe / método ) : A8MIU.clsMiu.
'' Executar
'Alterado   : Adilson G. Damasceno
'Data       : 10/12/2010
'Solicitação: RATS 1028
'Descrição  : Incluido If p/ não validarcampos na Exclusão
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim strRetorno              As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

  If strOperacao <> gstrOperExcluir Then
    strRetorno = flValidarCampos()
  End If
    
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
           strKeyItemSelected = ";" & fgObterCodigoCombo(cboEmpresa.List(cboEmpresa.ListIndex)) & _
                                ";" & fgObterCodigoCombo(cboSistema.List(cboSistema.ListIndex))
        End If
        strOperacao = gstrOperAlterar
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
             strKeyItemSelected = ";" & .selectSingleNode("//CO_EMPR").Text & _
                                  ";" & .selectSingleNode("//SG_SIST").Text
        End With
    End If
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flSalvar", 0

End Sub

'Valida o preenchimento dos campos da tela
Private Function flValidarCampos() As String
    
On Error GoTo ErrorHandler

    If cboEmpresa.ListIndex < 0 Then
        flValidarCampos = "Informe a Empresa do horário limite para Integração Sistemas Legados."
        cboEmpresa.SetFocus
        Exit Function
    End If
    
    If cboSistema.ListIndex < 0 Then
        flValidarCampos = "Informe o Sistema do horário limite para Integração Sistemas Legados."
        cboSistema.SetFocus
        Exit Function
    End If
            
    
    If Not IsNull(dtpHora.value) Then
        If dtpHora.value = gstrDataVazia Then
            flValidarCampos = "Informe o Horário do limite para Integração Sistemas Legados."
            dtpHora.SetFocus
            Exit Function
        End If
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
    
    cboEmpresa.ListIndex = -1
    cboEmpresa.Enabled = True
    
    cboSistema.ListIndex = -1
    cboSistema.Enabled = False
    
    dtpHora.value = fgDataHoraServidor(DataAux)
    dtpHora.Enabled = True
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimpaCampos", 0

End Sub
'' Encaminhar a solicitação (Leitura de detalhes do horário limite de integração
'' para sistemas legados) à camada controladora de caso de uso (componente /
'' classe / metodo ) : A8MIU.clsMiu.Executar
'' O método retornará uma String XML para a camada de interface.
Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    cboEmpresa.Enabled = False
    cboSistema.Enabled = False
        
    With xmlLer.documentElement
        .selectSingleNode("//@Operacao").Text = "Ler"
        .selectSingleNode("//SG_SIST").Text = arrKey(2)
        .selectSingleNode("//CO_EMPR").Text = arrKey(1)
    End With
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    xmlLer.loadXML objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
   
    With xmlLer.documentElement
        
         Call fgSearchItemCombo(cboEmpresa, , .selectSingleNode("CO_EMPR").Text)
         Call fgSearchItemCombo(cboSistema, , .selectSingleNode("SG_SIST").Text)
   
         If Trim(.selectSingleNode("HO_LIMI_ENVI_ITGR").Text) <> gstrDataVazia Then
             dtpHora.value = fgDtHrStr_To_DateTime(.selectSingleNode("HO_LIMI_ENVI_ITGR").Text)
         Else
             dtpHora.value = fgDataHoraServidor(DataAux)
             dtpHora.value = Null
         End If
        
    End With
        
Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flXmlToInterface", 0
    
End Sub
'Preenche o conteúdo do documento XML com o conteúdo dos campos em tela
Private Function flInterfaceToXml() As String

On Error GoTo ErrorHandler

    With xmlLer.documentElement
    
         Me.lstSistemasLegados.SetFocus
         .selectSingleNode("@Operacao").Text = strOperacao
    
         If strOperacao = "Incluir" Then
            .selectSingleNode("SG_SIST").Text = fgObterCodigoCombo(cboSistema.List(cboSistema.ListIndex))
            .selectSingleNode("CO_EMPR").Text = fgObterCodigoCombo(cboEmpresa.List(cboEmpresa.ListIndex))
         ElseIf strOperacao = gstrOperExcluir Then
            Exit Function
         End If
         
         If Not IsNull(dtpHora.value) Then
            .selectSingleNode("HO_LIMI_ENVI_ITGR").Text = fgDateHr_To_DtHrXML(dtpHora.value)
         Else
            .selectSingleNode("HO_LIMI_ENVI_ITGR").Text = ""
         End If
        
    End With
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0
    
End Function
'Carrega o Combo com a lista de sistemas de uma empresa
Private Sub flCarregacboSistema(ByVal plngCO_EMPR As Long)

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim strPropriedades                         As String
Dim strLerTodos                             As String
Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    cboSistema.Clear
    
    Set xmlDomNode = Nothing
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema/@Operacao").Text = "LerTodos"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema/CO_EMPR").Text = plngCO_EMPR
    strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema").xml
    strLerTodos = objMIU.Executar(strPropriedades, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    If strLerTodos = "" Then
        cboSistema.Enabled = False
        frmMural.Caption = Me.Caption
        frmMural.Display = "Não existem sistemas cadastrados para esta empresa."
        frmMural.Show vbModal
        Exit Sub
    End If
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLerTodos.loadXML(strLerTodos)

    For Each xmlDomNode In xmlLerTodos.documentElement.selectNodes("Grupo_Sistema")
        With xmlDomNode
        cboSistema.AddItem .selectSingleNode("SG_SIST").Text & " - " & _
                           .selectSingleNode("NO_SIST").Text
        End With
    Next
    
    If cboSistema.ListCount > 0 Then
       cboSistema.Enabled = True
    End If
    
    Set xmlDomNode = Nothing
    
Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregacboSistema", 0

End Sub
'' Encaminhar a solicitação (Obtenção das propriedades da tabela do cadastro em
'' questão) à camada controladora de caso de uso (componente / classe / metodo ) :
'' A8MIU.clsMiu.ObterMapaNavegacao
'' O método retornará uma String XML para a camada de interface.
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
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmCadastroHorarioLimiteIntegracao", "flInicializar")
    End If
    
    If xmlLer Is Nothing Then
       Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
       xmlLer.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_HorarioLimite").xml
    End If
    
Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0
    
End Sub
'' Encaminhar a solicitação (Leitura de todos os registros da tabela em questão,
'' para o preenchimento do listview) à camada controladora de caso de uso
'' (componente / classe / metodo ) : A8MIU.clsMiu.Executar
'' O método retornará uma String XML para a camada de interface.
Private Sub flCarregaListView()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem
Dim strPropriedades                         As String
Dim strLerTodos                             As String
Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    lstSistemasLegados.ListItems.Clear
    
    Set xmlDomNode = Nothing
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_HorarioLimite/@Operacao").Text = "LerTodos"
    strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_HorarioLimite").xml
    strLerTodos = objMIU.Executar(strPropriedades, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    If strLerTodos = "" Then Exit Sub
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLerTodos.loadXML(strLerTodos)

    For Each xmlDomNode In xmlLerTodos.selectNodes("//Repeat_HorarioLimite/*")
        With xmlDomNode
        
            '";" & .selectSingleNode("HO_LIMI_ENVI_ITGR").Text,
            
            Set objListItem = lstSistemasLegados.ListItems.Add(, ";" & .selectSingleNode("CO_EMPR").Text & _
                                                                 ";" & .selectSingleNode("SG_SIST").Text, _
                                                                 .selectSingleNode("CO_EMPR").Text & " - " & _
                                                                 .selectSingleNode("NO_REDU_EMPR").Text)
                                                                 
            objListItem.SubItems(1) = .selectSingleNode("SG_SIST").Text & " - " & .selectSingleNode("NO_SIST").Text
            
            If CStr(.selectSingleNode("HO_LIMI_ENVI_ITGR").Text) <> CStr(gstrDataVazia) Then
                objListItem.SubItems(2) = Format(Right(.selectSingleNode("HO_LIMI_ENVI_ITGR").Text, 6), "00:00:00")
            End If
        
        End With
    Next

Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlLerTodos = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregaListView", 0

End Sub

Private Sub cboEmpresa_Click()

On Error GoTo ErrorHandler

    fgCursor True

    If cboEmpresa.ListIndex < 0 Then Exit Sub
    Call flCarregacboSistema(CLng(fgObterCodigoCombo(cboEmpresa.List(cboEmpresa.ListIndex))))

    fgCursor

Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cboEmpresa_Click"
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
        
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    flFormataListView
    
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    
    Call flLimpaCampos
    Call fgCursor(True)
    Call flInicializar
    fgCarregarCombos cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_REDU_EMPR"
    Call flCarregaListView
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmCadastroHorarioLimiteIntegracao - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmHorarioLimiteIntegracao = Nothing
End Sub

Private Sub lstSistemasLegados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

    lstSistemasLegados.Sorted = True
    lstSistemasLegados.SortKey = ColumnHeader.Index - 1

    If lstSistemasLegados.SortOrder = lvwAscending Then
        lstSistemasLegados.SortOrder = lvwDescending
    Else
        lstSistemasLegados.SortOrder = lvwAscending
    End If

End Sub

Private Sub lstSistemasLegados_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
            
    Call fgCursor(True)
    
    Call flLimpaCampos
    strOperacao = gstrOperAlterar
    
    arrKey = Split(Item.Key, ";")
    
    strKeyItemSelected = Item.Key
    Call flXmlToInterface
    
    cboEmpresa.Enabled = False
    cboSistema.Enabled = False
    
    If dtpHora.value > gstrDataVazia Then
        tlbCadastro.Buttons(gstrOperExcluir).Enabled = gblnPerfilManutencao
    End If
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmCadastroHorarioLimiteIntegracao - lstSistemasLegados_ItemClick", Me.Caption
    
    flLimpaCampos
    flCarregaListView

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
        Case "Limpar"
            Call flLimpaCampos
            If fraDetalhes.Enabled = True And cboEmpresa.Enabled = True Then
               cboEmpresa.SetFocus
            End If
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
    End Select
    
    fgCursor False
    
Exit Sub
ErrorHandler:

    fgCursor False
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmCadastroHorarioLimiteIntegracao - tlbCadastro_ButtonClick", Me.Caption
    
    Call flCarregaListView
    
    If strOperacao = gstrOperExcluir Then
        flLimpaCampos
    ElseIf strOperacao <> gstrOperExcluir Then
        flPosicionaItemListView
    End If

End Sub

