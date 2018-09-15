VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmParamComunicacaoSistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Parâmetros de Comunicação com Sistemas"
   ClientHeight    =   4545
   ClientLeft      =   2655
   ClientTop       =   3885
   ClientWidth     =   10695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   10695
   Begin VB.Frame fraDetalhe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4080
      Left            =   4260
      TabIndex        =   4
      Top             =   -45
      Width           =   6405
      Begin VB.TextBox txtNomeFila 
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   435
         Width           =   6165
      End
      Begin MSComctlLib.ListView lstOcorrencias 
         Height          =   2970
         Left            =   120
         TabIndex        =   2
         Top             =   1005
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   5239
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Nome da Fila Padrão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   210
         Width           =   2145
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ocorrências Tratadas pelo Sistema Chamador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   795
         Width           =   3885
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   7110
      TabIndex        =   3
      Top             =   4140
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   582
      ButtonWidth     =   1535
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "Limpar"
            ImageKey        =   "Limpar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "Excluir"
            ImageKey        =   "Excluir"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "Salvar"
            ImageKey        =   "Salvar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons1 
      Left            =   60
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":0000
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":0112
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":042C
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":077E
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":0890
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":0BAA
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":0EC4
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":11DE
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treSistema 
      Height          =   3930
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   6932
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlIcons"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   915
      Top             =   4095
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
            Picture         =   "frmParamComunicacaoSistema.frx":14F8
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":160A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":1EE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":27BE
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":3098
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":3972
            Key             =   "Sistema"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":424C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":4B26
            Key             =   "Sistema1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":5400
            Key             =   "SistemaDestino"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":571A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":5A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":5D4E
            Key             =   "Regra"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":6068
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":669C
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":69B6
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":6D08
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":6E1A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":7134
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":744E
            Key             =   "Evento"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParamComunicacaoSistema.frx":7768
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmParamComunicacaoSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela configuração da comunicação com os sistema destino (configura nome da fila de destino).
Option Explicit

'Este objeto ObjDomDocument é carregado com as propriedades do objParametroPostagem
' e todas as coleções que este form utiliza
Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlParamComunicaoSistema            As MSXML2.DOMDocument40

Private strOperacao                         As String
Private Const strFuncionalidade             As String = "frmParamComunicacaoSistema"

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

'Limpar campos do formulário.
Private Sub flLimpaCampos()

Dim objListItem                             As ListItem

On Error GoTo ErrorHandler
    
    txtNomeFila.Text = ""
    
    For Each objListItem In lstOcorrencias.ListItems
        objListItem.Checked = False
    Next
        
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
        
    fgRaiseError App.EXEName, "frmParamComunicacaoSistema", "flLimpaCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Salvar configuração de comunicação corrente.
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim strRetorno          As String
Dim strPropriedades     As String
Dim strLer              As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    strRetorno = flValidarCampos()
    
    If strRetorno <> "" Then
        frmMural.Display = strRetorno
        frmMural.Show vbModal
        Exit Sub
    End If
        
    Call fgCursor(True)
    
    Call flInterfaceToXml

    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    Call objMiu.Executar(xmlParamComunicaoSistema.documentElement.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    With xmlParamComunicaoSistema.documentElement
        .selectSingleNode("Grupo_EnderecoFilaMqseries/@Operacao").Text = "Ler"
        .selectSingleNode("Grupo_EnderecoFilaMqseries/SG_SIST_DEST").Text = Mid(treSistema.SelectedItem.Key, 8, 3)
        .selectSingleNode("Grupo_EnderecoFilaMqseries/CO_EMPR_DEST").Text = CLng(Mid(treSistema.SelectedItem.Key, 2, 5))
        strPropriedades = .selectSingleNode("Grupo_EnderecoFilaMqseries").xml
    End With
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    strLer = objMiu.Executar(strPropriedades, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    If strLer <> "" Then
        xmlParamComunicaoSistema.loadXML strLer
        strOperacao = "Alterar"
    Else
        If strOperacao = "Excluir" Then
            flLimpaCampos
            tlbCadastro.Buttons("Excluir").Enabled = False
            strOperacao = "Incluir"
        End If
    End If
    
    Call fgCursor(False)
        
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    
    Exit Sub

ErrorHandler:

    Set objMiu = Nothing
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmParamComunicacaoSistema", "flSalvar", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)
    
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
        
    strOperacao = "Incluir"
        
    fgCenterMe Me
    
    Me.Icon = mdiBUS.Icon
    
    Me.Show
    DoEvents
   
    Call fgCursor(True)
    
    Call flInicializar
    Call flCarregartreSistemas
    
    Call flLimpaCampos
    
    Call flDefinirTamanhoMaximoCampos
    Call flCarregarOcorrencias
    
    fraDetalhe.Enabled = False
    tlbCadastro.Buttons("Salvar").Enabled = False
    tlbCadastro.Buttons("Limpar").Enabled = False
    tlbCadastro.Buttons("Excluir").Enabled = False

    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    mdiBUS.uctLogErros.MostrarErros Err, "frmParamComunicacaoSistema - Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmParamComunicacaoSistema = Nothing
    
End Sub

'Validar os valores informados para o parâmetro de comunicação.
Private Function flValidarCampos() As String
    
Dim objListItem                             As ListItem
    
    If treSistema.SelectedItem.Parent Is Nothing Then
        flValidarCampos = "Selecione um sistema."
        Exit Function
    End If
    
    For Each objListItem In lstOcorrencias.ListItems
        
        If objListItem.Checked Then
            If Trim(txtNomeFila) = "" Then
                flValidarCampos = "Digite o Nome da Fila ."
                txtNomeFila.SetFocus
                Exit Function
            End If
            Exit For
        End If
    Next
    
    flValidarCampos = ""
    
End Function

'Mover valores do formulário para XML para envio ao objeto de negócio.
Private Function flInterfaceToXml() As String
    
Dim objListItem                             As ListItem
Dim xmlPropriedades                         As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlNodeAux                              As MSXML2.IXMLDOMNode
Dim blnEncontrou                            As Boolean

On Error GoTo ErrorHandler
     
    With xmlParamComunicaoSistema.documentElement
        .selectSingleNode("Grupo_EnderecoFilaMqseries/@Operacao").Text = strOperacao
        .selectSingleNode("Grupo_EnderecoFilaMqseries/SG_SIST_DEST").Text = Mid(treSistema.SelectedItem.Key, 8, 3)
        .selectSingleNode("Grupo_EnderecoFilaMqseries/CO_EMPR_DEST").Text = CLng(Mid(treSistema.SelectedItem.Key, 2, 5))
        .selectSingleNode("Grupo_EnderecoFilaMqseries/NO_FILA_MQSE").Text = txtNomeFila.Text
    End With
    
    Set xmlPropriedades = CreateObject("MSXML2.DOMDocument.4.0")
                
    blnEncontrou = False
    'Remover todos os nós e evento_atributo
    For Each xmlNode In xmlParamComunicaoSistema.selectNodes("//Repeat_RespostaOcorrenciaSistema/*")
        blnEncontrou = True
        xmlParamComunicaoSistema.selectSingleNode("//Repeat_RespostaOcorrenciaSistema").removeChild xmlNode
    Next
    
    If Not blnEncontrou Then
       fgAppendNode xmlParamComunicaoSistema, "Grupo_ParametrosSistema", "Repeat_RespostaOcorrenciaSistema", "", ""
    End If
        
    For Each objListItem In lstOcorrencias.ListItems
        If objListItem.Checked Then
            Set xmlNodeAux = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ParametrosSistema/Grupo_RespostaOcorrenciaSistema").cloneNode(True)
            
            
            xmlNodeAux.selectSingleNode("CO_OCOR_MESG").Text = Mid$(objListItem.Key, 5)
            xmlNodeAux.selectSingleNode("SG_SIST").Text = Mid(treSistema.SelectedItem.Key, 8, 3)
            xmlNodeAux.selectSingleNode("CO_EMPR").Text = CLng(Mid(treSistema.SelectedItem.Key, 2, 5))
                        
            Call fgAppendXML(xmlParamComunicaoSistema, "Repeat_RespostaOcorrenciaSistema", xmlNodeAux.xml)
        End If
    Next
    
    Exit Function

ErrorHandler:
    
    Call fgCursor(False)
    
    fgRaiseError App.EXEName, "frmParamComunicacaoSistema", "flInterfaceToXml", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

'Configurar o tamanho máximo para campos de nome de fila de comunicação.
Private Sub flDefinirTamanhoMaximoCampos()

On Error GoTo ErrorHandler
                
    With xmlMapaNavegacao
        txtNomeFila.MaxLength = .selectSingleNode("//Grupo_Propriedades/Grupo_ParametrosSistema/Grupo_EnderecoFilaMqseries/NO_FILA_MQSE/@Tamanho").Text
    End With
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    
    fgRaiseError App.EXEName, "frmParamComunicacaoSistema", "flDefinirTamanhoMaximoCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Carregar os campos do formulário com os valores recebidos da camada de negócio.
Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim xmlLer              As MSXML2.DOMDocument40
Dim strPropriedades     As String
Dim strLer              As String
Dim xmlNode             As MSXML2.IXMLDOMNode
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    With xmlParamComunicaoSistema.documentElement
        .selectSingleNode("Grupo_EnderecoFilaMqseries/@Operacao").Text = "Ler"
        .selectSingleNode("Grupo_EnderecoFilaMqseries/SG_SIST_DEST").Text = Mid(treSistema.SelectedItem.Key, 8, 3)
        .selectSingleNode("Grupo_EnderecoFilaMqseries/CO_EMPR_DEST").Text = CLng(Mid(treSistema.SelectedItem.Key, 2, 5))
        strPropriedades = .selectSingleNode("Grupo_EnderecoFilaMqseries").xml
    End With
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    strLer = objMiu.Executar(strPropriedades, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
        
    If strLer = "" Then
        strOperacao = "Incluir"
        tlbCadastro.Buttons("Excluir").Enabled = False
        Exit Sub
    End If
        
    strOperacao = "Alterar"
    
    If strLer <> "" Then
    
        xmlParamComunicaoSistema.loadXML strLer
        
        txtNomeFila.Text = xmlParamComunicaoSistema.documentElement.selectSingleNode("Grupo_EnderecoFilaMqseries/NO_FILA_MQSE").Text
        
        For Each xmlNode In xmlParamComunicaoSistema.documentElement.selectNodes("//Repeat_RespostaOcorrenciaSistema/*")
            lstOcorrencias.ListItems("OCOR" & xmlNode.selectSingleNode("CO_OCOR_MESG").Text).Checked = True
        Next
    End If
    Exit Sub

ErrorHandler:
    Set objMiu = Nothing
   
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmParamComunicacaoSistema", "flXmlToInterface", lngCodigoErroNegocio, intNumeroSequencialErro
    
End Sub

'Obter as propriedades necessárias para o formulário através de interação com a camada controladora de caso de uso MIU.
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim strMapaNavegacao    As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    Set xmlMapaNavegacao = Nothing
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    strMapaNavegacao = objMiu.ObterMapaNavegacao(enumSistemaSLCC.BUS, strFuncionalidade, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmParamComunicacaoSistema", "flInicializar")
    End If
    
    If xmlParamComunicaoSistema Is Nothing Then
        Set xmlParamComunicaoSistema = CreateObject("MSXML2.DOMDocument.4.0")
        xmlParamComunicaoSistema.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_ParametrosSistema").xml
    End If
    
    Exit Sub

ErrorHandler:
    
    Set objMiu = Nothing
    Set xmlMapaNavegacao = Nothing
    
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmParamComunicacaoSistema", "flInicializar", lngCodigoErroNegocio, intNumeroSequencialErro
    
End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
        Case "Limpar"
            flLimpaCampos
            If fraDetalhe.Enabled = True Then
               txtNomeFila.SetFocus
            End If
        Case "Salvar"
            Call flSalvar
            If strOperacao = "Alterar" Then
              flPosicionaItemListView
            End If
        Case "Excluir"
            strOperacao = "Excluir"
            Call flSalvar
        Case "Sair"
            Unload Me
            strOperacao = ""
    End Select
        
    
    fgCursor False
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmParamComunicacaoSistema - tlbCadastro_ButtonClick"
    
    If strOperacao = "Excluir" Then
        flLimpaCampos
    ElseIf strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If

End Sub

'Posiciona item no listview de sistema de destino.
Private Sub flPosicionaItemListView()

    If treSistema.SelectedItem Is Nothing Then
        flLimpaCampos
        Exit Sub
    End If
    
    treSistema_NodeClick treSistema.SelectedItem

End Sub
   
'Carregar listview com ocorrências para a comunicação.
Private Sub flCarregarOcorrencias()

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem

On Error GoTo ErrorHandler
        
    lstOcorrencias.ListItems.Clear
    
    For Each xmlNode In xmlMapaNavegacao.documentElement.selectNodes("Grupo_Dados/Repeat_OcorrenciaMensagem/*")
        Set objListItem = lstOcorrencias.ListItems.Add(, "OCOR" & xmlNode.selectSingleNode("CO_OCOR_MESG").Text, _
                                                                 xmlNode.selectSingleNode("CO_OCOR_MESG").Text)
        objListItem.SubItems(1) = xmlNode.selectSingleNode("DE_OCOR_MESG").Text
    Next
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    
    fgRaiseError App.EXEName, "frmParamComunicacaoSistema", "flCarregarOcorrencias", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Carregar treeview com informação de empresa e sistema de destino.
Private Sub flCarregartreSistemas()

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strKey                                  As String

On Error GoTo ErrorHandler

    For Each xmlNode In xmlMapaNavegacao.documentElement.selectNodes("Grupo_Dados/Repeat_Sistema/Grupo_Sistema")
        
        On Error Resume Next
            strKey = "E" & Format(xmlNode.selectSingleNode("CO_EMPR").Text, "00000")
            treSistema.Nodes.Add , , strKey, _
                                 xmlNode.selectSingleNode("CO_EMPR").Text & " - " & _
                                 xmlNode.selectSingleNode("NO_REDU_EMPR").Text, _
                                 "Empresa"
            treSistema.Nodes(strKey).Expanded = True
            
        On Error GoTo 0
    
        treSistema.Nodes.Add strKey, _
                             tvwChild, _
                             strKey & "S" & xmlNode.selectSingleNode("SG_SIST").Text, _
                             xmlNode.selectSingleNode("SG_SIST").Text & " - " & xmlNode.selectSingleNode("NO_SIST").Text, _
                             "Sistema"
    
    Next
    
    Exit Sub
ErrorHandler:
    Call fgRaiseError(App.EXEName, Me.Name, "flCarregartreSistemas", lngCodigoErroNegocio)
End Sub

Private Sub treSistema_NodeClick(ByVal Node As MSComctlLib.Node)

On Error GoTo ErrorHandler
    
    If Node Is Nothing Then Exit Sub
    
    flLimpaCampos
        
    If Node.Parent Is Nothing Then
       fraDetalhe.Enabled = False
       tlbCadastro.Buttons("Salvar").Enabled = False
       tlbCadastro.Buttons("Limpar").Enabled = False
       tlbCadastro.Buttons("Excluir").Enabled = False
       Exit Sub
    Else
       fraDetalhe.Enabled = True
    End If
    
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao 'True
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao 'True
    tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao 'True
            
    fraDetalhe.Enabled = True
    
    Call fgCursor(True)
    
    Call flXmlToInterface
       
    Call fgCursor(False)
    
    Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmParamComunicacaoSistema - treSistema_NodeClick"
    
End Sub
