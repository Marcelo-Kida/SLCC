VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Begin VB.Form frmTipoMensagemSaida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Tipos de Mensagens de Saída"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7860
   Begin VB.Frame Frame1 
      Height          =   2340
      Left            =   45
      TabIndex        =   4
      Top             =   0
      Width           =   7800
      Begin MSComctlLib.ListView lstTipoMsgSaida 
         Height          =   2070
         Left            =   90
         TabIndex        =   5
         Top             =   165
         Width           =   7650
         _ExtentX        =   13494
         _ExtentY        =   3651
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   9085
         EndProperty
      End
   End
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
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   2325
      Width           =   7815
      Begin VB.TextBox txtDescricao 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   1050
         Width           =   7530
      End
      Begin NumBox.Number txtCodigo 
         Height          =   330
         Left            =   135
         TabIndex        =   7
         Top             =   405
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSelecao     =   0   'False
         AceitaNegativo  =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         TabIndex        =   3
         Top             =   150
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Descrição"
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
         Left            =   120
         TabIndex        =   2
         Top             =   810
         Width           =   2145
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   4275
      TabIndex        =   6
      Top             =   3915
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   582
      ButtonWidth     =   1535
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "Limpar"
            ImageKey        =   "Limpar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   45
      Top             =   3780
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
            Picture         =   "frmTipoMensagemSaida.frx":0000
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemSaida.frx":0112
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemSaida.frx":042C
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemSaida.frx":077E
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemSaida.frx":0890
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemSaida.frx":0BAA
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemSaida.frx":0EC4
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemSaida.frx":11DE
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTipoMensagemSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Empresa        : Regerbanc
'Pacote         :
'Classe         : frmAtributo
'Data Criação   : 01/07/2003
'Objetivo       :
'
'Analista       : Marcelo Kida
'
'Programador    : Marcelo Kida
'Data           : 01/07/2003
'
'Teste          :
'Autor          :
'
'Data Alteração : Douglas Cavalcante
'Autor          : 23/09/2003
'Objetivo       : Seguindo os Padrões Definidos foi alterado o Form.
'
'Data Alteração : 26/09/2003
'Autor          : Eder Andrade
'Objetivo       : Impedir que sejam selecionadas datas de vigência inválidas

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlTipoMensagemSaida                As MSXML2.DOMDocument40

Private strOperacao                         As String
Private strKeyItemSelected                  As String
Private Const strFuncionalidade             As String = "frmTipoMensagemSaida"

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

On Error GoTo ErrorHandler

    If lstTipoMsgSaida.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub
    
    blnEncontrou = False
    For Each objListItem In lstTipoMsgSaida.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lstTipoMsgSaida_ItemClick objListItem
           lstTipoMsgSaida.ListItems(strKeyItemSelected).EnsureVisible
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


Private Sub flLimpaCampos()

On Error GoTo ErrorHandler
        
    strOperacao = "Incluir"
    
    txtCodigo.Valor = 0
    txtCodigo.Enabled = True
    txtDescricao.Text = ""
    
    lstTipoMsgSaida.Sorted = False
    
    tlbCadastro.Buttons("Excluir").Enabled = False
    
    Exit Sub
    
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flLimpaCampos", 0

End Sub

Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMiu                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                              As A7Miu.clsMIU
#End If

Dim strRetorno                              As String
Dim strKey                                  As String

On Error GoTo ErrorHandler

    strRetorno = flValidarCampos()
        
    If strRetorno <> "" Then
        If strRetorno = "TipoMensagemSaida" Then
            If MsgBox("Este Tipo de Mensagem de Saída está associado a um tipo de mensagem." & vbCrLf & _
                      "Somente a Descrição será alterada." & vbCrLf & _
                      "Confirma ?", vbYesNo, "Tipo de Mensagem de Saída") = vbNo Then
                Exit Sub
            End If
        Else
            frmMural.Display = strRetorno
            frmMural.Show vbModal
            Exit Sub
       End If
    End If
    
    Call flInterfaceToXml
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    Call objMiu.Executar(xmlTipoMensagemSaida.xml)
    Set objMiu = Nothing
    
    If Not lstTipoMsgSaida.SelectedItem Is Nothing Then
        If strOperacao = "Incluir" Then
            strKey = "K" & Format(txtCodigo.Valor, "0000")
        Else
            strKey = lstTipoMsgSaida.SelectedItem.Key
        End If
    End If
    
    strKeyItemSelected = strKey
    
    Call flCarregaListView
            
    If strKey <> "" Then
        lstTipoMsgSaida.ListItems.Item(strKey).Selected = True
        lstTipoMsgSaida.ListItems.Item(strKey).EnsureVisible
    End If
        
    With xmlTipoMensagemSaida.documentElement
        .selectSingleNode("@Operacao").Text = "Ler"
    End With
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlTipoMensagemSaida.loadXML objMiu.Executar(xmlTipoMensagemSaida.xml)
    Set objMiu = Nothing
    
    If strOperacao = "Incluir" Then
       strKeyItemSelected = ""
    End If
        
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
                
    If strOperacao = "Incluir" Then
        flProtegerChave
    End If
    
    Exit Sub

ErrorHandler:
    
    Set objMiu = Nothing

    fgRaiseError App.EXEName, "frmAtributo", "flSalvar", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
        
    fgCenterMe Me
    Me.Icon = mdiBUS.Icon
    
    DoEvents
    
    fgCursor True
    
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao 'True
    tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao 'True
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao 'True
    
    Call flLimpaCampos
    Call flInicializar
    Call flDefinirTamanhoMaximoCampos
    
    Me.Show

    Call flCarregaListView
    Call fgCursor(False)
    
    Exit Sub
    
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmAtributo - Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTipoMensagemSaida = Nothing
End Sub

Private Function flValidarCampos() As String
    
#If EnableSoap = 1 Then
    Dim objMiu                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                              As A7Miu.clsMIU
#End If
    
Dim strPropriedades                         As String
Dim strXML                                  As String
    
On Error GoTo ErrorHandler
    
    If txtCodigo.Valor = 0 Then
        flValidarCampos = "Digite o código para o tipo de mensagem de saída."
        txtCodigo.SetFocus
        Exit Function
    End If
    
    If txtCodigo.Valor > 9999 Or txtCodigo.Valor < 1 Then
        flValidarCampos = "Código para o tipo de mensagem de saída inválido."
        txtCodigo.SetFocus
        Exit Function
    End If
    
    If Trim(txtDescricao.Text) = "" Then
        flValidarCampos = "Digite a descrição para o tipo de mensagem de saída."
        txtDescricao.SetFocus
        Exit Function
    End If
    
    If strOperacao = "Alterar" Then
    ' =====  Apos definir função para verificação de uso implementá-la neste ponto.
    
    
    
'        xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoMensagemAtributo/@Operacao").Text = "Ler"
'        xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoMensagemAtributo/CO_MESG_SAID").Text = Trim(txtNomeFisico)
'
'        strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoMensagemAtributo").xml
'
'        Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
'        strPropriedades = objMiu.Executar(strPropriedades)
'        Set objMiu = Nothing
'
'        If Trim(strPropriedades) <> "" Then
'
'            With xmlTipoMensagemSaida.documentElement
'
'                If optTipoDado(0) Then
'                    If .selectSingleNode("TP_DADO_ATRB_MESG").Text <> enumTipoDadoAtributo.Numerico Then
'                        flValidarCampos = "Atributo"
'                        Exit Function
'                    End If
'                End If
'
'                If optTipoDado(1) Then
'                    If .selectSingleNode("TP_DADO_ATRB_MESG").Text <> enumTipoDadoAtributo.Alfanumerico Then
'                        flValidarCampos = "Atributo"
'                        Exit Function
'                    End If
'                End If
'
'                If .selectSingleNode("QT_CTER_ATRB").Text <> numTamanho.Valor Then
'                    flValidarCampos = "Atributo"
'                    Exit Function
'                End If
'
'                If .selectSingleNode("QT_CASA_DECI_ATRB").Text <> numDecimais.Valor Then
'                    flValidarCampos = "Atributo"
'                    Exit Function
'                End If
'
'
'            End With
'
'            flValidarCampos = vbNullString
'
'            Exit Function
'        End If
    End If
    
    flValidarCampos = ""

    Exit Function
ErrorHandler:
    
    Set objMiu = Nothing

    fgRaiseError App.EXEName, TypeName(Me), "flValidarCampos", 0

End Function

Private Function flInterfaceToXml() As String
    
On Error GoTo ErrorHandler
        
    With xmlTipoMensagemSaida.documentElement
        
        .selectSingleNode("@Operacao").Text = strOperacao
        .selectSingleNode("CO_MESG_SAID").Text = txtCodigo.Valor
        .selectSingleNode("DE_MESG_SAID").Text = Trim(txtDescricao.Text)
    
    End With
    
    Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0

End Function

Private Sub flDefinirTamanhoMaximoCampos()

On Error GoTo ErrorHandler
                
    With xmlMapaNavegacao.documentElement
        'txtCodigo.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_TipoMensagemSaida/CO_MESG_SAID/@Tamanho").Text
        txtDescricao.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_TipoMensagemSaida/DE_MESG_SAID/@Tamanho").Text
    End With
        
    Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flDefinirTamanhoMaximoCampos", 0

End Sub

Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMiu                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                              As A7Miu.clsMIU
#End If

On Error GoTo ErrorHandler
    
    With xmlTipoMensagemSaida.documentElement
        .selectSingleNode("@Operacao").Text = "Ler"
        .selectSingleNode("CO_MESG_SAID").Text = Mid$(lstTipoMsgSaida.SelectedItem.Key, 2)
    End With
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlTipoMensagemSaida.loadXML objMiu.Executar(xmlTipoMensagemSaida.xml)
    Set objMiu = Nothing
       
    tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao 'True
       
    With xmlTipoMensagemSaida.documentElement
   
        txtCodigo.Valor = CLng(.selectSingleNode("CO_MESG_SAID").Text)
        txtDescricao.Text = .selectSingleNode("DE_MESG_SAID").Text
        
    End With
        
    Exit Sub
    
ErrorHandler:

    fgRaiseError App.EXEName, TypeName(Me), "flXmlToInterface", 0
    
End Sub

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMiu                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                              As A7Miu.clsMIU
#End If

Dim strMapaNavegacao                        As String

On Error GoTo ErrorHandler
    
    Set xmlMapaNavegacao = Nothing
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    strMapaNavegacao = objMiu.ObterMapaNavegacao(enumSistemaSLCC.BUS, strFuncionalidade)
    Set objMiu = Nothing
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmTipoMensagemSaida", "flInicializar")
    End If
    
    If xmlTipoMensagemSaida Is Nothing Then
       Set xmlTipoMensagemSaida = CreateObject("MSXML2.DOMDocument.4.0")
       xmlTipoMensagemSaida.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_TipoMensagemSaida").xml
    End If
    
    Exit Sub
ErrorHandler:

    Set objMiu = Nothing
    Set xmlMapaNavegacao = Nothing

    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0
    
End Sub

Private Sub lstTipoMsgSaida_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

    lstTipoMsgSaida.Sorted = True
    lstTipoMsgSaida.SortKey = ColumnHeader.Index - 1

    If lstTipoMsgSaida.SortOrder = lvwAscending Then
        lstTipoMsgSaida.SortOrder = lvwDescending
    Else
        lstTipoMsgSaida.SortOrder = lvwAscending
    End If

End Sub

Private Sub lstTipoMsgSaida_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    Call flLimpaCampos
    
    'Não permitir que a chave da tabela seja alterada
    txtCodigo.Enabled = False
    
    strOperacao = "Alterar"
    strKeyItemSelected = Item.Key
    Call flXmlToInterface
    Call fgCursor(False)
    
    Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmAtributo - lstTipoMensagemSaida_ItemClick"

    Call flCarregaListView
    
    If strOperacao = "Excluir" Then
        flLimpaCampos
    ElseIf strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
        Case "Limpar"
            Call flLimpaCampos
        Case "Salvar"
            Call flSalvar
        Case "Excluir"
            strOperacao = "Excluir"
            flExcluir
            Call flLimpaCampos
        Case "Sair"
            Unload Me
            strOperacao = ""
            fgCursor
            Exit Sub
    End Select
                
    txtCodigo.SetFocus
    
    If strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If
            
    fgCursor
            
    Exit Sub

ErrorHandler:

    fgCursor

    mdiBUS.uctLogErros.MostrarErros Err, "frmAtributo - tlbCadastro_ButtonClick"
    
    Call flCarregaListView
    
    If strOperacao = "Excluir" Then
        flLimpaCampos
    ElseIf strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If

End Sub

Private Sub flCarregaListView()

#If EnableSoap = 1 Then
    Dim objMiu                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                              As A7Miu.clsMIU
#End If

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem
Dim strPropriedades                         As String
Dim strLerTodos                             As String
Dim xmlLerTodos                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
        
    lstTipoMsgSaida.ListItems.Clear
    lstTipoMsgSaida.HideSelection = False
    
    Set xmlNode = Nothing
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoMensagemSaida/@Operacao").Text = "LerTodos"
    strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoMensagemSaida").xml
    strLerTodos = objMiu.Executar(strPropriedades)
    
    Set objMiu = Nothing
    
    If strLerTodos = "" Then Exit Sub
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLerTodos.loadXML(strLerTodos)
    
    For Each xmlNode In xmlLerTodos.selectSingleNode("//Repeat_TipoMensagemSaida").childNodes
        
        With xmlNode
                
            Set objListItem = lstTipoMsgSaida.ListItems.Add(, "K" & Format(CLng(.selectSingleNode("CO_MESG_SAID").Text), "0000"), .selectSingleNode("CO_MESG_SAID").Text)
            
            objListItem.SubItems(1) = .selectSingleNode("DE_MESG_SAID").Text
        
        End With
    Next

    Exit Sub
ErrorHandler:

    Set objMiu = Nothing
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaListView", 0

End Sub

Private Sub flExcluir()

#If EnableSoap = 1 Then
    Dim objMiu                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                              As A7Miu.clsMIU
#End If

On Error GoTo ErrorHandler
      
    If MsgBox("Confirma Exclusão ?", vbYesNo + vbQuestion, "Tipo de Mensagem de Saída") = vbNo Then Exit Sub
   
    Call fgCursor(True)
        
    Call flInterfaceToXml

    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    Call objMiu.Executar(xmlTipoMensagemSaida.xml)
    Set objMiu = Nothing
    
    Call flCarregaListView
        
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption

    Call fgCursor(False)

    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    
    Set objMiu = Nothing
    
    fgRaiseError App.EXEName, TypeName(Me), "flExcluir", 0

End Sub
Private Sub flProtegerChave()
    
   strOperacao = "Alterar"
   txtCodigo.Enabled = False
   tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao
 
End Sub



