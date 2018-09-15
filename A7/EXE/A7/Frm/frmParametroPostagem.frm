VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Begin VB.Form frmParametroPostagem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Parâmetros de Controle de Postagem"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5685
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
      Height          =   1185
      Left            =   0
      TabIndex        =   5
      Top             =   30
      Width           =   5640
      Begin NumBox.Number numTentativas 
         Height          =   285
         Left            =   2175
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1890
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   503
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
      Begin NumBox.Number numCicloVerificacao 
         Height          =   285
         Left            =   2175
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2235
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   503
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
      Begin NumBox.Number numLimiteEntrega 
         Height          =   285
         Left            =   2175
         TabIndex        =   2
         Top             =   330
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
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
      Begin NumBox.Number numLimiteRetirada 
         Height          =   285
         Left            =   2175
         TabIndex        =   3
         Top             =   675
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
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
      Begin VB.Label Label6 
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2820
         TabIndex        =   11
         Top             =   690
         Width           =   120
      End
      Begin VB.Label Label5 
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2820
         TabIndex        =   10
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label1 
         Caption         =   "Tempo Limite Retirada"
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
         Left            =   195
         TabIndex        =   9
         Top             =   705
         Width           =   1920
      End
      Begin VB.Label Label3 
         Caption         =   "Ciclo Verificação"
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
         Left            =   180
         TabIndex        =   8
         Top             =   2265
         Width           =   2145
      End
      Begin VB.Label Label4 
         Caption         =   "Número de Tentativas"
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
         Left            =   165
         TabIndex        =   7
         Top             =   1920
         Width           =   2145
      End
      Begin VB.Label Label2 
         Caption         =   "Tempo Limite Entrega "
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
         Left            =   195
         TabIndex        =   6
         Top             =   360
         Width           =   1860
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   3885
      TabIndex        =   4
      Top             =   1275
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   582
      ButtonWidth     =   1508
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Limpar"
            Key             =   "Limpar"
            ImageKey        =   "Limpar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
      Left            =   0
      Top             =   0
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
            Picture         =   "frmParametroPostagem.frx":0000
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametroPostagem.frx":0112
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametroPostagem.frx":042C
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametroPostagem.frx":077E
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametroPostagem.frx":0890
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametroPostagem.frx":0BAA
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametroPostagem.frx":0EC4
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametroPostagem.frx":11DE
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmParametroPostagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela manutenção de parâmetros de tempo de espera para retirada e entrega de mensagens.
Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private strOperacao                         As String
Private Const strFuncionalidade             As String = "frmParametroPostagem"

'Salvar as configurações correntes.
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim strRetorno          As String
Dim strPropriedades     As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    strRetorno = flValidarCampos()
    
    Call fgCursor(True)
    
    strOperacao = "Incluir"
    
    Call flInterfaceToXml

    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_ParametroControlePostagem").xml
    Call objMiu.Executar(strPropriedades, vntCodErro, vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set objMiu = Nothing
    
    flCarregaInterface
        
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
        
    Call fgCursor(False)

    Exit Sub

ErrorHandler:
    Call fgCursor(False)
    
    Set objMiu = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmParametroControlePostagem", "flSalvar", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    fgCenterMe Me
    Me.Icon = mdiBUS.Icon
    Me.Show
    DoEvents
    
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao 'True
    
    Call fgCursor(True)
    
    Call flInit
    
    flCarregaInterface
    
    Call fgCursor(False)
    
    Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmParametroControlePostagem - Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlMapaNavegacao = Nothing
    
End Sub

'Validar os valores informados para a parametrização.
Private Function flValidarCampos() As String
    
    flValidarCampos = ""
    
End Function

'Mover valores do formulário para XML para envio ao objeto de negócio.
Private Function flInterfaceToXml() As String
    
On Error GoTo ErrorHandler
          
    With xmlMapaNavegacao.documentElement
        .selectSingleNode("Grupo_Propriedades/Grupo_ParametroControlePostagem/@Operacao").Text = strOperacao
        .selectSingleNode("Grupo_Propriedades/Grupo_ParametroControlePostagem/DH_PARM").Text = "20030712160000"
        .selectSingleNode("Grupo_Propriedades/Grupo_ParametroControlePostagem/QT_TENT_ENVI_MESG").Text = numTentativas.Valor
        .selectSingleNode("Grupo_Propriedades/Grupo_ParametroControlePostagem/QT_FREQ_VERI").Text = numCicloVerificacao.Valor
        .selectSingleNode("Grupo_Propriedades/Grupo_ParametroControlePostagem/QT_TEMP_ENTG_MESG").Text = numLimiteEntrega.Valor
        .selectSingleNode("Grupo_Propriedades/Grupo_ParametroControlePostagem/QT_TEMP_RETI_MESG").Text = numLimiteRetirada.Valor
    End With
        
    Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, "frmParametroControlePostagem", "flInterfaceToXml", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

'Carregar os campos do formulário com os valores recebidos da camada de negócio.
Private Sub flXmlToInterface(ByVal pobjNode As MSXML2.IXMLDOMNode)

On Error GoTo ErrorHandler
    
    fraDetalhe.Caption = "Ultima Atualização: " & Format(fgDtHrStr_To_DateTime(pobjNode.selectSingleNode("DH_PARM").Text), gstrMascaraDataHoraDtp)
    
    numTentativas.Valor = pobjNode.selectSingleNode("QT_TENT_ENVI_MESG").Text
    numCicloVerificacao.Valor = pobjNode.selectSingleNode("QT_FREQ_VERI").Text
    numLimiteEntrega.Valor = pobjNode.selectSingleNode("QT_TEMP_ENTG_MESG").Text
    numLimiteRetirada.Valor = pobjNode.selectSingleNode("QT_TEMP_RETI_MESG").Text
        
    Exit Sub
ErrorHandler:

    fgRaiseError App.EXEName, "frmParametroControlePostagem", "flXmlToInterface", lngCodigoErroNegocio, intNumeroSequencialErro
    
End Sub

'Obter as propriedades necessárias para o formulário através de interação com a camada controladora de caso de uso MIU.
Private Sub flInit()

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
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmParametroControlePostagem", "flInit")
    Else
    
    End If
    
    Exit Sub
ErrorHandler:

    Set objMiu = Nothing
    Set xmlMapaNavegacao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmParametroControlePostagem", "flInit", lngCodigoErroNegocio, intNumeroSequencialErro
    
End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    Select Case Button.Key
        Case "Limpar"
            'Não se aplica
        Case "Excluir"
            'Não se aplica
        Case "Salvar"
            Call flSalvar
        Case "Sair"
            Unload Me
    End Select
    
    Exit Sub
ErrorHandler:

    mdiBUS.uctLogErros.MostrarErros Err, "frmParametroControlePostagem - tlbCadastro_ButtonClick"

End Sub

'Carregar valores de parametrização.
Private Sub flCarregaInterface()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim objNodeList         As MSXML2.IXMLDOMNodeList
Dim strPropriedades     As String
Dim strLerTodos         As String
Dim objLerTodos         As MSXML2.DOMDocument40
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
        
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlMapaNavegacao.selectSingleNode("//Grupo_Propriedades/Grupo_ParametroControlePostagem/@Operacao").Text = "LerTodos"
    strPropriedades = xmlMapaNavegacao.selectSingleNode("//Grupo_ParametroControlePostagem").xml
    strLerTodos = objMiu.Executar(strPropriedades, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    If strLerTodos = "" Then Exit Sub
    
    Set objLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    Call objLerTodos.loadXML(strLerTodos)
    
    Set objNodeList = objLerTodos.selectSingleNode("//Repeat_ParametroControlePostagem").childNodes
    
    Call flXmlToInterface(objLerTodos.selectSingleNode("//Repeat_ParametroControlePostagem").childNodes.Item(0))
    
    Set objLerTodos = Nothing
    
    Exit Sub
ErrorHandler:

    Set objMiu = Nothing
    Set objLerTodos = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmParametroControlePostagem", "flCarregaInterface", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub
