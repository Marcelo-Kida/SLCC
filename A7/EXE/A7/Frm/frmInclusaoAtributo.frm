VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInclusaoAtributo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de Atributos"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "frmInclusaoAtributo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   420
      Left            =   5715
      TabIndex        =   2
      Top             =   4590
      Width           =   1185
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   420
      Left            =   4455
      TabIndex        =   1
      Top             =   4590
      Width           =   1185
   End
   Begin MSComctlLib.ListView lstAtributo 
      Height          =   4005
      Left            =   45
      TabIndex        =   0
      Top             =   540
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   7064
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nome Lógico"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome Físico"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tamanho"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Decimais"
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmInclusaoAtributo.frx":0442
      Top             =   45
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Selecione os Atributos para serem incluídos na mensagem"
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
      Left            =   765
      TabIndex        =   3
      Top             =   225
      Width           =   4965
   End
End
Attribute VB_Name = "frmInclusaoAtributo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela exibição e seleção de atributos disponíveis para inclusão em tipos de mensagem.
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlTipoMensagem                     As MSXML2.DOMDocument40

Private strOperacao                         As String
Private strKeyItemSelected                  As String

Private Const strFuncionalidade             As String = "frmAtributo"
Private strEstruturaAtributo                As String

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Enum enumMoveUpDown
    Up = 1
    Down = 2
End Enum

Public Event AtributosEscolhidos(ByVal strXMLAtributo As String)

Private Sub cmdCancelar_Click()

    Me.Hide

End Sub

Private Sub cmdOk_Click()

Dim xmlSelecao                              As DOMDocument40
Dim objlvwItem                              As ListItem

On Error GoTo ErrorHandler

    Set xmlSelecao = New DOMDocument40
    
    fgAppendNode xmlSelecao, vbNullString, "XML", vbNullString
    
    For Each objlvwItem In lstAtributo.ListItems
        
        If objlvwItem.Selected Then
            fgAppendNode xmlSelecao, _
                         "XML", _
                         objlvwItem.SubItems(1), _
                         vbNullString
                         
            fgAppendAttribute xmlSelecao, _
                              objlvwItem.SubItems(1), _
                              "TP_DADO_ATRB_MESG", _
                              Mid$(objlvwItem.Tag, 2)

            fgAppendAttribute xmlSelecao, _
                              objlvwItem.SubItems(1), _
                              "QT_CTER_ATRB", _
                              objlvwItem.SubItems(3)

            fgAppendAttribute xmlSelecao, _
                              objlvwItem.SubItems(1), _
                              "QT_CASA_DECI_ATRB", _
                              objlvwItem.SubItems(4)
                         
            fgAppendAttribute xmlSelecao, _
                              objlvwItem.SubItems(1), _
                              "IN_ATRB_PRMT_VALO_NEGT", _
                              Left$(objlvwItem.Tag, 1)
                         
        '.SubItems(2) = flTipoDadoToSTR(CLng(xmlNode.selectSingleNode("TP_DADO_ATRB_MESG").Text))
        '.SubItems(3) = xmlNode.selectSingleNode("QT_CTER_ATRB").Text
        '.SubItems(4) = xmlNode.selectSingleNode("QT_CASA_DECI_ATRB").Text
        '.Tag = xmlNode.selectSingleNode("TP_DADO_ATRB_MESG").Text
                
        End If

    Next

    RaiseEvent AtributosEscolhidos(xmlSelecao.xml)
    
    Set xmlSelecao = Nothing

    Me.Hide
    Exit Sub
ErrorHandler:
    
    Set xmlSelecao = Nothing
    
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmInclusaoAtributo - cmdOK_Click")

End Sub

'Carregar listview de atributos.
Private Sub flCarregarlistAtributo()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim xmlNode             As MSXML2.IXMLDOMNode
Dim strPropriedades     As String
Dim strLerTodos         As String
Dim xmlLerTodos         As MSXML2.DOMDocument40
Dim dtmDataServidor     As Date
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
        
    lstAtributo.ListItems.Clear
    lstAtributo.HideSelection = False
        
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    
    xmlMapaNavegacao.selectSingleNode("//Grupo_Propriedades/Grupo_AtributoMensagem/@Operacao").Text = "LerTodos"
    xmlMapaNavegacao.selectSingleNode("//Grupo_Propriedades/Grupo_AtributoMensagem/IN_VIGE").Text = "S"
    strPropriedades = xmlMapaNavegacao.selectSingleNode("//Grupo_AtributoMensagem").xml
    
    strLerTodos = objMiu.Executar(strPropriedades, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    If strLerTodos = "" Then Exit Sub
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLerTodos.loadXML(strLerTodos)
    
    dtmDataServidor = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    
    For Each xmlNode In xmlLerTodos.selectNodes("//Repeat_AtributoMensagem/Grupo_AtributoMensagem")
            
        If Left$(xmlNode.selectSingleNode("NO_ATRB_MESG").Text, 1) <> "/" Then
            
            With lstAtributo.ListItems.Add(, "K" & xmlNode.selectSingleNode("NO_ATRB_MESG").Text, xmlNode.selectSingleNode("NO_TRAP_ATRB").Text)
                
                .SubItems(1) = xmlNode.selectSingleNode("NO_ATRB_MESG").Text
                .SubItems(2) = flTipoDadoToSTR(CLng(xmlNode.selectSingleNode("TP_DADO_ATRB_MESG").Text))
                .SubItems(3) = xmlNode.selectSingleNode("QT_CTER_ATRB").Text
                .SubItems(4) = xmlNode.selectSingleNode("QT_CASA_DECI_ATRB").Text
                .Tag = Left$(Trim$(xmlNode.selectSingleNode("IN_ATRB_PRMT_VALO_NEGT").Text), 1) & _
                       xmlNode.selectSingleNode("TP_DADO_ATRB_MESG").Text
                
            End With
        End If
    Next
    
    lstAtributo.SortKey = 0
    lstAtributo.SortOrder = lvwAscending
    lstAtributo.Sorted = False

    Exit Sub
ErrorHandler:
    
    Set objMiu = Nothing
    Set xmlLerTodos = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flCarregaListView", 0
    
End Sub

'Obter as propriedades necessárias para o formulário através de interação com a camada controladora de caso de uso MIU.
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim xmlNode             As MSXML2.IXMLDOMNode
Dim strPropriedades     As String
Dim strLerTodos         As String
Dim xmlLerTodos         As MSXML2.DOMDocument40
Dim dtmDataServidor     As Date
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    Set objMiu = Nothing

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    
    xmlMapaNavegacao.loadXML objMiu.ObterMapaNavegacao(1, "frmAtributo", vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    Exit Sub
ErrorHandler:

    Set objMiu = Nothing
    Set xmlLerTodos = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmInclusaoAtributo", "flInicializar", lngCodigoErroNegocio, intNumeroSequencialErro
    
End Sub
'Converter o domínio numérico de tipo de dado para literais.
Private Function flTipoDadoToSTR(plngTipoDado As Long) As String
    
    Select Case plngTipoDado
        Case enumTipoDadoAtributo.Alfanumerico
            flTipoDadoToSTR = "Alfanumérico"
        Case enumTipoDadoAtributo.Numerico
            flTipoDadoToSTR = "Numérico"
    End Select

End Function

'Converter as literais de tipo de dado para o domínio numérico.
Private Function flTipoDadoToEnum(pstrTipoDado As String) As Long

    Select Case pstrTipoDado
        Case "Alfanumérico"
            flTipoDadoToEnum = enumTipoDadoAtributo.Alfanumerico
        Case "Numérico"
            flTipoDadoToEnum = enumTipoDadoAtributo.Numerico
    End Select

End Function

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    fgCursor True
    
    fgCenterMe Me
    
    flInicializar
    flCarregarlistAtributo
    
    fgCursor

    Exit Sub
ErrorHandler:
    
    fgCursor False
    
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmInclusaoAtributo - Form_Load")

End Sub

Private Sub lstAtributo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    lstAtributo.Sorted = True
    lstAtributo.SortKey = ColumnHeader.Index - 1

    If lstAtributo.SortOrder = lvwAscending Then
        lstAtributo.SortOrder = lvwDescending
    Else
        lstAtributo.SortOrder = lvwAscending
    End If

End Sub
