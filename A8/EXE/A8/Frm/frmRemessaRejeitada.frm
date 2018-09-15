VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRemessaRejeitada 
   Caption         =   "Consulta - Remessa Rejeitada"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   12525
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstRemessaRejeitada 
      Height          =   5595
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9869
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Data/Hora Rejeição"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Empresa"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sistema"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tipo"
         Object.Width           =   10583
      EndProperty
   End
   Begin MSComctlLib.ImageList imgOrder 
      Left            =   5940
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":0000
            Key             =   "Cima"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":0192
            Key             =   "Baixo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbFiltro 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   7245
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   582
      ButtonWidth     =   2434
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Definir &Filtro"
            Key             =   "Filtro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "Atualizar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Sair"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   4980
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":0324
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":0436
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":0548
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":089A
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":0BEC
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":0F3E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":1290
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":15AA
            Key             =   "confirmar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":19FC
            Key             =   "agendar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":1D16
            Key             =   "Cima"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemessaRejeitada.frx":1EA8
            Key             =   "Baixo"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRemessaRejeitada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:35:10
'-------------------------------------------------
'' Objeto responsável pelo envio da solicitação (Consulta às remessas enviadas por
'' outros sistemas, e rejeitadas pelo sistema A8) à camada controladora de caso de
'' uso A8MIU.
''
'' São consideradas especificamente as classes destino:
''      A8MIU.clsRemessaRejeitada
''
Option Explicit

'Este objeto xmlMapaNavegacao é carregado com as propriedades do objRemessaRejeitada
' e todas as coleções que este for utilizar
Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private strOperacao                         As String

'Variaveis para a utilização do Filtro
Private strFiltroXML                        As String
Private blnPrimeiraConsulta                 As Boolean

Private Const strFuncionalidade             As String = "REMESSAREJEITADA"
Private Const strTableComboInicial          As String = "Empresa"

Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1

Private intCompareCol                       As Integer
Private intCompareOrder                     As Integer

Private Const COMPARE_ASC                   As Integer = 1
Private Const COMPARE_DESC                  As Integer = -1

'Definição de Colunas e Linha do Grid
Private Const COL_DESCRICAO                 As Integer = 0

'Definição das Colunas do Grid
Private Const COL_DH_REME_REJE              As Integer = 0
Private Const COL_EMPRESA                   As Integer = 1
Private Const COL_SISTEMA                   As Integer = 2
Private Const COL_TIPO_MENSAGEM             As Integer = 3
Private Const COL_KEY                       As Integer = 4

Private lngCodigoErroNegocio                As Long

Private lngIndexClassifList                 As Long

'' Encaminhar a solicitação (Leitura de todas as remessas rejeitadas, obedecendo
'' aos critérios de seleção informados) à camada controladora de caso de uso
'' (componente / classe / metodo ) :
''
'' A8MIU.clsRemessaRejeitada.LerTodos
''
'' O método retornará uma String XML para a camada de interface.
''
Private Sub flCarregarFlexGrid(ByRef pxmlDocFiltros As String)

#If EnableSoap = 1 Then
    Dim objRemessaRejeitada                 As MSSOAPLib30.SoapClient30
#Else
    Dim objRemessaRejeitada                 As A8MIU.clsRemessaRejeitada
#End If

Dim xmlRejeitada                            As MSXML2.DOMDocument40
Dim xmlDomRejeitada                         As MSXML2.IXMLDOMNode
Dim strXMLRetorno                           As String

Dim strKey                                  As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Call fgCursor(True)
    
    gintRowPositionAnt = 0

    Set objRemessaRejeitada = fgCriarObjetoMIU("A8MIU.clsRemessaRejeitada")

    Set xmlRejeitada = CreateObject("MSXML2.DOMDocument.4.0")
    
    strXMLRetorno = objRemessaRejeitada.LerTodos(pxmlDocFiltros, _
                                                 vntCodErro, _
                                                 vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    lstRemessaRejeitada.ListItems.Clear
    
    'caso a tabela esteja sem registros não tem como carregar um XML,
    'sendo assim vai para o fim da rotina.
    If Trim(strXMLRetorno) <> "" Then
       If Not xmlRejeitada.loadXML(strXMLRetorno) Then
          '100 - Documento XML Inválido.
          lngCodigoErroNegocio = 100
          GoTo ErrorHandler
       End If
    Else
       Call fgCursor(False)
       Exit Sub
    End If
    
    For Each xmlDomRejeitada In xmlRejeitada.documentElement.selectNodes("//Repeat_Erro/*")
        strKey = ";" & xmlDomRejeitada.selectSingleNode(".//SG_SIST_ORIG_INFO").Text & _
                ";" & xmlDomRejeitada.selectSingleNode(".//TP_MESG_INTE").Text & _
                ";" & xmlDomRejeitada.selectSingleNode(".//CO_EMPR").Text & _
                ";" & xmlDomRejeitada.selectSingleNode(".//CO_TEXT_XML_REJE").Text & _
                ";" & xmlDomRejeitada.selectSingleNode(".//DH_REME_REJE").Text & _
                ";" & xmlDomRejeitada.selectSingleNode(".//OWNER").Text
        
        With lstRemessaRejeitada.ListItems.Add(, strKey, CStr(fgDtHrStr_To_DateTime(xmlDomRejeitada.selectSingleNode(".//DH_REME_REJE").Text)))
            .SubItems(1) = xmlDomRejeitada.selectSingleNode(".//CO_EMPR").Text & _
                           " - " & _
                           xmlDomRejeitada.selectSingleNode(".//NO_REDU_EMPR").Text
            .SubItems(2) = xmlDomRejeitada.selectSingleNode(".//SG_SIST_ORIG_INFO").Text & _
                           " - " & _
                           xmlDomRejeitada.selectSingleNode(".//NO_SIST").Text
            .SubItems(3) = xmlDomRejeitada.selectSingleNode(".//TP_MESG_INTE").Text & _
                           " - " & _
                           xmlDomRejeitada.selectSingleNode(".//NO_TIPO_MESG").Text
        End With
        
    Next xmlDomRejeitada
    
    Call fgClassificarListview(Me.lstRemessaRejeitada, lngIndexClassifList, True)
    
    Set xmlRejeitada = Nothing
    Set objRemessaRejeitada = Nothing
    Call fgCursor(False)

Exit Sub
ErrorHandler:

    Set xmlRejeitada = Nothing
    Set objRemessaRejeitada = Nothing
    Call fgCursor(False)
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, Me.Name, "frmRemessaRejeitada - flCarregarFlexGrid", 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        fgLockWindow Me.hwnd
        Call tlbFiltro_ButtonClick(tlbFiltro.Buttons(gstrAtualizar))
        fgLockWindow 0
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    fgLockWindow 0
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmRemessaRejeitada - Form_KeyDown", Me.Caption

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    intCompareOrder = COMPARE_ASC
    
    fgLockWindow Me.hwnd
    
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    Call fgCursor(True)
    
    blnPrimeiraConsulta = True
    
    fgLockWindow 0

    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA8.frmRemessaRejeitada
    Load objFiltro
    
    Call objFiltro.fgCarregarPesquisaAnterior

    Me.Refresh

    Call fgCursor(False)

Exit Sub
ErrorHandler:

    Call fgCursor(False)

    mdiLQS.uctlogErros.MostrarErros Err, "frmRemessaRejeitada - Form_Load", Me.Caption

End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    With Me
        
        .lstRemessaRejeitada.Left = 0
        '.lstRemessaRejeitada.Top = .lblGrupoVeiculo.Height
        .lstRemessaRejeitada.Width = .ScaleWidth
        .lstRemessaRejeitada.Height = .ScaleHeight - .lstRemessaRejeitada.Top - .tlbFiltro.Height
    End With

End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

'Inicializa o conteúdo dos controles de tela e variáveis
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim strMapaNavegacao                        As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Set xmlMapaNavegacao = Nothing

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    'strMapaNavegacao = objMIU.ObterMapaNavegacao(strFuncionalidade, vntCodErro, vntMensagemErro)
    'If vntCodErro <> 0 Then
    '    GoTo ErrorHandler
    'End If
    
    Set objMIU = Nothing

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmRemessaRejeitada", "flInicializar")
    End If

Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, Me.Name, "frmRemessaRejeitada - flInicializar", 0

End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Set frmRemessaRejeitada = Nothing
    gintRowPositionAnt = 0

End Sub

Private Sub lstRemessaRejeitada_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   
On Error GoTo ErrorHandler
        
    Call fgClassificarListview(Me.lstRemessaRejeitada, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index

Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - lstRemessaRejeitada_ColumnClick", Me.Caption
End Sub

Private Sub lstRemessaRejeitada_DblClick()

#If EnableSoap = 1 Then
    Dim objRemessa                          As MSSOAPLib30.SoapClient30
#Else
    Dim objRemessa                          As A8MIU.clsRemessaRejeitada
#End If

Dim arrChave()                              As String
Dim strXML                                  As String
Dim xmlLer                                  As MSXML2.DOMDocument40
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    If lstRemessaRejeitada.SelectedItem Is Nothing Then Exit Sub

    fgCursor True
    arrChave = Split(lstRemessaRejeitada.SelectedItem.Key, ";")

    Set objRemessa = fgCriarObjetoMIU("A8MIU.clsRemessaRejeitada")

    strXML = objRemessa.Ler(UCase(arrChave(1)), _
                            CInt(arrChave(2)), _
                            CLng(arrChave(3)), _
                            CLng(arrChave(4)), _
                            arrChave(5), _
                            arrChave(6), _
                            vntCodErro, _
                            vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If Len(strXML) = 0 Then Exit Sub
    
    Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlLer.loadXML(strXML) Then
        fgCursor False
        fgErroLoadXML xmlLer, "lstRemessaRejeitada_DblClick", "", ""
    End If
    
    strXML = xmlLer.documentElement.selectSingleNode("TX_XML_ERRO").Text
    xmlLer.loadXML strXML
    
    If UCase(arrChave(6)) = "A8HIST" Then
        frmDetalheRemessa.lngCO_TEXT_XML_REJE = CLng(arrChave(4) * -1)
    Else
        frmDetalheRemessa.lngCO_TEXT_XML_REJE = CLng(arrChave(4))
    End If
    frmDetalheRemessa.strXMLErro = strXML
    frmDetalheRemessa.Show
    
    fgCursor False

Exit Sub
ErrorHandler:

    fgCursor False
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmRemessaRejeitada - Form_Load", Me.Caption

End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, lsTituloTableCombo As String)

On Error GoTo ErrorHandler

    strFiltroXML = xmlDocFiltros
    
    If Trim(strFiltroXML) <> vbNullString Then
        If fgMostraFiltro(strFiltroXML, blnPrimeiraConsulta) Then
            blnPrimeiraConsulta = False
            Call tlbFiltro_ButtonClick(tlbFiltro.Buttons("Filtro"))
        End If
        
        If InStr(1, strFiltroXML, "DataIni") = 0 Then
            frmMural.Caption = Me.Caption
            frmMural.Display = "Obrigatória a seleção do filtro DATA."
            frmMural.Show vbModal
        Else
            Call fgCursor(True)
            Call flCarregarFlexGrid(strFiltroXML)
            Call fgCursor(False)
        End If
    
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmRemessaRejeitada - objFiltro_AplicarFiltro", Me.Caption
    
End Sub

Private Sub tlbFiltro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    fgCursor True

    Select Case Button.Key
           Case "Sair"
                Unload Me
           
           Case gstrAtualizar
                If InStr(1, strFiltroXML, "DataIni") = 0 Then
                    frmMural.Caption = Me.Caption
                    frmMural.Display = "Obrigatória a seleção do filtro DATA."
                    frmMural.Show vbModal
                Else
                    Call fgCursor(True)
                    Call flCarregarFlexGrid(strFiltroXML)
                    Call fgCursor(False)
                End If
           
           Case Else
                blnPrimeiraConsulta = False
            
                Set objFiltro = New frmFiltro
                Set objFiltro.FormOwner = Me
                objFiltro.TipoFiltro = enumTipoFiltroA8.frmRemessaRejeitada
                objFiltro.Show vbModal
    
    End Select
    
    fgCursor

Exit Sub
ErrorHandler:
    fgCursor True
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbFiltro_ButtonClick"
    
End Sub
