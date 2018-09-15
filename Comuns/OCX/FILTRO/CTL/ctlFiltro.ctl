VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ctlFiltro 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5640
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   5640
   Begin MSComctlLib.ImageCombo cboFiltro 
      Height          =   330
      Index           =   0
      Left            =   2565
      TabIndex        =   8
      Top             =   360
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin MSComCtl2.DTPicker dtpFiltro 
      Height          =   330
      Index           =   0
      Left            =   2565
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   71041025
      CurrentDate     =   37826
   End
   Begin VB.TextBox txtFiltro 
      Height          =   330
      Index           =   0
      Left            =   2565
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.TextBox txtOrdem 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Index           =   0
      Left            =   2070
      TabIndex        =   0
      Text            =   "0"
      Top             =   360
      Width           =   270
   End
   Begin MSComCtl2.UpDown udClassificacao 
      Height          =   330
      Index           =   0
      Left            =   2340
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      OrigLeft        =   2340
      OrigTop         =   360
      OrigRight       =   2580
      OrigBottom      =   660
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ImageList imgFiltro 
      Left            =   450
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlFiltro.ctx":0000
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlFiltro.ctx":0452
            Key             =   "Aplicar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Height          =   330
      Left            =   3195
      TabIndex        =   2
      Top             =   3240
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   582
      ButtonWidth     =   2011
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgFiltro"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Aplicar"
            Object.ToolTipText     =   "Aplicar Filtro"
            ImageIndex      =   2
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            Object.ToolTipText     =   "Cancelar Filtro"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFechar 
      BackColor       =   &H80000010&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5445
      TabIndex        =   4
      Top             =   0
      Width           =   390
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H8000000C&
      Caption         =   " Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5685
   End
   Begin VB.Label lblFiltro 
      AutoSize        =   -1  'True
      Height          =   330
      Index           =   0
      Left            =   195
      TabIndex        =   1
      Top             =   375
      Visible         =   0   'False
      Width           =   1890
   End
End
Attribute VB_Name = "ctlFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''Empresa            : Regerbanc - Participações , Negócios e Serviços LTDA
''Componente         : ctlFiltro
''Classe             :
''Data Criação       : 18/07/2003
''Objetivo           : OCX responsável pela aplicação de filtros em telas onde sejam acessados bancos de dados
''Analista           : Adilson G. Damasceno/Carlos Fortes/Marcelo /Marcelo
''
''Programador        : Marcelo /Marcelo /Cassiano Nicolosi
''Data               : 22/07/2003
''
''Data Teste         :
''Autor              :
''
''Data Alteração     :
''Autor              :
''Objetivo           :
'
'Option Explicit
'
''Event Declarations:
'Event AplicarFiltro(xmlDocFiltros As String, lsTituloTableCombo As String, xmlDocPrimSelecao As String)
'Event CancelarFiltro()
'Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
'Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
'Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
'
''Private Variables:
'Private objXMLMapaNavegacao                 As MSXML2.DOMDocument40
'Private objXMLMapaFiltro                    As MSXML2.DOMDocument40
'Private xmlDocFiltros                       As String
'Private xmlDocPrimSelecao                   As String
'Private lsTituloTableCombo                  As String
'Private lcControlesFiltro                   As Collection
'
'Private Enum enumMovimentacao
'    Proximo = 1
'    Anterior = 2
'End Enum
'
'Private Enum enumTipoTexto
'    Alfanumerico = 1
'    Numerico = 2
'End Enum
'
'Private Enum enumTipoControle
'    ctlComboBox = 1
'    ctlTextBox = 2
'    ctlDTPicker = 3
'End Enum
'
''Utiliza o xml "Mapa de Filtro" para arranjar dinamicamente os controles que farão parte da OCX
''Utiliza o xml "Mapa de Navegação" para a leitura de tabelas necessária para o carregamento do(s) combo(s),
''caso existam
'Public Function fgInicializar(ByVal piSistemaSLCC As enumSistemaSLCC) As Boolean
'
'Dim liCount                                 As Integer
'Dim objMIU                                  As MIU.clsMIU
'Dim objDomNode                              As MSXML2.IXMLDOMNode
'Dim objArgControle                          As Object
'
'    On Error GoTo ErrorHandler
'
'    Set objMIU = CreateObject("MIU.clsMIU")
'    Set objXMLMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
'    Set objXMLMapaFiltro = CreateObject("MSXML2.DOMDocument.4.0")
'
'    If Not objXMLMapaFiltro.loadXML(objMIU.ObterMapaFiltro(piSistemaSLCC, UserControl.Parent.Name)) Then
'        Call fgErroLoadXML(objXMLMapaFiltro, App.EXEName, "ctlFiltro", "flInicializar")
'    End If
'
'    Set lcControlesFiltro = New Collection
'
'    For Each objDomNode In objXMLMapaFiltro.documentElement.selectNodes("//" & UserControl.Parent.Name & "/*")
'        Select Case Val(objDomNode.selectSingleNode("Tipo").Text)
'            Case enumTipoControle.ctlComboBox
'                If liCount > 0 Then Load cboFiltro(liCount)
'                Set objArgControle = cboFiltro(liCount)
'
'            Case enumTipoControle.ctlTextBox
'                If liCount > 0 Then Load txtFiltro(liCount)
'                Set objArgControle = txtFiltro(liCount)
'                objArgControle.Tag = objDomNode.selectSingleNode("TipoTexto").Text
'                objArgControle.MaxLength = objDomNode.selectSingleNode("MaxLen").Text
'
'            Case enumTipoControle.ctlDTPicker
'                If liCount > 0 Then Load dtpFiltro(liCount)
'                Set objArgControle = dtpFiltro(liCount)
'                objArgControle.Value = Date
'                objArgControle.Value = Null
'
'        End Select
'
'        lcControlesFiltro.Add objArgControle
'
'        If liCount > 0 Then
'            Load lblFiltro(liCount)
'            Load txtOrdem(liCount)
'            Load udClassificacao(liCount)
'
'            txtOrdem(liCount).Top = txtOrdem(liCount - 1).Top + txtOrdem(liCount - 1).Height + 60
'            lblFiltro(liCount).Top = txtOrdem(liCount).Top + 60
'            udClassificacao(liCount).Top = txtOrdem(liCount).Top
'        End If
'
'        lblFiltro(liCount).Visible = True
'        txtOrdem(liCount).Visible = True
'        udClassificacao(liCount).Visible = True
'        objArgControle.Visible = True
'
'        udClassificacao(liCount).Enabled = False
'        udClassificacao(liCount).Min = 0
'        udClassificacao(liCount).Max = objXMLMapaFiltro.documentElement.selectNodes( '                                       "//" & UserControl.Parent.Name & "/*").length
'
'        lblFiltro(liCount).Caption = objDomNode.selectSingleNode("Descricao").Text
'
'        lblFiltro(liCount).Tag = objDomNode.selectSingleNode("TagXML").Text
'        txtOrdem(liCount).Tag = objDomNode.selectSingleNode("TagCodigo").Text
'
'        objArgControle.Top = txtOrdem(liCount).Top
'        objArgControle.TabIndex = liCount
'
'        liCount = liCount + 1
'    Next
'
'    If Not objXMLMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(piSistemaSLCC, UserControl.Parent.Name)) Then
'        Call fgErroLoadXML(objXMLMapaNavegacao, App.EXEName, "ctlFiltro", "flInicializar")
'    End If
'
'    Call flCarregarCombos(objXMLMapaFiltro, objXMLMapaNavegacao)
'    Call flDefinirHeight
'
'    Set objMIU = Nothing
'    Set objXMLMapaFiltro = Nothing
'    Set objXMLMapaNavegacao = Nothing
'
'    Call flAplicarSettingsRegistry
'
'    Exit Function
'
'ErrorHandler:
'    Set objMIU = Nothing
'    Set objXMLMapaFiltro = Nothing
'    Set objXMLMapaNavegacao = Nothing
'
''    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmSubReservaResumo - flInit")
'
'End Function
'
'' Verifica se já existe algum Setting anterior da OCX registrado na máquina.
'' Caso positivo, assume este como default.
'Private Sub flAplicarSettingsRegistry()
'
'Dim objDomRegistry                          As MSXML2.DOMDocument40
'Dim lsRegistry                              As String
'Dim liIndControle                           As Integer
'Dim objDomNode                              As MSXML2.IXMLDOMNode
'
'    lsRegistry = GetSetting("OCX Filtro", UserControl.Parent.Name, "Settings")
'    If lsRegistry = vbNullString Then Exit Sub
'
'    Set objDomRegistry = CreateObject("MSXML2.DOMDocument.4.0")
'    If Not objDomRegistry.loadXML(lsRegistry) Then
'        Call fgErroLoadXML(objDomRegistry, App.EXEName, "ctlFiltro", "flAplicarSettingsRegistry")
'    End If
'
'    For Each objDomNode In objDomRegistry.documentElement.selectNodes("//Registry/*")
'        liIndControle = objDomNode.selectSingleNode("IndiceControle").Text
'        txtOrdem(liIndControle).Text = objDomNode.selectSingleNode("OrdemSelecao").Text
'
'        If Val(txtOrdem(liIndControle).Text) > 0 Then
'            udClassificacao(liIndControle).Enabled = True
'        End If
'
'        If objDomNode.selectSingleNode("TipoControle").Text = "ImageCombo" Then
''            cboFiltro(liIndControle).ListIndex = objDomNode.selectSingleNode("ConteudoControle").Text
'            cboFiltro(liIndControle).ComboItems(objDomNode.selectSingleNode("ConteudoControle").Text).Selected = True
'
'        ElseIf objDomNode.selectSingleNode("TipoControle").Text = "TextBox" Then
'            txtFiltro(liIndControle).Text = objDomNode.selectSingleNode("ConteudoControle").Text
'
'        ElseIf objDomNode.selectSingleNode("TipoControle").Text = "DTPicker" Then
'            dtpFiltro(liIndControle).Value = objDomNode.selectSingleNode("ConteudoControle").Text
'
'        End If
'    Next
'
'    Set objDomRegistry = Nothing
'
'End Sub
'
''Preenche os combos com as tabelas e campos especificados na propriedade Campos\Tabelas
'Private Sub flCarregarCombos(ByRef xmlDOMFiltro As MSXML2.DOMDocument40, '                             ByRef xmlDOMNavegacao As MSXML2.DOMDocument40)
'
'Dim liCount                                 As Integer
'Dim objDomFiltro                            As MSXML2.IXMLDOMNode
'Dim objDomNavegacao                         As MSXML2.IXMLDOMNode
'
'    For Each objDomFiltro In xmlDOMFiltro.documentElement.selectNodes("//" & UserControl.Parent.Name & "/*")
'        If Val(objDomFiltro.selectSingleNode("Tipo").Text) = enumTipoControle.ctlComboBox Then
''            cboFiltro(liCount).AddItem "<-- Todos -->"
'            cboFiltro(liCount).ComboItems.Add , "cod_0", "<-- Todos -->"
'
'            For Each objDomNavegacao In xmlDOMNavegacao.documentElement.selectNodes( '                                        "//Repeat_" & objDomFiltro.selectSingleNode("TagXML").Text & "/*")
'
'                cboFiltro(liCount).ComboItems.Add , "cod_" & '                            objDomNavegacao.selectSingleNode(objDomFiltro.selectSingleNode("TagCodigo").Text).Text, '                            objDomNavegacao.selectSingleNode(objDomFiltro.selectSingleNode("TagDescricao").Text).Text
'
''                cboFiltro(liCount).ItemData(cboFiltro(liCount).NewIndex) = ''                            objDomNavegacao.selectSingleNode(objDomFiltro.selectSingleNode("TagCodigo").Text).Text
'
'            Next
'
''            cboFiltro(liCount).ListIndex = 0
'            cboFiltro(liCount).ComboItems("cod_0").Selected = True
'        End If
'
'        liCount = liCount + 1
'    Next
'
'End Sub
'
'' Define a altura da OCX, de acordo com o número de controles existentes
'Private Sub flDefinirHeight()
'    UserControl.Height = cboFiltro(0).Height * (lcControlesFiltro.Count + 1) + tlbComandos.Height + lblTitulo.Height + 200
'    UserControl.Width = 5640
'End Sub
'
'Private Sub flGravarSettingsRegistry()
'
'Dim objDomAux                               As MSXML2.DOMDocument40
'Dim objDomRegistry                          As MSXML2.DOMDocument40
'Dim objControleFiltro                       As Object
'
'    Set objDomRegistry = CreateObject("MSXML2.DOMDocument.4.0")
'    Call fgAppendNode(objDomRegistry, "", "Registry", "")
'
'    For Each objControleFiltro In lcControlesFiltro
'        Set objDomAux = CreateObject("MSXML2.DOMDocument.4.0")
'
'        Call fgAppendNode(objDomAux, "", "Grupo_ControleFiltro", "")
'        Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "IndiceControle", objControleFiltro.Index)
'        Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "OrdemSelecao", txtOrdem(objControleFiltro.Index).Text)
'        Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "TipoControle", TypeName(objControleFiltro))
'
'        If TypeName(objControleFiltro) = "ImageCombo" Then
''            Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "ConteudoControle", objControleFiltro.ListIndex)
'            Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "ConteudoControle", objControleFiltro.SelectedItem.Key)
'
'        ElseIf TypeName(objControleFiltro) = "TextBox" Then
'            Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "ConteudoControle", objControleFiltro.Text)
'
'        ElseIf TypeName(objControleFiltro) = "DTPicker" Then
'            Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "ConteudoControle", objControleFiltro.Value)
'
'        End If
'
'        Call fgAppendXML(objDomRegistry, "Registry", objDomAux.xml)
'
'        Set objDomAux = Nothing
'    Next
'
'    Call SaveSetting("OCX Filtro", UserControl.Parent.Name, "Settings", objDomRegistry.xml)
'
'    Set objDomRegistry = Nothing
'
'End Sub
'
'' Retorna qual a ordem seguinte de seleção, quando o usuário decide utilizar um dos controles da ocx.
'Private Function flIndiceDoValor(plValor As Long, plMaiorNumero As Long) As Long
'
'Dim llIndex As Long
'Dim lbAchou As Boolean
'
'    lbAchou = False
'    llIndex = 0
'
'    While txtOrdem.Count > llIndex And plValor <> txtOrdem(llIndex).Text
'        llIndex = llIndex + 1
'
'        If llIndex >= txtOrdem.Count Then
'            flIndiceDoValor = -1
'            Exit Function
'        End If
'        'If llIndex > plMaiorNumero Then
'        'End If
'
'    Wend
'
'    flIndiceDoValor = llIndex
'
'End Function
'
''Monta XML com os filtros selecionados e devolve ao formulário correspondente
'Private Sub flMontarXMLFiltro()
'
'Dim objDomFiltro                            As MSXML2.DOMDocument40
'Dim objDomAux                               As MSXML2.DOMDocument40
'Dim objDomRegistry                          As MSXML2.DOMDocument40
'Dim objControleFiltro                       As Object
'Dim liCount                                 As Integer
''Dim liIndCombo                              As Integer
'Dim objComboItem                            As ComboItem
'
'    Set objDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
'    Call fgAppendNode(objDomFiltro, "", "Filtros", "")
'
''    For Each objControleFiltro In lcControlesFiltro
''        If TypeName(objControleFiltro) = "ImageCombo" Then
'''            If objControleFiltro.ListIndex <= 0 Then ' "Todos"
''            If objControleFiltro.ComboItems("cod_0").Selected Then ' "Todos"
''                Call fgAppendNode(objDomFiltro, "Filtros", ''                                                "Repeat_" & lblFiltro(objControleFiltro.Index).Tag, "")
''
'''                For liIndCombo = 1 To objControleFiltro.ListCount - 1
''                For Each objComboItem In objControleFiltro.ComboItems
''                    If objComboItem.Key <> "cod_0" Then
''                        Set objDomAux = CreateObject("MSXML2.DOMDocument.4.0")
''
''                        Call fgAppendNode(objDomAux, "", "Grupo_" & lblFiltro(objControleFiltro.Index).Tag, "")
''    '                    Call fgAppendNode(objDomAux, "Grupo_" & lblFiltro(objControleFiltro.Index).Tag, ''    '                                                 lblFiltro(objControleFiltro.Index).Tag, ''    '                                                 objControleFiltro.ItemData(liIndCombo))
''                        Call fgAppendNode(objDomAux, "Grupo_" & lblFiltro(objControleFiltro.Index).Tag, ''                                                     lblFiltro(objControleFiltro.Index).Tag, ''                                                     Mid$(objComboItem.Key, 5))
''
''                        Call fgAppendXML(objDomFiltro, "Repeat_" & lblFiltro(objControleFiltro.Index).Tag, objDomAux.xml)
''
''                        Set objDomAux = Nothing
''                    End If
''                Next
''            End If
''        End If
''    Next objControleFiltro
'
'    For liCount = 1 To lcControlesFiltro.Count
'        For Each objControleFiltro In lcControlesFiltro
'            If txtOrdem(objControleFiltro.Index).Text = liCount Then
'                Call fgAppendNode(objDomFiltro, "Filtros", '                                                "Repeat_" & lblFiltro(liCount - 1).Tag, "")
'                Call fgAppendNode(objDomFiltro, "Repeat_" & lblFiltro(liCount - 1).Tag, '                                                "Grupo_" & lblFiltro(liCount - 1).Tag, "")
'
'                If TypeName(objControleFiltro) = "ImageCombo" Then
''                    Call fgAppendNode(objDomFiltro, "Grupo_" & lblFiltro(liCount - 1).Tag, ''                                                    lblFiltro(liCount - 1).Tag, ''                                                    objControleFiltro.ItemData(objControleFiltro.ListIndex))
'                    Call fgAppendNode(objDomFiltro, "Grupo_" & lblFiltro(liCount - 1).Tag, '                                                    lblFiltro(liCount - 1).Tag, '                                                    Mid$(objControleFiltro.SelectedItem.Key, 5))
'
'                ElseIf TypeName(objControleFiltro) = "TextBox" Then
'                    Call fgAppendNode(objDomFiltro, "Grupo_" & lblFiltro(liCount - 1).Tag, '                                                    lblFiltro(liCount - 1).Tag, '                                                    objControleFiltro.Text)
'
'                ElseIf TypeName(objControleFiltro) = "DTPicker" Then
'                    Call fgAppendNode(objDomFiltro, "Grupo_" & lblFiltro(liCount - 1).Tag, '                                                    lblFiltro(liCount - 1).Tag, '                                                    objControleFiltro.Value)
'
'                End If
'
'                Exit For
'            End If
'        Next objControleFiltro
'    Next liCount
'
'    xmlDocFiltros = objDomFiltro.xml
'
'    Set objDomFiltro = Nothing
'
'End Sub
'
'' Verifica qual é o primeiro controle selecionado (se existir).
'' Monta XML com o conteúdo, e devolve ao formulário correspondente.
'Private Sub flMontarXMLPrimeiroControleSelecionado()
'
'Dim objDomControle                          As MSXML2.DOMDocument40
'Dim objDomAux                               As MSXML2.DOMDocument40
'Dim objControleFiltro                       As Object
'Dim liIndCombo                              As Integer
'
'    Set objDomControle = CreateObject("MSXML2.DOMDocument.4.0")
'    Call fgAppendNode(objDomControle, "", "PrimeiraSelecao", "")
'
'    lsTituloTableCombo = vbNullString
'    xmlDocPrimSelecao = vbNullString
'
'    For Each objControleFiltro In lcControlesFiltro
'        If txtOrdem(objControleFiltro.Index).Text = "1" Then
'            If TypeName(objControleFiltro) = "ImageCombo" Then
'                lsTituloTableCombo = objControleFiltro.Text
'
''                For liIndCombo = 1 To objControleFiltro.ListCount - 1
'                For liIndCombo = 2 To objControleFiltro.ComboItems.Count
'
'                    Set objDomAux = CreateObject("MSXML2.DOMDocument.4.0")
'
'                    Call fgAppendNode(objDomAux, "", "Grupo_ControleFiltro", "")
''                    Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "Codigo", objControleFiltro.ItemData(liIndCombo))
''                    Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "Descricao", objControleFiltro.List(liIndCombo))
'                    Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "Codigo", Mid$(objControleFiltro.ComboItems(liIndCombo).Key, 5))
'                    Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "Descricao", objControleFiltro.ComboItems(liIndCombo).Text)
'
'                    Call fgAppendXML(objDomControle, "PrimeiraSelecao", objDomAux.xml)
'                    Set objDomAux = Nothing
'
'                Next
'
'            ElseIf TypeName(objControleFiltro) = "TextBox" Then
'                lsTituloTableCombo = objControleFiltro.Text
'
'                Set objDomAux = CreateObject("MSXML2.DOMDocument.4.0")
'
'                Call fgAppendNode(objDomAux, "", "Grupo_ControleFiltro", "")
'                Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "Codigo", "0")
'                Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "Descricao", objControleFiltro.Text)
'
'                Call fgAppendXML(objDomControle, "PrimeiraSelecao", objDomAux.xml)
'                Set objDomAux = Nothing
'
'            ElseIf TypeName(objControleFiltro) = "DTPicker" Then
'                lsTituloTableCombo = objControleFiltro.Value
'
'                Set objDomAux = CreateObject("MSXML2.DOMDocument.4.0")
'
'                Call fgAppendNode(objDomAux, "", "Grupo_ControleFiltro", "")
'                Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "Codigo", "0")
'                Call fgAppendNode(objDomAux, "Grupo_ControleFiltro", "Descricao", objControleFiltro.Value)
'
'                Call fgAppendXML(objDomControle, "PrimeiraSelecao", objDomAux.xml)
'                Set objDomAux = Nothing
'
'            End If
'
'            xmlDocPrimSelecao = objDomControle.xml
'
'            Exit For
'        End If
'    Next
'
'    Set objDomControle = Nothing
'
'End Sub
'
''Ordena a sequencia de classificação do filtro
'Private Sub flOrdenarClassificacao(ByVal UpDownIndex As Integer, '                                   ByVal OrdemMovimentacao As enumMovimentacao)
'
'Dim liCount                                 As Integer
'Dim liNumeroOriginal                        As Integer
'Dim llIndiceTroca                           As Long
'Dim liMaiorNumero                           As Long
'
'    For liCount = 0 To txtOrdem.Count - 1
'        If liMaiorNumero < Val(txtOrdem(liCount).Text) Then
'            liMaiorNumero = Val(txtOrdem(liCount).Text)
'        End If
'    Next
'
'    liNumeroOriginal = Val(txtOrdem(UpDownIndex).Text)
'
'    If OrdemMovimentacao = Anterior Then
'        If txtOrdem(UpDownIndex).Text = "1" Then Exit Sub
'
'        llIndiceTroca = flIndiceDoValor(liNumeroOriginal - 1, liMaiorNumero)
'        If llIndiceTroca >= 0 Then
'            txtOrdem(UpDownIndex).Text = liNumeroOriginal - 1
'            txtOrdem(llIndiceTroca) = liNumeroOriginal
'        End If
'    Else
'        If txtOrdem(UpDownIndex).Text = udClassificacao(UpDownIndex).Max Then Exit Sub
'
'        llIndiceTroca = flIndiceDoValor(liNumeroOriginal + 1, liMaiorNumero)
'        If llIndiceTroca >= 0 Then
'            txtOrdem(UpDownIndex).Text = liNumeroOriginal + 1
'            txtOrdem(llIndiceTroca) = liNumeroOriginal
'        End If
'    End If
'
'End Sub
'
''Subtrai 1 dos outros índices de classificação quando o usuário seleciona "Todos" no Combo,
''ou limpa os argumentos de pesquisa
'Private Sub flSubtrairIndiceClassificacao(ByVal Index As Integer)
'
'Dim liCount                                 As Integer
'
'    If txtOrdem(Index).Text = "0" Then Exit Sub
'
'    For liCount = 0 To txtOrdem.Count - 1
'        If Index <> liCount Then
'            If txtOrdem(liCount).Text > txtOrdem(Index).Text Then
'                txtOrdem(liCount).Text = Val(txtOrdem(liCount).Text) - 1
'            End If
'        End If
'    Next
'
'    txtOrdem(Index).Text = "0"
'    udClassificacao(Index).Enabled = False
'
'End Sub
'
''Classifica a sequencia de classsificação do filtro
'Private Sub flValidarClassificação(ByVal Index As Integer)
'
'Dim liCount                                 As Integer
'Dim liMaiorNumero                           As Integer
'
'    liMaiorNumero = 0
'
'    For liCount = 0 To txtOrdem.Count - 1
'        If Index <> liCount Then
'            If liMaiorNumero < Val(txtOrdem(liCount).Text) Then
'                liMaiorNumero = Val(txtOrdem(liCount).Text)
'            End If
'        End If
'    Next
'
'    If Val(txtOrdem(Index).Text) = 0 Then
'        txtOrdem(Index).Text = Val(liMaiorNumero) + 1
'        udClassificacao(Index).Enabled = True
'    End If
'
'End Sub
'
'Private Sub cboFiltro_Click(Index As Integer)
''    If cboFiltro(Index).ListIndex <> 0 Then
'    If Not cboFiltro(Index).ComboItems("cod_0").Selected Then
'        Call flValidarClassificação(Index)
'    Else
'        Call flSubtrairIndiceClassificacao(Index)
'    End If
'End Sub
'
'Private Sub cboFiltro_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = 0
'End Sub
'
'Private Sub dtpFiltro_Change(Index As Integer)
'    If IsNull(dtpFiltro(Index).Value) Then
'        Call flSubtrairIndiceClassificacao(Index)
'    Else
'        Call flValidarClassificação(Index)
'    End If
'End Sub
'
'Private Sub lblFechar_Click()
'    RaiseEvent CancelarFiltro
'End Sub
'
'Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)
'    Select Case Button.Index
'        Case 1 'AplicaFiltro
'            Call flMontarXMLFiltro
'            Call flMontarXMLPrimeiroControleSelecionado
'            Call flGravarSettingsRegistry
'            RaiseEvent AplicarFiltro(xmlDocFiltros, lsTituloTableCombo, xmlDocPrimSelecao)
'        Case 2 'Cancelar Filtro
'            RaiseEvent CancelarFiltro
'    End Select
'End Sub
'
'Private Sub txtFiltro_Change(Index As Integer)
'    If txtFiltro(Index).Text = vbNullString Then
'        Call flSubtrairIndiceClassificacao(Index)
'    Else
'        Call flValidarClassificação(Index)
'    End If
'End Sub
'
'Private Sub txtFiltro_KeyPress(Index As Integer, KeyAscii As Integer)
'    If txtFiltro(Index).Tag = enumTipoTexto.Numerico Then
'        If Not IsNumeric(Chr$(KeyAscii)) And KeyAscii <> Asc(vbBack) Then
'            KeyAscii = 0
'            Beep
'        End If
'    End If
'End Sub
'
'Public Function ObterFiltro() As String
'    ObterFiltro = xmlDocFiltros
'End Function
'
'Private Sub udClassificacao_DownClick(Index As Integer)
'    Call flOrdenarClassificacao(Index, enumMovimentacao.Anterior)
'End Sub
'
'Private Sub udClassificacao_UpClick(Index As Integer)
'    Call flOrdenarClassificacao(Index, enumMovimentacao.Proximo)
'End Sub
'
'Private Sub UserControl_Resize()
'    tlbComandos.Top = UserControl.Height - tlbComandos.Height
'End Sub
'
'Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyUp(KeyCode, Shift)
'End Sub
'
'Private Sub UserControl_KeyPress(KeyAscii As Integer)
'    RaiseEvent KeyPress(KeyAscii)
'End Sub
'
'Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyDown(KeyCode, Shift)
'End Sub
