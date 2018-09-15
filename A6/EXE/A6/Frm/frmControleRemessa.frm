VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmControleRemessa 
   Caption         =   "Sub-reserva - Controle de Remessa"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   9540
   WindowState     =   2  'Maximized
   Begin A6.TableCombo ctlTableCombo 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   661
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   5070
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
            Picture         =   "frmControleRemessa.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleRemessa.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleRemessa.frx":0224
            Key             =   "showtreeview2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleRemessa.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleRemessa.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleRemessa.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleRemessa.frx":0F6C
            Key             =   "anterior"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControleRemessa.frx":13BE
            Key             =   "posterior"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxMonitoracao 
      Height          =   3495
      Left            =   60
      TabIndex        =   0
      Tag             =   "Controle Remessa"
      Top             =   495
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   10
      Cols            =   8
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   -2147483638
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      BorderStyle     =   0
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   3960
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   635
      ButtonWidth     =   2328
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
            Caption         =   "Definir Filtro"
            Key             =   "showfiltro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Label lblGrupoVeiculo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4110
   End
End
Attribute VB_Name = "frmControleRemessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário a consulta aos registros de controle de remessas enviadas
' ao A6.

Option Explicit

'Este objeto ObjDomDocument é carregado com as propriedades para o formulário
' e todas as coleções que este for utilizar
Private xmlMapaNavegacao                    As MSXML2.DOMDocument40

'Variaveis para a utilização do Filtro
Private strFiltroXML                        As String
Private blnPrimeiraConsulta                 As Boolean

Private Const strFuncionalidade             As String = "CONTROLEREMESSA"
Private Const strTableComboInicial          As String = "Empresa"

Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1

'Definição de Colunas e Linha do Grid
Private Const COL_DESCRICAO                 As Integer = 0

'Definição das Colunas do Grid
Private Const COL_SIGLA_SISTEMA             As Integer = 0
Private Const COL_DESCR_SISTEMA             As Integer = 1
Private Const COL_CO_VEIC_LEGA              As Integer = 2
Private Const COL_NO_VEIC_LEGA              As Integer = 3
Private Const COL_CODIGO_EMPRESA            As Integer = 4
Private Const COL_NOME_EMPRESA              As Integer = 5
Private Const COL_DATA_MENSAGEM             As Integer = 6
Private Const COL_TIPO_MENSAGEM             As Integer = 7
Private Const COL_QUANT_REG_INFO_SIST_ORI   As Integer = 8
Private Const COL_QUANT_REG_REJEITADO       As Integer = 9
Private Const COL_QUANT_REG_RECEBIDO        As Integer = 10
Private Const COL_DH_INI_REMESSA_ORIGEM     As Integer = 11
Private Const COL_DH_FIM_REMESSA_ORIGEM     As Integer = 12
Private Const COL_DH_INI_PROCESS_REMESSA    As Integer = 13
Private Const COL_DH_FIM_PROCESS_REMESSA    As Integer = 14
Private Const COL_CODIGO_REMESSA_PRIMEIRO   As Integer = 15
Private Const COL_CODIGO_REMESSA_ULTIMO     As Integer = 16
Private Const COL_VALOR_TOTAL_REMESSA       As Integer = 17
Private Const COL_CODIGO_SITUACAO_REMESSA   As Integer = 18
Private Const COL_DH_ULTIMA_ATUALIZACAO     As Integer = 19

Private Const MAX_FIXED_ROWS                As Integer = 100

' Retorna a descrição do tipo de mensagem.

Private Function flTipoMensagem(ByVal penumTipo As enumTipoMensagem) As String

Dim strRetornaMensagem                      As String

On Error GoTo ErrorHandler

Select Case penumTipo
       Case MensagemXML
            strRetornaMensagem = "MensagemXML"
       Case MensagemString
            strRetornaMensagem = "MensagemString"
       Case MensagemCSV
            strRetornaMensagem = "MensagemCSV"
       Case MensagemStringXML
            strRetornaMensagem = "MensagemStringXML"
       Case MensagemCSVXML
            strRetornaMensagem = "MensagemCSVXML"
End Select

flTipoMensagem = strRetornaMensagem

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flTipoMensagem", 0

End Function

' Carrega lista com o resultado da pesquisa dos registros de controle de remessa.

Private Sub flCarregarFlexGrid(ByRef pstrXMLDocFiltros As String)

#If EnableSoap = 1 Then
    Dim objControleRemessa  As MSSOAPLib30.SoapClient30
#Else
    Dim objControleRemessa  As A6MIU.clsConsultaControleRemessa
#End If

Dim xmlRemessa              As MSXML2.DOMDocument40
Dim xmlDomNode              As MSXML2.IXMLDOMNode
Dim lngLinhaGrid            As Long
Dim strXMLRetorno           As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    gintRowPositionAnt = 0

    Set objControleRemessa = fgCriarObjetoMIU("A6MIU.clsConsultaControleRemessa")

    Set xmlRemessa = CreateObject("MSXML2.DOMDocument.4.0")
    
    strXMLRetorno = objControleRemessa.LerTodos(pstrXMLDocFiltros, _
                                                vntCodErro, _
                                                vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Call flFormatarFlxMonitoracao
    
    'Caso a tabela esteja sem registros não tem como carregar um XML, sendo assim vai para o fim da rotina.
    If Trim(strXMLRetorno) <> "" Then
       If Not xmlRemessa.loadXML(strXMLRetorno) Then
          '100 - Documento XML Inválido.
          lngCodigoErroNegocio = 100
          GoTo ErrorHandler
       End If
    Else
       Exit Sub
    End If
    
    lngLinhaGrid = 2

    With Me.flxMonitoracao
    
        .Redraw = False
        
        .MergeCol(0) = False
        .MergeCol(1) = False
        .MergeCol(2) = False
        .MergeCol(3) = False
        .MergeCol(4) = False
        .MergeCol(5) = False
        .MergeCol(6) = False
        .MergeCol(7) = False
        .MergeCol(8) = False
        .MergeCol(9) = False
        .MergeCol(10) = False
        .MergeCol(11) = False
        .MergeCol(12) = False
        .MergeCol(13) = False
        .MergeCol(14) = False
        .MergeCol(15) = False
        .MergeCol(16) = False
        .MergeCol(17) = False
        .MergeCol(18) = False
        .MergeCol(19) = False
        
        'Define alinhamentos colunas
        .ColAlignment(COL_CODIGO_SITUACAO_REMESSA) = flexAlignRightBottom
        .ColAlignment(COL_CODIGO_REMESSA_ULTIMO) = flexAlignRightBottom
        .ColAlignment(COL_CODIGO_REMESSA_PRIMEIRO) = flexAlignRightBottom
        .ColAlignment(COL_QUANT_REG_INFO_SIST_ORI) = flexAlignRightBottom
        .ColAlignment(COL_QUANT_REG_REJEITADO) = flexAlignRightBottom
        .ColAlignment(COL_QUANT_REG_RECEBIDO) = flexAlignRightBottom
        .ColAlignment(COL_CODIGO_EMPRESA) = flexAlignRightBottom
        .ColAlignment(COL_VALOR_TOTAL_REMESSA) = flexAlignRightBottom
        .ColAlignment(COL_DESCR_SISTEMA) = flexAlignLeftBottom
        .ColAlignment(COL_NOME_EMPRESA) = flexAlignLeftBottom
        .ColAlignment(COL_TIPO_MENSAGEM) = flexAlignLeftBottom
        .ColAlignment(COL_CO_VEIC_LEGA) = flexAlignLeftBottom
        .ColAlignment(COL_NO_VEIC_LEGA) = flexAlignLeftBottom
        
        .ColWidth(COL_SIGLA_SISTEMA) = 600
        .ColWidth(COL_DESCR_SISTEMA) = 1000
        .ColWidth(COL_CODIGO_EMPRESA) = 600
        .ColWidth(COL_NOME_EMPRESA) = 2500
        .ColWidth(COL_DATA_MENSAGEM) = 1000
        .ColWidth(COL_QUANT_REG_INFO_SIST_ORI) = 900
        .ColWidth(COL_QUANT_REG_REJEITADO) = 900
        .ColWidth(COL_QUANT_REG_RECEBIDO) = 900
        .ColWidth(COL_DH_INI_REMESSA_ORIGEM) = 1700
        .ColWidth(COL_DH_FIM_REMESSA_ORIGEM) = 1700
        .ColWidth(COL_DH_INI_PROCESS_REMESSA) = 1700
        .ColWidth(COL_DH_FIM_PROCESS_REMESSA) = 1700
        .ColWidth(COL_DH_ULTIMA_ATUALIZACAO) = 1700
        .ColWidth(COL_CO_VEIC_LEGA) = 1000
        .ColWidth(COL_NO_VEIC_LEGA) = 3000
        
        .FillStyle = flexFillRepeat
        .Row = 0
        .Col = 0
        .RowSel = 1
        .ColSel = .Cols - 1
        .CellAlignment = flexAlignCenterBottom
        
        For Each xmlDomNode In xmlRemessa.documentElement.selectNodes("//Repeat_Sistema/*")
                            
            'Finalizar a Mesclagem de Colunas e Linhas
            .MergeRow(lngLinhaGrid) = False
                
            .TextMatrix(lngLinhaGrid, COL_SIGLA_SISTEMA) = xmlDomNode.selectSingleNode("SG_SIST").Text
            .TextMatrix(lngLinhaGrid, COL_DESCR_SISTEMA) = xmlDomNode.selectSingleNode("NO_SIST").Text
            .TextMatrix(lngLinhaGrid, COL_CO_VEIC_LEGA) = xmlDomNode.selectSingleNode("CO_VEIC_LEGA").Text
            .TextMatrix(lngLinhaGrid, COL_NO_VEIC_LEGA) = xmlDomNode.selectSingleNode("NO_VEIC_LEGA").Text
            .TextMatrix(lngLinhaGrid, COL_CODIGO_EMPRESA) = xmlDomNode.selectSingleNode("CO_EMPR").Text
            .TextMatrix(lngLinhaGrid, COL_NOME_EMPRESA) = xmlDomNode.selectSingleNode("NO_EMPR").Text
            .TextMatrix(lngLinhaGrid, COL_DATA_MENSAGEM) = fgDtXML_To_Interface(xmlDomNode.selectSingleNode("DT_REME").Text)
            .TextMatrix(lngLinhaGrid, COL_TIPO_MENSAGEM) = xmlDomNode.selectSingleNode("NO_TIPO_MESG").Text
            .TextMatrix(lngLinhaGrid, COL_QUANT_REG_INFO_SIST_ORI) = xmlDomNode.selectSingleNode("QT_REGT_INFO_SIST_ORIG").Text
            .TextMatrix(lngLinhaGrid, COL_QUANT_REG_REJEITADO) = xmlDomNode.selectSingleNode("QT_REGT_REJE").Text
            .TextMatrix(lngLinhaGrid, COL_QUANT_REG_RECEBIDO) = xmlDomNode.selectSingleNode("QT_REGT_RECB").Text
            .TextMatrix(lngLinhaGrid, COL_DH_INI_REMESSA_ORIGEM) = fgDtHrStr_To_DateTime(xmlDomNode.selectSingleNode("DH_INIC_REME_ORIG").Text)
            
            If xmlDomNode.selectSingleNode("DH_FIM_REME_ORIG").Text <> gstrDataVazia Then
                .TextMatrix(lngLinhaGrid, COL_DH_FIM_REMESSA_ORIGEM) = fgDtHrStr_To_DateTime(xmlDomNode.selectSingleNode("DH_FIM_REME_ORIG").Text)
            End If
            
            .TextMatrix(lngLinhaGrid, COL_CODIGO_REMESSA_PRIMEIRO) = xmlDomNode.selectSingleNode("CO_REME_PRMR").Text
            .TextMatrix(lngLinhaGrid, COL_DH_INI_PROCESS_REMESSA) = fgDtHrStr_To_DateTime(xmlDomNode.selectSingleNode("DH_INIC_PROC_REME").Text)
            .TextMatrix(lngLinhaGrid, COL_CODIGO_REMESSA_ULTIMO) = xmlDomNode.selectSingleNode("CO_REME_ULTI").Text
            .TextMatrix(lngLinhaGrid, COL_VALOR_TOTAL_REMESSA) = fgVlrXml_To_Interface(xmlDomNode.selectSingleNode("VA_TOTL_REME").Text)
            
            If xmlDomNode.selectSingleNode("DH_FIM_PROC_REME").Text <> gstrDataVazia Then
               .TextMatrix(lngLinhaGrid, COL_DH_FIM_PROCESS_REMESSA) = fgDtHrStr_To_DateTime(xmlDomNode.selectSingleNode("DH_FIM_PROC_REME").Text)
            End If
            
            If Val(xmlDomNode.selectSingleNode("CO_SITU_REME").Text) = enumSituacaoRemessa.EmProcessamento Then
                .TextMatrix(lngLinhaGrid, COL_CODIGO_SITUACAO_REMESSA) = "Em Processamento"
            Else
                .TextMatrix(lngLinhaGrid, COL_CODIGO_SITUACAO_REMESSA) = "Finalizado"
            End If
            
            .TextMatrix(lngLinhaGrid, COL_DH_ULTIMA_ATUALIZACAO) = fgDtHrStr_To_DateTime(xmlDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)
            
            If (.Rows - 2) > lngLinhaGrid Then
               lngLinhaGrid = lngLinhaGrid + 1
            Else
               lngLinhaGrid = lngLinhaGrid + 1
               .Rows = .Rows + 1
            End If
            
        Next xmlDomNode
        
        .Rows = .Rows - 2
        .Redraw = True
        
    End With

    Set xmlRemessa = Nothing
    Set objControleRemessa = Nothing
    
    Exit Sub

ErrorHandler:
    Me.flxMonitoracao.Redraw = True
    
    Set xmlRemessa = Nothing
    Set objControleRemessa = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    mdiSBR.uctLogErros.MostrarErros Err, "flCarregaFlexGrid"

End Sub

Private Sub ctlTableCombo_AplicarFiltro(xmlDocFiltros As String)

Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    fgCursor True
    
    If Trim(strFiltroXML) = vbNullString Then
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomLeitura.loadXML(xmlDocFiltros)
        
        Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
        
        Call fgAppendNode(xmlDomFiltro, "", "Repeat_Filtros", "")
        Call fgAppendNode(xmlDomFiltro, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
        Call fgAppendNode(xmlDomFiltro, "Grupo_BancoLiquidante", "BancoLiquidante", _
                                        xmlDomLeitura.selectSingleNode("//BancoLiquidante").Text)
                                        
        If xmlDomFiltro.selectSingleNode("//Grupo_Data") Is Nothing Then
            fgAppendNode xmlDomFiltro, "Repeat_Filtros", "Grupo_Data", ""
            fgAppendNode xmlDomFiltro, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)))
            fgAppendNode xmlDomFiltro, "Grupo_Data", "DataFim", fgDtXML_To_Oracle("99991231")
        End If
        
        Call flCarregarFlexGrid(xmlDomFiltro.xml)
        strFiltroXML = xmlDomFiltro.xml
        
    Else
        Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomFiltro.loadXML(strFiltroXML)
        
        If Not xmlDomFiltro.selectSingleNode("//Grupo_BancoLiquidante") Is Nothing Then
            Call fgRemoveNode(xmlDomFiltro, "Grupo_BancoLiquidante")
        End If
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomLeitura.loadXML(xmlDocFiltros)
        
        Call fgAppendNode(xmlDomFiltro, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
        Call fgAppendNode(xmlDomFiltro, "Grupo_BancoLiquidante", "BancoLiquidante", _
                                        xmlDomLeitura.selectSingleNode("//BancoLiquidante").Text)
                                        
        If xmlDomFiltro.selectSingleNode("//Grupo_Data") Is Nothing Then
            fgAppendNode xmlDomFiltro, "Repeat_Filtros", "Grupo_Data", ""
            fgAppendNode xmlDomFiltro, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)))
            fgAppendNode xmlDomFiltro, "Grupo_Data", "DataFim", fgDtXML_To_Oracle("99991231")
        End If
        
        Call flCarregarFlexGrid(xmlDomFiltro.xml)
        strFiltroXML = xmlDomFiltro.xml
        
    End If
    
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing
    fgCursor

    Exit Sub

ErrorHandler:
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing

    mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - ctlTableCombo_AplicarFiltro"
    
End Sub

Private Sub ctlTableCombo_DropDown()

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgCarregarCombo
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaD0 - ctlTableCombo_DropDown"

End Sub

Private Sub ctlTableCombo_MouseMove()

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButton3D
    
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "ctlTableCombo_MouseMove"

End Sub

Private Sub flxMonitoracao_Click()

On Error GoTo ErrorHandler

    fgPositionRowFlexGrid flxMonitoracao.Row, flxMonitoracao
    
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "flxMonitoracao_Click"

End Sub

Private Sub flxMonitoracao_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat
    
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "flxMonitoracao_MouseMove"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        If Trim(strFiltroXML) = vbNullString Then Exit Sub
        
        fgCursor True
        Call flCarregarFlexGrid(strFiltroXML)
        fgCursor
    End If

    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmControleRemessa - Form_KeyDown"

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCenterMe Me
     
    Me.Icon = mdiSBR.Icon
     
    Me.Show
    DoEvents

    fgCursor True
    
    blnPrimeiraConsulta = True
    ctlTableCombo.TituloCombo = strTableComboInicial
    
    flFormatarFlxMonitoracao

    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA6.frmControleRemessa
    Load objFiltro
    
    Call objFiltro.fgCarregarPesquisaAnterior
    objFiltro.Show vbModal
    
    blnPrimeiraConsulta = False
    fgCursor
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "Form_Load"

End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    With Me
        .lblGrupoVeiculo.Left = 0
        .lblGrupoVeiculo.Top = 0
        .lblGrupoVeiculo.Width = .ScaleWidth
        
        .ctlTableCombo.Left = 0
        .ctlTableCombo.Top = 0
        
        .flxMonitoracao.Left = 0
        .flxMonitoracao.Top = .lblGrupoVeiculo.Height
        .flxMonitoracao.Width = .ScaleWidth
        .flxMonitoracao.Height = .ScaleHeight - .flxMonitoracao.Top - .tlbButtons.Height
    End With

End Sub

' Formatação inicial do grid de pesquisa.

Private Sub flFormatarFlxMonitoracao()

Dim intCount                                 As Integer
Dim intLinhaGrid                             As Integer
Dim intLinhaGridFix                          As Integer

On Error GoTo ErrorHandler

    intLinhaGridFix = 0
    intLinhaGrid = 1
    
    With Me.flxMonitoracao
        
        .Clear
        .Redraw = False
        
        .Rows = MAX_FIXED_ROWS
        .Cols = 20
        .FixedRows = 2
        
        For intCount = 0 To .Cols - 1
            .ColAlignment(intCount) = MSFlexGridLib.flexAlignCenterCenter
            .ColWidth(intCount) = 1600
        Next
        
        .GridColorFixed = &HE6E6E6

        'EMPRESA - Coluna Fixa
        .TextMatrix(intLinhaGridFix, COL_CODIGO_EMPRESA) = "Empresa"
        .TextMatrix(intLinhaGridFix, COL_NOME_EMPRESA) = "Empresa"
        
        .TextMatrix(intLinhaGridFix, COL_SIGLA_SISTEMA) = "Sistema"
        .TextMatrix(intLinhaGridFix, COL_DESCR_SISTEMA) = "Sistema"
        
        .TextMatrix(intLinhaGridFix, COL_CO_VEIC_LEGA) = "Veículo Legal"
        .TextMatrix(intLinhaGridFix, COL_NO_VEIC_LEGA) = "Veículo Legal"
        
        'Veiculo Legal
        .TextMatrix(intLinhaGrid, COL_CO_VEIC_LEGA) = "Codigo"
        .TextMatrix(intLinhaGrid, COL_NO_VEIC_LEGA) = "Descrição"
        
        .TextMatrix(intLinhaGridFix, COL_DATA_MENSAGEM) = "Controle"
        .TextMatrix(intLinhaGridFix, COL_TIPO_MENSAGEM) = "Controle"

        'Controle
        .TextMatrix(intLinhaGrid, COL_CODIGO_EMPRESA) = "Codigo"
        .TextMatrix(intLinhaGrid, COL_NOME_EMPRESA) = "Descrição"
        
        .TextMatrix(intLinhaGrid, COL_SIGLA_SISTEMA) = "Sigla"
        .TextMatrix(intLinhaGrid, COL_DESCR_SISTEMA) = "Descrição"
        
        .TextMatrix(intLinhaGrid, COL_DATA_MENSAGEM) = "Data"
        .TextMatrix(intLinhaGrid, COL_TIPO_MENSAGEM) = "Tipo"
        
        'Quantidade Registros - Coluna Fixa
        .TextMatrix(intLinhaGridFix, COL_QUANT_REG_INFO_SIST_ORI) = "Quantidade Registros"
        .TextMatrix(intLinhaGridFix, COL_QUANT_REG_REJEITADO) = "Quantidade Registros"
        .TextMatrix(intLinhaGridFix, COL_QUANT_REG_RECEBIDO) = "Quantidade Registros"
        
        'Quantidade Registros
        .TextMatrix(intLinhaGrid, COL_QUANT_REG_INFO_SIST_ORI) = "Informado"
        .TextMatrix(intLinhaGrid, COL_QUANT_REG_REJEITADO) = "Rejeitado"
        .TextMatrix(intLinhaGrid, COL_QUANT_REG_RECEBIDO) = "Recebido"
        
        'Data Remessa - Coluna Fixa
        .TextMatrix(intLinhaGridFix, COL_DH_INI_REMESSA_ORIGEM) = "Data Remessa"
        .TextMatrix(intLinhaGridFix, COL_DH_FIM_REMESSA_ORIGEM) = "Data Remessa"
        
        'Data Remessa
        .TextMatrix(intLinhaGrid, COL_DH_INI_REMESSA_ORIGEM) = "Inicial"
        .TextMatrix(intLinhaGrid, COL_DH_FIM_REMESSA_ORIGEM) = "Final"
        
        'Codigo Remessa - Coluna Fixa
        .TextMatrix(intLinhaGridFix, COL_CODIGO_REMESSA_PRIMEIRO) = "Codigo Remessa"
        .TextMatrix(intLinhaGridFix, COL_CODIGO_REMESSA_ULTIMO) = "Codigo Remessa"
        
        'Codigo Remessa
        .TextMatrix(intLinhaGrid, COL_CODIGO_REMESSA_PRIMEIRO) = "Primeiro"
        .TextMatrix(intLinhaGrid, COL_CODIGO_REMESSA_ULTIMO) = "Ultimo"
        
        'Processo - Coluna Fixa
        .TextMatrix(intLinhaGridFix, COL_DH_INI_PROCESS_REMESSA) = "Processo"
        .TextMatrix(intLinhaGridFix, COL_DH_FIM_PROCESS_REMESSA) = "Processo"
        
        'Processo
        .TextMatrix(intLinhaGrid, COL_DH_INI_PROCESS_REMESSA) = "Inicio"
        .TextMatrix(intLinhaGrid, COL_DH_FIM_PROCESS_REMESSA) = "Final"
                    
        'Remessa - Coluna Fixa
        .TextMatrix(intLinhaGridFix, COL_CODIGO_SITUACAO_REMESSA) = "Remessa"
        .TextMatrix(intLinhaGridFix, COL_VALOR_TOTAL_REMESSA) = "Remessa"
        .TextMatrix(intLinhaGridFix, COL_DH_ULTIMA_ATUALIZACAO) = "Remessa"
        
        'Remessa
        .TextMatrix(intLinhaGrid, COL_CODIGO_SITUACAO_REMESSA) = "Situação"
        .TextMatrix(intLinhaGrid, COL_VALOR_TOTAL_REMESSA) = "Valor Total"
        .TextMatrix(intLinhaGrid, COL_DH_ULTIMA_ATUALIZACAO) = "Atualizado em:"

        .MergeCells = flexMergeFree

        .MergeRow(0) = True
        
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
        .MergeCol(6) = True
        .MergeCol(7) = True
        .MergeCol(8) = True
        .MergeCol(9) = True
        .MergeCol(10) = True
        .MergeCol(11) = True
        .MergeCol(12) = True
        .MergeCol(13) = True
        .MergeCol(14) = True
        .MergeCol(15) = True
        .MergeCol(16) = True
        .MergeCol(17) = True
        .MergeCol(18) = True
        .MergeCol(19) = True
        
        .GridLinesFixed = flexGridInset
        
        .Redraw = True
        
    End With

Exit Sub
ErrorHandler:
   
End Sub

' Carrega configurações iniciais do formulário.

Private Sub flInit()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim strMapaNavegacao    As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    Set xmlMapaNavegacao = Nothing

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    'strMapaNavegacao = objMIU.ObterMapaNavegacao(strFuncionalidade, vntCodErro, vntMensagemErro)
    Set objMIU = Nothing
    
    'If vntCodErro <> 0 Then
    '    GoTo ErrorHandler
    'End If

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmControleRemessa", "flInit")
    End If

    Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    'Err.Number = vntCodErro
    'Err.Description = vntMensagemErro
    
    Call mdiSBR.uctLogErros.MostrarErros(Err, "flInit")

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmControleRemessa = Nothing
    gintRowPositionAnt = 0
    
End Sub

Private Sub lblGrupoVeiculo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmControleRemessa - lblGrupoVeiculo_MouseMove", Me.Caption
    
End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, strTituloTableCombo As String)

On Error GoTo ErrorHandler

    strFiltroXML = xmlDocFiltros
    
    If Trim(strFiltroXML) <> vbNullString Then
        If fgMostraFiltro(strFiltroXML, blnPrimeiraConsulta) Then
            If blnPrimeiraConsulta Then
                blnPrimeiraConsulta = False
                
                'Call tlbButtons_ButtonClick(tlbButtons.Buttons("showfiltro"))
            Else
                frmMural.Caption = Me.Caption
                frmMural.Display = "Obrigatória a seleção do filtro DATA."
                frmMural.Show vbModal
                
                Exit Sub
            End If
        End If
        
        fgCursor True
        Call flCarregarFlexGrid(strFiltroXML)
        fgCursor
    
        ctlTableCombo.TituloCombo = IIf(strTituloTableCombo = vbNullString, strTableComboInicial, strTituloTableCombo)
        
    End If
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmControleRemessa - objFiltro_AplicarFiltro", Me.Caption
    
End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    Select Case Button.Key
    Case "showfiltro"
        Set objFiltro = New frmFiltro
        Set objFiltro.FormOwner = Me
        objFiltro.TipoFiltro = enumTipoFiltroA6.frmControleRemessa
        objFiltro.Show vbModal
    
    Case "refresh"
        If Trim(strFiltroXML) = vbNullString Then Exit Sub
        
        fgCursor True
        Call flCarregarFlexGrid(strFiltroXML)
        fgCursor
    
    End Select
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmControleRemessa - tlbButtons_ButtonClick", Me.Caption
    
End Sub

