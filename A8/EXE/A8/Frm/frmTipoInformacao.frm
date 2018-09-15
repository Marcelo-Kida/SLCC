VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmTipoInformacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Divergência A8 X MBS"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   14055
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid flxOcorrencias 
      Height          =   8295
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   14631
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      SelectionMode   =   1
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   5040
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
            Picture         =   "frmTipoInformacao.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoInformacao.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoInformacao.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoInformacao.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoInformacao.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoInformacao.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoInformacao.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   13140
      TabIndex        =   0
      Top             =   8460
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      ButtonWidth     =   1376
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTipoInformacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:35:17
'-------------------------------------------------
'' Objeto genérico responsável pela exibição de ocorrências por tipo de informação
'' do Cadastro de Conversão MBS.
''
'' É acionado a partir do botão 'Divergência A8 X MBS' do Cadastro de Conversão
'' MBS.
''
Option Explicit

Public pstrXMLOcorrencias                   As String

Private Const COL_TIPO_INFORMADO            As Integer = 0
Private Const COL_CODIGO_INFORMADO          As Integer = 1
Private Const COL_NOME_INFORMADO            As Integer = 2
Private Const COL_DESCRICAO_OCORRENCIA      As Integer = 3

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

'Carregar o conteúdo do grid com os dados da consulta
Private Sub flCarregarFlexGrid(ByVal pxmlOcorrencias As String)

Dim xmlOcorrencias                          As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim xmlDomNodeChild                         As MSXML2.IXMLDOMNode
Dim lngLinhaGrid                            As Long
Dim strQuebraGrupo                          As String

On Error GoTo ErrorHandler

    Set xmlOcorrencias = CreateObject("MSXML2.DOMDocument.4.0")
    
    gintRowPositionAnt = 0

    Call flFormatarFlxMonitoracao

    'caso a tabela esteja sem registros não tem como carregar um XML,
    'sendo assim vai para o fim da rotina.
    If Trim(pxmlOcorrencias) <> "" Then
       If Not xmlOcorrencias.loadXML(pxmlOcorrencias) Then
          '100 - Documento XML Inválido.
          lngCodigoErroNegocio = 100
          GoTo ErrorHandler
       End If
    Else
       Call fgCursor(False)
       Exit Sub
    End If

    lngLinhaGrid = 1
    strQuebraGrupo = "XXX"

    With Me.flxOcorrencias

        .ReDraw = False

        For Each xmlDomNode In xmlOcorrencias.documentElement.selectNodes("Grupo_TipoInformacao")

            'Finalizar a Mesclagem de Colunas e Linhas
            .MergeRow(lngLinhaGrid) = False

            .MergeCol(0) = False
            .MergeCol(1) = False
            .MergeCol(2) = False
            .MergeCol(3) = False

            If strQuebraGrupo <> xmlDomNode.selectSingleNode("TP_INFO").Text Then
            
                'Formata Células do Grid
                .Row = lngLinhaGrid
                .Col = COL_TIPO_INFORMADO
                .CellFontBold = True
                .CellAlignment = flexAlignLeftBottom
                
                .Row = lngLinhaGrid
                .Col = COL_CODIGO_INFORMADO
                .CellAlignment = flexAlignRightBottom
                .Col = COL_NOME_INFORMADO
                .CellAlignment = flexAlignLeftBottom
    
                .Row = lngLinhaGrid
                .Col = COL_DESCRICAO_OCORRENCIA
                .CellAlignment = flexAlignLeftBottom

                .TextMatrix(lngLinhaGrid, COL_TIPO_INFORMADO) = xmlDomNode.selectSingleNode("DE_INFO").Text
                
                If (.Rows - 1) > lngLinhaGrid Then
                   lngLinhaGrid = lngLinhaGrid + 1
                Else
                   lngLinhaGrid = lngLinhaGrid + 1
                   .Rows = .Rows + 1
                End If
                
                For Each xmlDomNodeChild In xmlDomNode.selectNodes("Repeat_Ocorrencias/Grupo_Ocorrencias")
                
                    'Formata Células do Grid
                    .Row = lngLinhaGrid
                    .Col = COL_TIPO_INFORMADO
                    .CellAlignment = flexAlignLeftBottom
        
                    .Row = lngLinhaGrid
                    .Col = COL_CODIGO_INFORMADO
                    .CellAlignment = flexAlignRightBottom
                    .Col = COL_NOME_INFORMADO
                    .CellAlignment = flexAlignLeftBottom
        
                    .Row = lngLinhaGrid
                    .Col = COL_DESCRICAO_OCORRENCIA
                    .CellAlignment = flexAlignLeftBottom
                
                    Select Case xmlDomNode.selectSingleNode("TP_INFO").Text
                            
                           Case enumTipoInformacao.GrupoVeiculoLegal
                           
                                .TextMatrix(lngLinhaGrid, COL_CODIGO_INFORMADO) = xmlDomNodeChild.selectSingleNode("CO_GRUP_VEIC_LEGA").Text
                                .TextMatrix(lngLinhaGrid, COL_NOME_INFORMADO) = xmlDomNodeChild.selectSingleNode("NO_GRUP_VEIC_LEGA").Text
                           
                           Case enumTipoInformacao.GrupoUsuario
                           
                                .TextMatrix(lngLinhaGrid, COL_CODIGO_INFORMADO) = xmlDomNodeChild.selectSingleNode("CO_GRUP_USUA").Text
                                .TextMatrix(lngLinhaGrid, COL_NOME_INFORMADO) = xmlDomNodeChild.selectSingleNode("NO_GRUP_USUA").Text
                           
                           Case enumTipoInformacao.TipoBackOffice
                           
                                .TextMatrix(lngLinhaGrid, COL_CODIGO_INFORMADO) = xmlDomNodeChild.selectSingleNode("TP_BKOF").Text
                                .TextMatrix(lngLinhaGrid, COL_NOME_INFORMADO) = xmlDomNodeChild.selectSingleNode("DE_BKOF").Text
                           
                           Case enumTipoInformacao.LocalLiquidacao
                    
                                .TextMatrix(lngLinhaGrid, COL_CODIGO_INFORMADO) = xmlDomNodeChild.selectSingleNode("CO_LOCA_LIQU").Text
                                .TextMatrix(lngLinhaGrid, COL_NOME_INFORMADO) = xmlDomNodeChild.selectSingleNode("DE_LOCA_LIQU").Text
                    
                    End Select
                
                    .TextMatrix(lngLinhaGrid, COL_DESCRICAO_OCORRENCIA) = xmlDomNodeChild.selectSingleNode("DE_OCOR").Text
                    
                    If (.Rows - 1) > lngLinhaGrid Then
                       lngLinhaGrid = lngLinhaGrid + 1
                    Else
                       lngLinhaGrid = lngLinhaGrid + 1
                       .Rows = .Rows + 1
                    End If
                
                Next

                strQuebraGrupo = xmlDomNode.selectSingleNode("TP_INFO").Text
                
            End If

        Next

        .Rows = .Rows - 1

        .ReDraw = True

    End With
    
    If flxOcorrencias.Rows > 1 Then
       flxOcorrencias.Col = COL_TIPO_INFORMADO
       flxOcorrencias.Row = 2
       flxOcorrencias_Click
    End If
    
    Set xmlOcorrencias = Nothing

    Exit Sub

ErrorHandler:

    Call fgCursor(False)
    Set xmlOcorrencias = Nothing
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarFlexGrid", 0

End Sub

'Carregar o grid de monitoração
Private Sub flFormatarFlxMonitoracao()

Dim intCount                                As Integer
Dim intLinhaGridFix                         As Integer

On Error GoTo ErrorHandler

    intLinhaGridFix = 0
    
    With Me.flxOcorrencias
    
        .Clear
        .Rows = 2
        .Cols = 5
        .FixedRows = 1
        .FixedCols = 1
        
        For intCount = 0 To .Cols - 1
           .ColAlignment(intCount) = MSFlexGridLib.flexAlignCenterCenter
        Next
        
        .GridColorFixed = &HE6E6E6
        
        .ColWidth(COL_TIPO_INFORMADO) = 2300
        .ColWidth(COL_CODIGO_INFORMADO) = 1000
        .ColWidth(COL_NOME_INFORMADO) = 4000
        .ColWidth(COL_DESCRICAO_OCORRENCIA) = 6300
        
        .Col = 0
        .Row = intLinhaGridFix
        .CellFontBold = True

        .Col = 1
        .Row = intLinhaGridFix
        .CellFontBold = True

        .Col = 2
        .Row = intLinhaGridFix
        .CellFontBold = True

        .Col = 3
        .Row = intLinhaGridFix
        .CellFontBold = True

        .TextMatrix(intLinhaGridFix, COL_TIPO_INFORMADO) = "Tipo Informação"
        .TextMatrix(intLinhaGridFix, COL_CODIGO_INFORMADO) = "Código A8"
        .TextMatrix(intLinhaGridFix, COL_NOME_INFORMADO) = "Descrição A8"
        .TextMatrix(intLinhaGridFix, COL_DESCRICAO_OCORRENCIA) = "Ocorrências"
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        
        .GridLinesFixed = flexGridInset
        
    End With

Exit Sub

ErrorHandler:

    fgRaiseError App.EXEName, TypeName(Me), "flFormatarFlxMonitoracao", 0
   
End Sub

Private Sub flxOcorrencias_Click()

On Error GoTo ErrorHandler

    fgPositionRowFlexGrid flxOcorrencias.Row, flxOcorrencias

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flxOcorrencias_Click"
    
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    Call fgCenterMe(Me)
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents

    fgCursor True
    Call flCarregarFlexGrid(pstrXMLOcorrencias)
    fgCursor False
    
Exit Sub

ErrorHandler:

    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, "frmTipoInformacao - Form_Load", Me.Caption
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTipoInformacao = Nothing
    gintRowPositionAnt = 0
End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)
    Unload Me
End Sub
