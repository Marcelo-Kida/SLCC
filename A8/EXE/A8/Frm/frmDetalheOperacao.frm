VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDetalheOperacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalhe Operação"
   ClientHeight    =   8730
   ClientLeft      =   2685
   ClientTop       =   1245
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   9375
   Begin TabDlg.SSTab sstDetalhe 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4471
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Mensagens Internas"
      TabPicture(0)   =   "frmDetalheOperacao.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flxMsgInterna"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Situações"
      TabPicture(1)   =   "frmDetalheOperacao.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxSituacao"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Mensagens Associadas"
      TabPicture(2)   =   "frmDetalheOperacao.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "flxMsgAssociada"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Conciliação"
      TabPicture(3)   =   "frmDetalheOperacao.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "flxConciliacao"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Corretoras"
      TabPicture(4)   =   "frmDetalheOperacao.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "flxCorretoras"
      Tab(4).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid flxMsgInterna 
         Height          =   1935
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3413
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxSituacao 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3413
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxMsgAssociada 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3413
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxConciliacao 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3413
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxCorretoras 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3413
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   120
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalheOperacao.frx":008C
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblBotoes 
      Height          =   330
      Left            =   8520
      TabIndex        =   0
      Top             =   8400
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      ButtonWidth     =   1376
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser wbDetalhe 
      Height          =   5475
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   9135
      ExtentX         =   16113
      ExtentY         =   9657
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmDetalheOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:33:55
'-------------------------------------------------
'' Objeto responsável pela consulta dos detalhes sobre uma operação ou mensagem,
'' através de interação com a camada de controle de caso de uso MIU.
''
'' Classes especificamente consideradas de destino:
''  A8MIU.clsMensagem
''  A8MIU.clsOperacao
''
Option Explicit

'para guardar a chave da operacao
Private strOwner                            As String

'para guardar a chave da operacao
Private vntSequenciaOperacao                As Variant

Private lngCodigoEmpresa                    As Long

'para guardar a chave da mensagem
Private strNumeroControleIF                 As String
Private datDataRegistroMensagem             As Date
Private lngNumeroSequenciaRepeticaoMensagem As Long

Private strConteudoBrowser                  As String

Private blnDetalheMensagem                  As Boolean

Private Const SEM_MENSAGEM                  As String = "Inexistente"

'Abas do Objeto SSTab
Private Const ABA_MENSAGENS_INTERNAS        As Integer = 0
Private Const ABA_SITUACOES                 As Integer = 1
Private Const ABA_MENSAGENS_ASSOCIADAS      As Integer = 2  '<-- Depende da tela de onde o
Private Const ABA_OPERACOES_ASSOCIADAS      As Integer = 2  '    FORM foi chamado
Private Const ABA_CONCILIACAO               As Integer = 3
Private Const ABA_CORRETORAS                As Integer = 4

'Colunas FlexGrid Mensagens Internas (MI)
Private Const COL_MI_DATA                   As Integer = 0
Private Const COL_MI_TIPO_MENSAGEM          As Integer = 1
Private Const COL_MI_TIPO_SOLICITACAO       As Integer = 2
Private Const COL_MI_CODIGO_XML_OCULTO      As Integer = 3

'Colunas FlexGrid Situações (ST)
Private Const COL_ST_DATA                   As Integer = 0
Private Const COL_ST_SITUACAO               As Integer = 1

'Situações para operações
Private Const COL_ST_JUSTIFICATIVA          As Integer = 2
Private Const COL_ST_ACAO_OP                As Integer = 3
Private Const COL_ST_TEXTOANTERIOR_OP       As Integer = 4
Private Const COL_ST_USUARIO_OP             As Integer = 5

'Situações para mensagem
Private Const COL_ST_ACAO_MS                As Integer = 2
Private Const COL_ST_TEXTOANTERIOR_MS       As Integer = 3
Private Const COL_ST_USUARIO_MS             As Integer = 4

'Colunas FlexGrid Mensagens Associadas (MA)
Private Const COL_MA_MENSAGEM               As Integer = 0
Private Const COL_MA_DATA                   As Integer = 1
Private Const COL_MA_SITUACAO               As Integer = 2
Private Const COL_MA_CODIGO_XML_OCULTO      As Integer = 3

'Colunas FlexGrid Conciliação (CN)
Private Const COL_CN_MENSAGEM               As Integer = 0
Private Const COL_CN_DATA                   As Integer = 1
Private Const COL_CN_JUSTIFICATIVA          As Integer = 2
Private Const COL_CN_TEXTO_JUSTIFICATIVA    As Integer = 3
Private Const COL_CN_CODIGO_XML_OCULTO      As Integer = 4

'Colunas FlexGrid Corretoras (CR)
Private Const COL_CR_DESCRICAO              As Integer = 0
Private Const COL_CR_INDICADOR_COMPRA       As Integer = 1
Private Const COL_CR_QUANTIDADE             As Integer = 2
Private Const COL_CR_PRECO                  As Integer = 3
Private Const COL_CR_VALOR_FINCEIRO         As Integer = 4

Private gstrTipoMovimento                   As String

Private lngCodigoErroNegocio                As Long

Private Sub flxConciliacao_Click()
    
On Error GoTo ErrorHandler
    
    fgPositionRowFlexGrid flxConciliacao.Row, flxConciliacao
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "flxConciliacao_Click", Me.Caption
    
End Sub

Private Sub flxConciliacao_DblClick()
    
On Error GoTo ErrorHandler
    
    With flxConciliacao
        If .Rows <= 1 Then Exit Sub
        If .TextMatrix(.RowSel, COL_CN_CODIGO_XML_OCULTO) = vbNullString Then Exit Sub
        
        Call fgCursor(True)
        Call flCarregaMensagemHTML(.TextMatrix(.RowSel, COL_CN_CODIGO_XML_OCULTO), lngCodigoEmpresa)
    
    End With
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "flxConciliacao_DblClick", Me.Caption
    
End Sub

Private Sub flxMsgAssociada_Click()
    
On Error GoTo ErrorHandler
    
    fgPositionRowFlexGrid flxMsgAssociada.Row, flxMsgAssociada
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "flxMsgAssociada_Click", Me.Caption
    
End Sub

'Verificação do código da mensagem SPB se deverá ou não ser exibida

Private Function flExibicaoXMLPermitido(ByVal pstrMensagem As String) As Boolean

On Error GoTo ErrorHandler

    Select Case pstrMensagem
'        Case "LTR0008", "LTR0002", "LTR0004", "STR0004", "LTR0003", _
              "LDL0004", "LDL0003", "SEL1023", "BMA0004", "LDL1003", "LDL1002"
        Case "SEL1023", "BMA0004", "LDL1003", "LDL1002"
            flExibicaoXMLPermitido = False
        Case Else
            flExibicaoXMLPermitido = True
    End Select

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flExibicaoXMLPermitido", 0
End Function

Private Sub flxMsgAssociada_DblClick()

On Error GoTo ErrorHandler
    
    With flxMsgAssociada
        If .Rows <= 1 Then Exit Sub
        If .TextMatrix(.RowSel, COL_MA_CODIGO_XML_OCULTO) = vbNullString Then Exit Sub
        If flExibicaoXMLPermitido(.TextMatrix(.RowSel, COL_MA_MENSAGEM)) Then
            Call fgCursor(True)
            Call flCarregaMensagemHTML(.TextMatrix(.RowSel, COL_MA_CODIGO_XML_OCULTO), lngCodigoEmpresa)
            
        Else
            Call flApresentacaoDetalhe(False, "Esta mensagem não pode ser exibida")
        End If
    End With
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "flxMsgAssociada_DblClick", Me.Caption

End Sub

Private Sub flxMsgInterna_Click()
    
On Error GoTo ErrorHandler
    
    fgPositionRowFlexGrid flxMsgInterna.Row, flxMsgInterna
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "flxMsgInterna_Click", Me.Caption
    
End Sub

Private Sub flxMsgInterna_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    
On Error GoTo ErrorHandler

    With flxMsgInterna
        If .TextMatrix(Row1, COL_MI_DATA) = vbNullString Then
            If .TextMatrix(Row2, COL_MI_DATA) = vbNullString Then
                Cmp = 0
            Else
                Cmp = 1
            End If
            Exit Sub
        End If
        
        If .TextMatrix(Row2, COL_MI_DATA) = vbNullString Then
            Cmp = -1
            Exit Sub
        End If
        
        If CDate(.TextMatrix(Row1, COL_MI_DATA)) > CDate(.TextMatrix(Row2, COL_MI_DATA)) Then
            Cmp = -1
        ElseIf CDate(.TextMatrix(Row1, COL_MI_DATA)) < CDate(.TextMatrix(Row2, COL_MI_DATA)) Then
            Cmp = 1
        Else
            Cmp = 0
        End If
    End With

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flxMsgInterna_Compare"
End Sub

Private Sub flxMsgInterna_DblClick()

On Error GoTo ErrorHandler
    
    With flxMsgInterna
        If .Rows <= 1 Then Exit Sub
        If .TextMatrix(.RowSel, COL_MI_CODIGO_XML_OCULTO) = vbNullString Then Exit Sub
        
        Call fgCursor(True)
        Call flCarregaMensagemHTML(.TextMatrix(.RowSel, COL_MI_CODIGO_XML_OCULTO), lngCodigoEmpresa)
    End With
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "flxMsgInterna_DblClick", Me.Caption
    
End Sub

Private Sub flxSituacao_Click()
    
On Error GoTo ErrorHandler
    
    fgPositionRowFlexGrid flxSituacao.Row, flxSituacao
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "flxSituacao_Click", Me.Caption
    
End Sub

Private Sub Form_Initialize()
 
    lngNumeroSequenciaRepeticaoMensagem = 1

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    Set Me.Icon = mdiLQS.Icon
    blnDetalheMensagem = strNumeroControleIF <> vbNullString
    
    If blnDetalheMensagem Then
        sstDetalhe.TabVisible(4) = False
        Me.Caption = "Detalhe Mensagem"
    End If
    
    'Limpa o conteúdo do Browser
    Call flAtualizaConteudoBrowser
    
    If blnDetalheMensagem Then
        sstDetalhe.TabCaption(0) = "Operação Associada"
    End If
    
    'Configura a formatação dos Grids
    Call flFormatarFlexGrid(flxMsgInterna)          'Ou Mensagem Associada
    Call flFormatarFlexGrid(flxSituacao)
    Call flFormatarFlexGrid(flxMsgAssociada)        'Ou Operação Associada
    Call flFormatarFlexGrid(flxConciliacao)
    Call flFormatarFlexGrid(flxCorretoras)
    
    'Carrega o Grid de Mensagens Internas
    Call flCarregarFlexGrid(flxMsgInterna)
    Call flApresentacaoDetalhe(True)
    
    '>>> -------------------------------------------
    'O restante dos Grids será carregado por demanda
    '>>> -------------------------------------------
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "Form_Load", Me.Caption
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmDetalheOperacao = Nothing
    gintRowPositionAnt = 0
    
End Sub

Private Sub sstDetalhe_Click(PreviousTab As Integer)

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    Select Case sstDetalhe.Tab
           Case ABA_MENSAGENS_INTERNAS
                Call flCarregarFlexGrid(flxMsgInterna)
                Call flApresentacaoDetalhe(True)
                
           Case ABA_SITUACOES
                Call flCarregarFlexGrid(flxSituacao)
                Call flApresentacaoDetalhe(False)
                
           Case ABA_MENSAGENS_ASSOCIADAS
                Call flCarregarFlexGrid(flxMsgAssociada)
                Call flApresentacaoDetalhe(True)
                
           Case ABA_CONCILIACAO
                Call flCarregarFlexGrid(flxConciliacao)
                Call flApresentacaoDetalhe(True)
                
            Case ABA_CORRETORAS
                Call flCarregarFlexGrid(flxCorretoras)
                Call flApresentacaoDetalhe(False, gstrTipoMovimento)
                
    End Select
    
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "sstDetalhe_Click", Me.Caption
    
End Sub

Private Sub tblBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    Select Case Button.Key
        Case gstrSair
            Unload Me
    
    End Select

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tblBotoes_ButtonClick"

End Sub

'' Formata o FlexGrid de acordo com o uso do objeto(detalhe de operação ou
'' mensagem).
Private Sub flFormatarFlexGrid(ByVal flxFlexGrid As MSFlexGrid)

On Error GoTo ErrorHandler

    With flxFlexGrid
        .ReDraw = False
        .Rows = 0
        .Rows = 1000
        .FixedRows = 1
        .FixedCols = 0
        
        Select Case .Name
            Case "flxMsgInterna"
                If Not blnDetalheMensagem Then
                    .Cols = 4
                    .ColWidth(COL_MI_DATA) = 2000
                    .ColWidth(COL_MI_TIPO_MENSAGEM) = 2000
                    .ColWidth(COL_MI_TIPO_SOLICITACAO) = 2000
                    .ColWidth(COL_MI_CODIGO_XML_OCULTO) = 0
                    .TextMatrix(0, COL_MI_DATA) = "Data/Hora Receb./Envio"
                    .TextMatrix(0, COL_MI_TIPO_MENSAGEM) = "Tipo Mensagem Interna"
                    .TextMatrix(0, COL_MI_TIPO_SOLICITACAO) = "Tipo Solicitação"
                Else
                    .Cols = 4
                    .ColWidth(COL_MA_MENSAGEM) = 1500
                    .ColWidth(COL_MA_DATA) = 2000
                    .ColWidth(COL_MA_SITUACAO) = 2000
                    .ColWidth(COL_MA_CODIGO_XML_OCULTO) = 0
                    .TextMatrix(0, COL_MA_MENSAGEM) = "Código Mensagem Interna"
                    .TextMatrix(0, COL_MA_DATA) = "Data/Hora Envio"
                    .TextMatrix(0, COL_MA_SITUACAO) = "Situação"
                End If

            Case "flxSituacao"
                .Cols = IIf(blnDetalheMensagem, 5, 6)
                .ColWidth(COL_ST_DATA) = 2000
                .TextMatrix(0, COL_ST_DATA) = "Data/Hora"
                .ColWidth(COL_ST_SITUACAO) = 2000
                .TextMatrix(0, COL_ST_SITUACAO) = "Situação"
                If blnDetalheMensagem Then
                    .ColWidth(COL_ST_ACAO_MS) = 2700
                    .ColAlignment(COL_ST_ACAO_MS) = flexAlignLeftCenter
                    .TextMatrix(0, COL_ST_ACAO_MS) = "Ação"
                    .ColWidth(COL_ST_TEXTOANTERIOR_MS) = 2500
                    .TextMatrix(0, COL_ST_TEXTOANTERIOR_MS) = "Texto Anterior"
                    .ColWidth(COL_ST_USUARIO_MS) = 2000
                    .TextMatrix(0, COL_ST_USUARIO_MS) = "Código Usuário"
                Else
                    .ColWidth(COL_ST_JUSTIFICATIVA) = 2500
                    .TextMatrix(0, COL_ST_JUSTIFICATIVA) = "Justificativa Situação"
                    .ColWidth(COL_ST_ACAO_OP) = 2000
                    .ColAlignment(COL_ST_ACAO_OP) = flexAlignLeftCenter
                    .TextMatrix(0, COL_ST_ACAO_OP) = "Ação"
                    .ColWidth(COL_ST_TEXTOANTERIOR_OP) = 2500
                    .TextMatrix(0, COL_ST_TEXTOANTERIOR_OP) = "Texto Anterior / Justif. Ação"
                    .ColWidth(COL_ST_USUARIO_OP) = 2000
                    .TextMatrix(0, COL_ST_USUARIO_OP) = "Código Usuário"
                End If
                
            Case "flxMsgAssociada"      'Ou Operação Associada
                .Cols = 4
                .ColWidth(COL_MA_MENSAGEM) = 1500
                .ColWidth(COL_MA_DATA) = 2000
                .ColWidth(COL_MA_SITUACAO) = 2000
                .ColWidth(COL_MA_CODIGO_XML_OCULTO) = 0
                .TextMatrix(0, COL_MA_MENSAGEM) = IIf(blnDetalheMensagem, "Código Mensagem", "Código Operação")
                .TextMatrix(0, COL_MA_DATA) = "Data/Hora Envio"
                .TextMatrix(0, COL_MA_SITUACAO) = "Situação"
                
            Case "flxConciliacao"
                .Cols = 5
                .ColWidth(COL_CN_MENSAGEM) = 1500
                .ColWidth(COL_CN_DATA) = 2000
                .ColWidth(COL_CN_JUSTIFICATIVA) = 2500
                .ColWidth(COL_CN_TEXTO_JUSTIFICATIVA) = 2500
                .ColWidth(COL_CN_CODIGO_XML_OCULTO) = 0
                If blnDetalheMensagem Then
                    .TextMatrix(0, COL_CN_MENSAGEM) = "Código Operação"
                Else
                    .TextMatrix(0, COL_CN_MENSAGEM) = "Código Mensagem"
                End If
                .TextMatrix(0, COL_CN_DATA) = "Data/Hora Conciliação"
                .TextMatrix(0, COL_CN_JUSTIFICATIVA) = "Justificativa"
                .TextMatrix(0, COL_CN_TEXTO_JUSTIFICATIVA) = "Texto Justificativa"
                
            Case "flxCorretoras"
            
                .Cols = 5
                .ColWidth(COL_CR_DESCRICAO) = 1500
                .ColWidth(COL_CR_INDICADOR_COMPRA) = 2000
                .ColWidth(COL_CR_QUANTIDADE) = 1500
                .ColWidth(COL_CR_PRECO) = 1500
                .ColWidth(COL_CR_VALOR_FINCEIRO) = 1500
                .TextMatrix(0, COL_CR_DESCRICAO) = "Descrição"
                .TextMatrix(0, COL_CR_INDICADOR_COMPRA) = "Indicador Compra/Venda"
                .TextMatrix(0, COL_CR_QUANTIDADE) = "Quantidade"
                .TextMatrix(0, COL_CR_PRECO) = "Preço Unitário"
                .TextMatrix(0, COL_CR_VALOR_FINCEIRO) = "Valor Financeiro"
            
        End Select
        
        .ReDraw = True
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flFormatarFlexGrid", 0
    
End Sub

'' Preenche o flexgrid informado mcomo parâmetro com os dados de acordo com o
'' flexgrid e com o tipo de uso do objeto(mensagem ou operação).Métodos incocados
'' na camada de controle de caso de uso:  A8MIU.clsOperacao.ObterMensagensInternas
'' A8MIU.clsOperacao.ObterSituacoesOperacao  A8MIU.clsMensagem.
'' ObterHistoricoMensagem  A8MIU.clsOperacao.ObterMensagensAssociadas  A8MIU.
'' clsMensagem.ObterMensagensAssociadas  A8MIU.clsOperacao.
'' ObterConciliacaoMensagem  A8MIU.clsMensagem.ObterConciliacaoOperacao
Private Sub flCarregarFlexGrid(ByVal flxFlexGrid As MSFlexGrid)

#If EnableSoap = 1 Then
    Dim objOperacao         As MSSOAPLib30.SoapClient30
    Dim objMensagem         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao         As A8MIU.clsOperacao
    Dim objMensagem         As A8MIU.clsMensagem
#End If

Dim objDomNode              As IXMLDOMNode
Dim xmlDomFiltros           As MSXML2.DOMDocument40
Dim xmlDomLeitura           As MSXML2.DOMDocument40
Dim strRetLeitura           As String
Dim lngLinhaGrid            As Long

Dim strTipoAcao             As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Call fgCursor(True)
    Call flFormatarFlexGrid(flxFlexGrid)
    gstrTipoMovimento = vbNullString
    
    gintRowPositionAnt = 0
    
    With flxFlexGrid
        .ReDraw = False
        
        '>>> Formata XML Filtro padrão... --------------------------------------------------------------
        Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
        
        Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")

        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_SequenciaOperacao", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_SequenciaOperacao", _
                                         "SequenciaOperacao", vntSequenciaOperacao)
                                         
        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_NumeroCtrlIF", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_NumeroCtrlIF", "NumeroCtrlIF", strNumeroControleIF)
        
        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_DataRegistro", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_DataRegistro", "DataRegistro", fgDtHr_To_Xml(datDataRegistroMensagem))
        
        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_NumeroSequenciaControleRepeticao", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_NumeroSequenciaControleRepeticao", "NumeroSequenciaControleRepeticao", lngNumeroSequenciaRepeticaoMensagem)
        '>>> -------------------------------------------------------------------------------------------
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
        Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
        
        'Verifica qual Grid será preenchido
        Select Case .Name               'Mensagem Interna
            Case "flxMsgInterna"
                If vntSequenciaOperacao <> 0 Then
                    strRetLeitura = objOperacao.ObterMensagensInternas(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
                    
                    If vntCodErro <> 0 Then
                        GoTo ErrorHandler
                    End If
                Else
                    strRetLeitura = vbNullString
                End If

                If strRetLeitura <> vbNullString Then
                    If Not xmlDomLeitura.loadXML(strRetLeitura) Then
                        Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarFlexGrid")
                    End If

                    For Each objDomNode In xmlDomLeitura.documentElement.childNodes
                        lngLinhaGrid = lngLinhaGrid + 1
                        
                        .TextMatrix(lngLinhaGrid, COL_MI_DATA) = fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("DH_MESG_INTE").Text)
                        .TextMatrix(lngLinhaGrid, COL_MI_TIPO_MENSAGEM) = objDomNode.selectSingleNode("NO_TIPO_MESG").Text
                        
                        Select Case Val(objDomNode.selectSingleNode("TP_SOLI_MESG_INTE").Text)
                            Case enumTipoSolicitacao.Inclusao
                                strTipoAcao = "Inclusão"
                            Case enumTipoSolicitacao.Complementacao
                                strTipoAcao = "Complementação"
                            Case enumTipoSolicitacao.Cancelamento
                                strTipoAcao = "Cancelamento"
                            Case enumTipoSolicitacao.RetornoLegado
                                strTipoAcao = "Retorno Legado"
                            Case enumTipoSolicitacao.Alteracao
                                strTipoAcao = "Alteração"
                            Case enumTipoSolicitacao.CancelamentoComMensagem
                                strTipoAcao = "Cancelamento com Mensagem"
                            Case enumTipoSolicitacao.LivreMovimentacao
                                strTipoAcao = "Livre Movimentação"
                            Case enumTipoSolicitacao.Confirmacao
                                strTipoAcao = "Confirmação"
                            Case Else
                                strTipoAcao = ""
                        End Select
                        
                        .TextMatrix(lngLinhaGrid, COL_MI_TIPO_SOLICITACAO) = strTipoAcao
                        .TextMatrix(lngLinhaGrid, COL_MI_CODIGO_XML_OCULTO) = objDomNode.selectSingleNode("CO_TEXT_XML").Text
                    Next
                    
                    .Sort = 9
                    
                End If
                
            Case "flxSituacao"          'Situações
            
                If Not blnDetalheMensagem Then
                    'Situações da operação
                    strRetLeitura = objOperacao.ObterSituacoesOperacao(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
                    
                    If vntCodErro <> 0 Then
                        GoTo ErrorHandler
                    End If
                Else
                    'Situações da mensagem
                    strRetLeitura = objMensagem.ObterHistoricoMensagem(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
                    
                    If vntCodErro <> 0 Then
                        GoTo ErrorHandler
                    End If
                End If
    
                If strRetLeitura <> vbNullString Then
                    If Not xmlDomLeitura.loadXML(strRetLeitura) Then
                        Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarFlexGrid")
                    End If

                    For Each objDomNode In xmlDomLeitura.documentElement.childNodes
                        lngLinhaGrid = lngLinhaGrid + 1
                        
                        .TextMatrix(lngLinhaGrid, COL_ST_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                        If blnDetalheMensagem Then
                            .TextMatrix(lngLinhaGrid, COL_ST_DATA) = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_SITU_ACAO_MESG_SPB").Text)
                            If Not Trim$(objDomNode.selectSingleNode("TP_ACAO_MESG_SPB").Text) = vbNullString And fgVlrXml_To_Decimal(objDomNode.selectSingleNode("TP_ACAO_MESG_SPB").Text) <> 0 Then
                                .TextMatrix(lngLinhaGrid, COL_ST_ACAO_MS) = objDomNode.selectSingleNode("TP_ACAO_MESG_SPB").Text & " - " & fgDescricaoTipoAcao(CLng("0" & objDomNode.selectSingleNode("TP_ACAO_MESG_SPB").Text))
                            End If
                            .TextMatrix(lngLinhaGrid, COL_ST_TEXTOANTERIOR_MS) = objDomNode.selectSingleNode("TX_CNTD_ANTE_ACAO").Text
                            .TextMatrix(lngLinhaGrid, COL_ST_USUARIO_MS) = objDomNode.selectSingleNode("CO_USUA_ULTI_ATLZ").Text
                        Else
                            .TextMatrix(lngLinhaGrid, COL_ST_DATA) = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_SITU_ACAO_OPER_ATIV").Text)
                            .TextMatrix(lngLinhaGrid, COL_ST_JUSTIFICATIVA) = objDomNode.selectSingleNode("NO_TIPO_JUST_SITU_PROC").Text
                            If Trim$(objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV").Text) <> vbNullString And fgVlrXml_To_Decimal(objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV").Text) <> 0 Then
                                .TextMatrix(lngLinhaGrid, COL_ST_ACAO_OP) = objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV").Text & " - " & fgDescricaoTipoAcao(CLng("0" & objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV").Text))
                            End If
                            .TextMatrix(lngLinhaGrid, COL_ST_TEXTOANTERIOR_OP) = objDomNode.selectSingleNode("TX_CNTD_ANTE_ACAO").Text
                            .TextMatrix(lngLinhaGrid, COL_ST_USUARIO_OP) = objDomNode.selectSingleNode("CO_USUA_ATLZ").Text
                        End If
                    Next
                End If
                
            Case "flxMsgAssociada"      'Mensagens Associadas OU Operações Associadas
                                
                If Not blnDetalheMensagem Then
                    strRetLeitura = objOperacao.ObterMensagensAssociadas(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
                    
                    If vntCodErro <> 0 Then
                        GoTo ErrorHandler
                    End If
    
                    If strRetLeitura <> vbNullString Then
                        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
                            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarFlexGrid")
                        End If
    
                        For Each objDomNode In xmlDomLeitura.documentElement.childNodes
                            lngLinhaGrid = lngLinhaGrid + 1
                            
                            .TextMatrix(lngLinhaGrid, COL_MA_MENSAGEM) = objDomNode.selectSingleNode("CO_MESG_SPB").Text
                            .TextMatrix(lngLinhaGrid, COL_MA_DATA) = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text)
                            .TextMatrix(lngLinhaGrid, COL_MA_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                            .TextMatrix(lngLinhaGrid, COL_MA_CODIGO_XML_OCULTO) = objDomNode.selectSingleNode("CO_TEXT_XML").Text
                        Next
                    End If
                Else
                    strRetLeitura = objMensagem.ObterMensagensAsssociadas(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
                    
                    If vntCodErro <> 0 Then
                        GoTo ErrorHandler
                    End If
                                                            
                    If strRetLeitura <> vbNullString Then
                        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
                            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarFlexGrid")
                        End If
    
                        For Each objDomNode In xmlDomLeitura.documentElement.childNodes
                            lngLinhaGrid = lngLinhaGrid + 1
                            
                            .TextMatrix(lngLinhaGrid, COL_MA_MENSAGEM) = objDomNode.selectSingleNode("CO_MESG_SPB").Text
                            .TextMatrix(lngLinhaGrid, COL_MA_DATA) = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text)
                            .TextMatrix(lngLinhaGrid, COL_MA_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                            .TextMatrix(lngLinhaGrid, COL_MA_CODIGO_XML_OCULTO) = objDomNode.selectSingleNode("CO_TEXT_XML").Text
                        Next
                    End If
                End If
               
            Case "flxConciliacao"       'Conciliação
                If blnDetalheMensagem Then
                    strRetLeitura = objMensagem.ObterConciliacaoMensagem(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
                    
                    If vntCodErro <> 0 Then
                        GoTo ErrorHandler
                    End If
                    
                Else
                    strRetLeitura = objOperacao.ObterConciliacaoOperacao(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
                    
                    If vntCodErro <> 0 Then
                        GoTo ErrorHandler
                    End If
                End If

                If strRetLeitura <> vbNullString Then
                    If Not xmlDomLeitura.loadXML(strRetLeitura) Then
                        Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarFlexGrid")
                    End If
                    
                    For Each objDomNode In xmlDomLeitura.documentElement.childNodes
                        lngLinhaGrid = lngLinhaGrid + 1
                        
                        .TextMatrix(lngLinhaGrid, COL_CN_MENSAGEM) = objDomNode.selectSingleNode("CO_MESG_SPB").Text
                        If objDomNode.selectSingleNode("CO_MESG_SPB").Text = SEM_MENSAGEM Then
                            .Row = lngLinhaGrid
                            .Col = COL_CN_MENSAGEM
                            .CellForeColor = vbRed
                        End If
                        
                        .TextMatrix(lngLinhaGrid, COL_CN_DATA) = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_JUST_CNCL").Text)
                        .TextMatrix(lngLinhaGrid, COL_CN_JUSTIFICATIVA) = objDomNode.selectSingleNode("NO_TIPO_JUST_CNCL").Text
                        .TextMatrix(lngLinhaGrid, COL_CN_TEXTO_JUSTIFICATIVA) = objDomNode.selectSingleNode("TX_JUST").Text
                        .TextMatrix(lngLinhaGrid, COL_CN_CODIGO_XML_OCULTO) = objDomNode.selectSingleNode("CO_TEXT_XML").Text
                    Next
                End If
                
            Case "flxCorretoras"       'Corretoras
                strRetLeitura = objOperacao.ObterDetalheCorretoras(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)

                If vntCodErro <> 0 Then
                    GoTo ErrorHandler
                End If

                If strRetLeitura <> vbNullString Then
                    If Not xmlDomLeitura.loadXML(strRetLeitura) Then
                        Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarFlexGrid")
                    End If
                    
                    For Each objDomNode In xmlDomLeitura.documentElement.childNodes
                    
                        If objDomNode.baseName = "DE_MOVI" Then
                             gstrTipoMovimento = objDomNode.Text
                        Else
                            lngLinhaGrid = lngLinhaGrid + 1
                            
                            .TextMatrix(lngLinhaGrid, COL_CR_DESCRICAO) = objDomNode.selectSingleNode("DE_ATIV_MERC").Text
                            .TextMatrix(lngLinhaGrid, COL_CR_INDICADOR_COMPRA) = objDomNode.selectSingleNode("IN_CPRA_VEND").Text
                            .TextMatrix(lngLinhaGrid, COL_CR_QUANTIDADE) = objDomNode.selectSingleNode("QT_ATIV_MERC").Text
                            .TextMatrix(lngLinhaGrid, COL_CR_PRECO) = objDomNode.selectSingleNode("PU_ATIV_MERC").Text
                            .TextMatrix(lngLinhaGrid, COL_CR_VALOR_FINCEIRO) = objDomNode.selectSingleNode("VA_OPER_ATIV_LIQU").Text
                        End If
                    Next
                Else
                    gstrTipoMovimento = "Esta não é uma operação com corretoras"
                End If
                
        End Select
        
        .Rows = lngLinhaGrid + 1
        
        If .Rows > 1 Then
            .Col = 0
            .Row = 1
            Select Case flxFlexGrid
                   Case flxConciliacao
                        flxConciliacao_Click
                        
                   Case flxMsgAssociada
                        flxMsgAssociada_Click
                        
                   Case flxMsgInterna
                        flxMsgInterna_Click
                        
                   Case flxSituacao
                        flxSituacao_Click
                        
            End Select
        End If
        
        .ReDraw = True
        
    End With
    
    Set xmlDomFiltros = Nothing
    Set xmlDomLeitura = Nothing
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    
    Call fgCursor(False)

Exit Sub
ErrorHandler:

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarFlexGrid", 0
    
End Sub
Public Property Let BaseOwner(ByVal strBaseOwner As String)
    strOwner = UCase$(Trim$(strBaseOwner))
End Property

Public Property Let SequenciaOperacao(ByVal vntSeqOperacao As Variant)
    vntSequenciaOperacao = vntSeqOperacao
End Property

Public Property Let NumeroControleIF(ByVal strNewValue As String)
    strNumeroControleIF = strNewValue
End Property
Public Property Let NumeroSequenciaRepeticao(ByVal strNewValue As Long)
    lngNumeroSequenciaRepeticaoMensagem = strNewValue
End Property
Public Property Let DataRegistroMensagem(ByVal datNewValue As Date)
    datDataRegistroMensagem = datNewValue
End Property

Public Property Let CodigoEmpresa(ByVal plngCodigoEmpresa As Long)
    lngCodigoEmpresa = plngCodigoEmpresa
End Property

'Carregar o conteúdo do Browser com a mensagem xml
Private Sub flCarregaMensagemHTML(ByVal plngSequencial As Long, _
                                  ByVal plngCodigoEmpresa As Long)

#If EnableSoap = 1 Then
    Dim objMensagem         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem         As A8MIU.clsMensagem
#End If

Dim xmlDomFiltros           As MSXML2.DOMDocument40
Dim strMensagemHTML         As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Call fgCursor(True)
    Call flAtualizaConteudoBrowser

    If vntSequenciaOperacao < 0 Or strOwner = "A8HIST" Then
        plngSequencial = plngSequencial * -1
    End If

    '>>> Formata XML Filtro padrão... --------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Sequencial", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Sequencial", _
                                     "Sequencial", plngSequencial)
                                     
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Empresa", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Empresa", _
                                     "Empresa", plngCodigoEmpresa)
                                     
    '>>> -------------------------------------------------------------------------------------------
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    strMensagemHTML = objMensagem.ObterMensagemHTML(xmlDomFiltros.xml, _
                                                    vntCodErro, _
                                                    vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Call flFormataDatas(strMensagemHTML)
    
    'Força o Browser a atualizar a página com o conteúdo obtido
    Call flAtualizaConteudoBrowser(strMensagemHTML)
     
    Set objMensagem = Nothing
    Set xmlDomFiltros = Nothing
    
    Call fgCursor(False)

Exit Sub
ErrorHandler:
    fgCursor
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaMensagemHTML", 0

End Sub

'Formatar as datas para exibição
Private Sub flFormataDatas(ByRef strMensagemHTML As String)

Dim lngPosicao                              As Long
Dim strDataRawFormat                        As String
Dim strSaida                                As String
Dim datDataFormatada                        As Date

Const TAG_DATA_HORA                         As String = "|DH|"
Const TAG_DATA_HORA_SIZE                    As Long = 19

Const TAG_DATA                              As String = "|DT|"
Const TAG_DATA_SIZE                         As Long = 13

    Do
        lngPosicao = InStr(strMensagemHTML, TAG_DATA)
        If lngPosicao > 0 Then
            strSaida = Mid$(strMensagemHTML, 1, lngPosicao - 1)
            strDataRawFormat = Mid$(strMensagemHTML, lngPosicao, TAG_DATA_SIZE)
            datDataFormatada = fgDtXML_To_Date(Mid$(strDataRawFormat, Len(TAG_DATA) + 1, 8))
            strSaida = strSaida & datDataFormatada & Mid$(strMensagemHTML, lngPosicao + TAG_DATA_SIZE)
            strMensagemHTML = strSaida
        End If
    Loop While lngPosicao <> 0

    Do
        lngPosicao = InStr(strMensagemHTML, TAG_DATA_HORA)
        If lngPosicao > 0 Then
            strSaida = Mid$(strMensagemHTML, 1, lngPosicao - 1)
            strDataRawFormat = Mid$(strMensagemHTML, lngPosicao, TAG_DATA_HORA_SIZE)
            datDataFormatada = fgDtHrStr_To_DateTime(Mid$(strDataRawFormat, Len(TAG_DATA_HORA) + 1, 14))
            strSaida = strSaida & datDataFormatada & Mid$(strMensagemHTML, lngPosicao + TAG_DATA_HORA_SIZE)
            strMensagemHTML = strSaida
        End If
    Loop While lngPosicao <> 0

End Sub

Private Sub wbDetalhe_DocumentComplete(ByVal pDisp As Object, URL As Variant)

On Error GoTo ErrorHandler

    pDisp.Document.Body.innerHTML = strConteudoBrowser

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - wbDetalhe_DocumentComplete"

End Sub

'' Atualiza o html exibido sobre a mensagem
Private Sub flAtualizaConteudoBrowser(Optional pstrConteudo As String = vbNullString)

On Error GoTo ErrorHandler

    strConteudoBrowser = pstrConteudo
    wbDetalhe.Navigate "about:blank"

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flAtualizaConteudoBrowser", 0

End Sub

'' Cria o cabeçãlho do html do detalhe
Private Sub flApresentacaoDetalhe(ByVal pblnApresenta As Boolean, _
                         Optional ByVal pstrMensagem As String = vbNullString)

Dim strMensagemHTML                         As String

On Error GoTo ErrorHandler

    If pstrMensagem = vbNullString Then
        pstrMensagem = "Esta seção não apresenta detalhe"
    End If

    If pblnApresenta Then
        strMensagemHTML = "<Html><Body><Center><Table border = 0>" & _
                          "<TR><TD BGColor=""#BBBBBB""><Font Color=White Size=3 Face=Verdana>" & _
                          "Selecione um item na lista para a exibição do detalhe" & _
                          "</Font></TD></TR>" & _
                          "</Table></Center></Body></Html>"
    Else
        strMensagemHTML = "<Html><Body><Center><Table border = 0>" & _
                          "<TR><TD BGColor=""#BBBBBB""><Font Color=White Size=3 Face=Verdana>" & _
                          pstrMensagem & _
                          "</Font></TD></TR>" & _
                          "</Table></Center></Body></Html>"
    End If
    
    Call flAtualizaConteudoBrowser(strMensagemHTML)

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flApresentacaoDetalhe", 0

End Sub
