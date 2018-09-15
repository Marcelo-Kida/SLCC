VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSubReservaResumo 
   Caption         =   "Sub-reserva - Resumo"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   9540
   Tag             =   "Resumo Movimentação"
   WindowState     =   2  'Maximized
   Begin A6.TableCombo ctlTableCombo 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   661
   End
   Begin MSComctlLib.ImageList imgOutrosIcones 
      Left            =   5670
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaResumo.frx":0000
            Key             =   "itemgrupo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaResumo.frx":0352
            Key             =   "selectednode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaResumo.frx":06A4
            Key             =   "itemelementar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaResumo.frx":09F6
            Key             =   "treeminus"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaResumo.frx":0D18
            Key             =   "treeplus"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaResumo.frx":103A
            Key             =   "leaf"
         EndProperty
      EndProperty
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
            Picture         =   "frmSubReservaResumo.frx":138C
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaResumo.frx":149E
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaResumo.frx":15B0
            Key             =   "showtreeview2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaResumo.frx":1902
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaResumo.frx":1C54
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaResumo.frx":1FA6
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaResumo.frx":22F8
            Key             =   "anterior"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubReservaResumo.frx":274A
            Key             =   "posterior"
         EndProperty
      EndProperty
   End
   Begin FPSpread.vaSpread vasLista 
      Height          =   3405
      Left            =   60
      TabIndex        =   2
      Tag             =   "Sub Reserva Resumo"
      Top             =   480
      Width           =   9435
      _Version        =   196608
      _ExtentX        =   16642
      _ExtentY        =   6006
      _StockProps     =   64
      ColHeaderDisplay=   0
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridSolid       =   0   'False
      MaxCols         =   2
      MaxRows         =   1
      OperationMode   =   1
      RowHeaderDisplay=   0
      SpreadDesigner  =   "frmSubReservaResumo.frx":2B9C
      UnitType        =   2
      UserResize      =   1
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   3945
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
            ImageKey        =   "refresh"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "    "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   6705
      TabIndex        =   1
      Top             =   30
      Width           =   300
   End
   Begin VB.Label lblBarra 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   3750
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmSubReservaResumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário consulta do resumo de movimentação do
' caixa sub-reserva.

Option Explicit

Private strFiltroXML                        As String

Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1
Private Const strTableComboInicial          As String = "Grupos de Veículos Legais"
Private lngLinhaFinalGrid                   As Long
Private Const MAX_FIXED_ROWS                As Integer = 4

' Formatação inicial do grid de consulta ao resumo de movimentação.

Private Sub flFormartarSpread()

Dim lngColunas                              As Long

On Error GoTo ErrorHandler

    With Me.vasLista
        .Redraw = False

        .MaxRows = 0
        .MaxRows = MAX_FIXED_ROWS
        .RowsFrozen = 3
        
        .CursorStyle = CursorStyleArrow
        
        .MaxCols = 17
        .ColWidth(1) = 200
        .ColWidth(2) = 200
        .ColWidth(3) = 3500
        .ColWidth(4) = 100
        .ColWidth(5) = 15
        .ColsFrozen = 4
        
        .BlockMode = True
        .Col = 1
        .Row = 3
        .Col2 = 4
        .Row2 = .MaxRows
        .AllowCellOverflow = True
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = 4
        .Row2 = .MaxRows
        .BackColorStyle = BackColorStyleOverVertGridOnly
        .BackColor = &H8000000E
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = 1
        .BackColor = vbBlack
        .ForeColor = vbWhite
        .RowHeight(1) = 300
        .FontSize = 10
        .FontBold = True
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Row = 2
        .Col2 = .MaxCols
        .Row2 = 2
        .BackColor = vbBlack
        .ForeColor = vbWhite
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Row = 3
        .Col2 = .MaxCols
        .Row2 = 3
        .BackColor = &H8000000F
        .ForeColor = vbAutomatic
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
        .BlockMode = True
        .Col = 3
        .Row = 4
        .Col2 = .MaxCols
        .Row2 = .MaxRows
        .TypeHAlign = TypeHAlignRight
        .BlockMode = False
        
        For lngColunas = 6 To .MaxCols
            .ColWidth(lngColunas) = 2260
        Next
        
        'Grupo Abertura
        .SetText 7, 1, "Abertura"
        
        .SetText 6, 3, "Data"
        .SetText 7, 3, "Status Caixa"
        .SetText 8, 3, "Valor"
        
        'Grupo Movimentação
        .SetText 11, 1, "Movimentação"
        
        .SetText 11, 2, "Realizado"
        .SetText 13, 2, "Variação"
        
        .SetText 9, 3, "Previsto"
        .SetText 10, 3, "Total"
        .SetText 11, 3, "Solicitado"
        .SetText 12, 3, "Confirmado"
        .SetText 13, 3, "Real.-Prev."
        
        'Grupo Fechamento
        .SetText 15, 1, "Fechamento"
        
        .SetText 15, 2, "Realizado"
        .SetText 17, 2, "Variação"
        
        .SetText 14, 3, "Total"
        .SetText 15, 3, "Solicitado"
        .SetText 16, 3, "Confirmado"
        .SetText 17, 3, "Real.-Prev."
        
        .BlockMode = True
        .Col = 6
        .Row = 4
        .Col2 = 8
        .Row2 = .MaxRows
        .BackColor = &HC0FFFF
        .BlockMode = False
        
        .BlockMode = True
        .Col = 9
        .Row = 4
        .Col2 = 13
        .Row2 = .MaxRows
        .BackColor = &HC0FFC0
        .BlockMode = False
        
        .BlockMode = True
        .Col = 14
        .Row = 4
        .Col2 = 17
        .Row2 = .MaxRows
        .BackColor = &HFFFFC0
        .BlockMode = False
        
        .Redraw = True
        
    End With

    Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flFormartarSpread", 0

End Sub

' Carrega resumo de movimentação.

Private Sub flCarregarSpread(ByVal xmlDocFiltros As String)

#If EnableSoap = 1 Then
    Dim objMonitoracaoFluxoCaixaReserva     As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracaoFluxoCaixaReserva     As A6MIU.clsMonitoracaoSubReserva
#End If

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim strResumoCaixaSubReserva                As String
Dim lngLinhaGrid                            As Long
Dim strQuebraGrupo                          As String
Dim lenumEstadoCaixa                        As enumEstadoCaixa
Dim strValorAbertura                        As String
Dim strDataAbertura                         As String
Dim strSituacaoCaixa                        As String

Dim intCodGrupoVeicLegal                    As Integer
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Set objMonitoracaoFluxoCaixaReserva = fgCriarObjetoMIU("A6MIU.clsMonitoracaoSubReserva")
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")

    strResumoCaixaSubReserva = objMonitoracaoFluxoCaixaReserva.ObterResumoCaixaSubReserva(xmlDocFiltros, _
                                                                                          vntCodErro, _
                                                                                          vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strResumoCaixaSubReserva = vbNullString Then
       flFormartarSpread
       Exit Sub
    End If

    flFormartarSpread
    
    lngLinhaGrid = 4
    
    With Me.vasLista
    
         .Redraw = False
        
        If Not xmlLerTodos.loadXML(strResumoCaixaSubReserva) Then
            '100 - Documento XML Inválido.
            lngCodigoErroNegocio = 100
            GoTo ErrorHandler
        End If

        lngLinhaGrid = 4
        
        For Each objDomNode In xmlLerTodos.documentElement.childNodes
            
            strDataAbertura = objDomNode.selectSingleNode("DT_CAIX_DISP").Text
            strValorAbertura = objDomNode.selectSingleNode("VA_UTLZ_ABER_CAIX").Text
            strSituacaoCaixa = objDomNode.selectSingleNode("CO_SITU_CAIX").Text
            
            If intCodGrupoVeicLegal <> Val(objDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text) Then
                
                .BlockMode = False
                .Col = 1
                .Row = lngLinhaGrid
                .CellType = CellTypePicture
                .TypePictCenter = True
                .TypeHAlign = TypeHAlignCenter
                .TypePictPicture = imgOutrosIcones.ListImages("treeminus").Picture
                
                .SetText 2, lngLinhaGrid, objDomNode.selectSingleNode("NO_GRUP_VEIC_LEGA").Text
                
                intCodGrupoVeicLegal = Val(objDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text)
                
                lngLinhaGrid = lngLinhaGrid + 1
                .MaxRows = lngLinhaGrid
                flFormartarNovaLinha
                
            End If
                     
            .BlockMode = False
            .Col = 2
            .Row = lngLinhaGrid
            .CellType = CellTypePicture
            .TypePictCenter = True
            .TypeHAlign = TypeHAlignCenter
            .TypePictPicture = imgOutrosIcones.ListImages("itemelementar").Picture
            
            .Col = 3
            .TypeHAlign = TypeHAlignLeft
            .SetText 3, lngLinhaGrid, objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
            
            .SetText 6, lngLinhaGrid, fgDtXML_To_Interface(strDataAbertura)
            .SetText 7, lngLinhaGrid, fgDescricaoEstadoCaixa(Val(strSituacaoCaixa))
            .SetText 8, lngLinhaGrid, fgVlrXml_To_Interface(strValorAbertura)
            If fgVlrXml_To_Decimal(strValorAbertura) < 0 Then
               .ForeColor = vbRed
            End If
            
            .SetText 9, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_PREV").Text)
                    
            If fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_PREV").Text) < 0 Then
                .Row = lngLinhaGrid
                .Col = 9
                .ForeColor = vbRed
            End If
            
            .SetText 10, lngLinhaGrid, _
                    fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_TOT_MOV").Text)
            
            If fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_TOT_MOV").Text) < 0 Then
                .Row = lngLinhaGrid
                .Col = 10
                .ForeColor = vbRed
            End If
            
            .SetText 11, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_SOLI").Text)
                    
            If fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_SOLI").Text) < 0 Then
               .Row = lngLinhaGrid
               .Col = 11
               .ForeColor = vbRed
            End If
            
            .SetText 12, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_CONF").Text)
                    
            If fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_CONF").Text) < 0 Then
               .Row = lngLinhaGrid
               .Col = 12
               .ForeColor = vbRed
            End If
            
            .SetText 13, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_VAR_MOV").Text)
                    
            If fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_VAR_MOV").Text) < 0 Then
               .Row = lngLinhaGrid
               .Col = 13
               .ForeColor = vbRed
            End If
            
            .SetText 14, lngLinhaGrid, fgVlrXml_To_Interface(fgVlrXml_To_Decimal(strValorAbertura) + _
                                                             fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_TOT_MOV").Text))
                           
            If fgVlrXml_To_Decimal(strValorAbertura) + _
               fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_TOT_MOV").Text) < 0 Then
                
                .Row = lngLinhaGrid
                .Col = 14
                .ForeColor = vbRed
            End If
            
            .SetText 15, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_TOT_SOLI").Text)
                           
            If fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_TOT_SOLI").Text) < 0 Then
                .Row = lngLinhaGrid
                .Col = 15
                .ForeColor = vbRed
            End If
            
            .SetText 16, lngLinhaGrid, fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_TOT_CONF").Text)
                           
            If fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_TOT_CONF").Text) < 0 Then
                .Row = lngLinhaGrid
                .Col = 16
                .ForeColor = vbRed
            End If
            
            .SetText 17, lngLinhaGrid, fgVlrXml_To_Interface(fgVlrXml_To_Decimal(strValorAbertura) + _
                                                             fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_VAR_MOV").Text))
                               
            If fgVlrXml_To_Decimal(strValorAbertura) + _
               fgVlrXml_To_Decimal(objDomNode.selectSingleNode("VA_VAR_MOV").Text) < 0 Then
               
                .Row = lngLinhaGrid
                .Col = 17
                .ForeColor = vbRed
            End If
            
            lngLinhaGrid = lngLinhaGrid + 1
            .MaxRows = lngLinhaGrid
            flFormartarNovaLinha
        Next
        
        lngLinhaFinalGrid = lngLinhaGrid
        
        .Redraw = True
    End With
    
    Set xmlLerTodos = Nothing
    Set objMonitoracaoFluxoCaixaReserva = Nothing

    Exit Sub

ErrorHandler:
    Set xmlLerTodos = Nothing
    Set objMonitoracaoFluxoCaixaReserva = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "frmSubReservaResumo - flCarregarSpread", 0

End Sub

Private Sub ctlTableCombo_AplicarFiltro(xmlDocFiltros As String)

Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim xmlDomLeitura                           As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    If Trim(strFiltroXML) = vbNullString Then
        Call flCarregarSpread(xmlDocFiltros)
        strFiltroXML = xmlDocFiltros
    Else
        Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomFiltro.loadXML(strFiltroXML)
        
        If Not xmlDomFiltro.selectSingleNode("//Grupo_GrupoVeiculoLegal") Is Nothing Then
            Call fgRemoveNode(xmlDomFiltro, "Grupo_GrupoVeiculoLegal")
        End If
        
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlDomLeitura.loadXML(xmlDocFiltros)
        
        Call fgAppendNode(xmlDomFiltro, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
        Call fgAppendNode(xmlDomFiltro, "Grupo_GrupoVeiculoLegal", "GrupoVeiculoLegal", _
                                        xmlDomLeitura.selectSingleNode("//GrupoVeiculoLegal").Text)
                                        
        Call flCarregarSpread(xmlDomFiltro.xml)
        strFiltroXML = xmlDomFiltro.xml
    End If
    
    Call fgCursor(False)
    
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlDomFiltro = Nothing
    Set xmlDomLeitura = Nothing
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaResumo - ctlTableCombo_AplicarFiltro"

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

   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaResumo - ctlTableCombo_MouseMove"
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler

    If KeyCode = vbKeyF5 Then
        fgCursor True
        Call flCarregarSpread(strFiltroXML)
        fgCursor
    End If

    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaResumo - Form_KeyDown"

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    Me.Icon = mdiSBR.Icon

    Call fgCenterMe(Me)
    Call fgCursor(True)
    
    flFormartarSpread
    lblData.Caption = Date
    ctlTableCombo.TituloCombo = strTableComboInicial
    DoEvents
    
    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA6.frmSubReservaResumo
    Load objFiltro
    
    Call objFiltro.fgCarregarPesquisaAnterior
    
    Me.Show
    DoEvents
    
    Call fgCursor(False)

    Exit Sub
ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaResumo - Form_Load"

End Sub

Private Sub Form_Resize()

On Error Resume Next

    With Me
        .ctlTableCombo.Left = 0
        .ctlTableCombo.Top = 0
        
        .lblBarra.Left = 0
        .lblBarra.Top = 0
        .lblBarra.Width = .ScaleWidth
        .lblBarra.Height = .ctlTableCombo.Height
        
        .lblData.Top = 0
        .lblData.Left = .ScaleWidth - .lblData.Width - 120
        .lblData.Height = .ctlTableCombo.Height

        .vasLista.Left = 0
        .vasLista.Top = .lblBarra.Height
        .vasLista.Width = .ScaleWidth
        .vasLista.Height = .ScaleHeight - .vasLista.Top - .tlbButtons.Height
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmSubReservaResumo = Nothing

End Sub

Private Sub lblBarra_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaResumo - lblBarra_MouseMove"
   
End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, lsTituloTableCombo As String)

On Error GoTo ErrorHandler

    Call fgCursor(True)
    Call flCarregarSpread(xmlDocFiltros)
    strFiltroXML = xmlDocFiltros
    Call fgCursor(False)

    ctlTableCombo.TituloCombo = IIf(lsTituloTableCombo = vbNullString, strTableComboInicial, lsTituloTableCombo)

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaResumo - objFiltro_AplicarFiltro"
    
End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    Select Case Button.Key
    Case "showfiltro"
        Set objFiltro = New frmFiltro
        Set objFiltro.FormOwner = Me
        objFiltro.TipoFiltro = enumTipoFiltroA6.frmSubReservaResumo
        objFiltro.Show vbModal

    Case "refresh"
        fgCursor True
        Call flCarregarSpread(strFiltroXML)
        fgCursor
        
    End Select

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaResumo - tlbButtons_ButtonClick"
   
End Sub

Private Sub vasLista_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    Call ctlTableCombo.fgMakeButtonFlat

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, "frmSubReservaResumo - vasLista_MouseMove"
   
End Sub

' Formata linha grid de pesquisa.

Private Sub flFormartarNovaLinha()

Dim lngColunas                              As Long

On Error GoTo ErrorHandler

    With Me.vasLista
       
        .BlockMode = True
        .Col = 1
        .Row = 3
        .Col2 = 4
        .Row2 = .MaxRows
        .AllowCellOverflow = True
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = 4
        .Row2 = .MaxRows
        .BackColorStyle = BackColorStyleOverVertGridOnly
        .BackColor = &H8000000E
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = 1
        .BackColor = vbBlack
        .ForeColor = vbWhite
        .RowHeight(1) = 300
        .FontSize = 10
        .FontBold = True
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Row = 2
        .Col2 = .MaxCols
        .Row2 = 2
        .BackColor = vbBlack
        .ForeColor = vbWhite
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Row = 3
        .Col2 = .MaxCols
        .Row2 = 3
        .BackColor = &H8000000F
        .ForeColor = vbAutomatic
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
        .BlockMode = True
        .Col = 6
        .Row = 4
        .Col2 = .MaxCols
        .Row2 = .MaxRows
        .TypeHAlign = TypeHAlignRight
        .BlockMode = False
        
        .BlockMode = True
        .Col = 6
        .Row = 4
        .Col2 = 8
        .Row2 = .MaxRows
        .BackColor = &HC0FFFF
        .BlockMode = False
        
        .BlockMode = True
        .Col = 9
        .Row = 4
        .Col2 = 13
        .Row2 = .MaxRows
        .BackColor = &HC0FFC0
        .BlockMode = False
        
        .BlockMode = True
        .Col = 14
        .Row = 4
        .Col2 = 17
        .Row2 = .MaxRows
        .BackColor = &HFFFFC0
        .BlockMode = False
        
    End With

    Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flFormartarNovaLinha", 0

End Sub

