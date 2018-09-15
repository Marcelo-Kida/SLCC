VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFiltro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtro"
   ClientHeight    =   3240
   ClientLeft      =   4665
   ClientTop       =   3705
   ClientWidth     =   6645
   Icon            =   "frmFiltro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraFiltro 
      Height          =   2820
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   6555
      Begin VB.ComboBox cboTipoBackoffice 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2010
         Width           =   4695
      End
      Begin VB.TextBox txtVeiculoLegal 
         Height          =   315
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1650
         Width           =   1455
      End
      Begin VB.ComboBox cboVeiculoLegal 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1650
         Width           =   3195
      End
      Begin VB.ComboBox cboGrupoVeicLegal 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1290
         Width           =   4695
      End
      Begin VB.ComboBox cboSistema 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   570
         Width           =   4695
      End
      Begin VB.ComboBox cboBancLiqu 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   4695
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   2940
         TabIndex        =   2
         Top             =   930
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Format          =   54460417
         CurrentDate     =   37322
      End
      Begin MSComCtl2.DTPicker dtpFim 
         Height          =   315
         Left            =   4770
         TabIndex        =   3
         Top             =   930
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         Format          =   54460417
         CurrentDate     =   37322
      End
      Begin MSComctlLib.Toolbar tlbData 
         Height          =   330
         Left            =   1740
         TabIndex        =   8
         ToolTipText     =   "Clique para Habilitar/Desabilitar"
         Top             =   915
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         ButtonWidth     =   1349
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Após"
               Key             =   "Comparacao"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Antes"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Entre"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTipoBackoffice 
         Caption         =   "Tipo de Backoffice"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   2070
         Width           =   1455
      End
      Begin VB.Label lblSituacaoCaixa 
         AutoSize        =   -1  'True
         Caption         =   "Aberto | Fechado | Disponível | Inexistente"
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
         Left            =   1800
         TabIndex        =   16
         Top             =   2430
         Width           =   4380
      End
      Begin VB.Label lblDescSituacaoCaixa 
         AutoSize        =   -1  'True
         Caption         =   "Situação do Caixa"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   2460
         Width           =   1290
      End
      Begin VB.Label lblVeiculoLegal 
         AutoSize        =   -1  'True
         Caption         =   "Veículo Legal"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   1710
         Width           =   990
      End
      Begin VB.Label lblGrupoVeicLegal 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Veículo Legal"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   1350
         Width           =   1470
      End
      Begin VB.Label lblBancLiqu 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   990
         Width           =   345
      End
      Begin VB.Label lblSistema 
         AutoSize        =   -1  'True
         Caption         =   "Sistema"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   630
         Width           =   555
      End
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   3540
      Top             =   2670
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
            Picture         =   "frmFiltro.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltro.frx":045E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Height          =   330
      Left            =   4170
      TabIndex        =   14
      Top             =   2880
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   582
      ButtonWidth     =   2011
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Aplicar"
            Key             =   "aplicar"
            Object.ToolTipText     =   "Aplicar Filtro"
            ImageIndex      =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Filtro"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário o filtro dos dados a serem apresentados nas telas de consulta.

Option Explicit

Public Event AplicarFiltro(xmlDocFiltros As String, strTituloTableCombo As String)

Public TipoFiltro                           As enumTipoFiltroA6
Public FormOwner                            As Form

Private Const PROP_TOP_CAMPO_01             As Integer = 240
Private Const PROP_TOP_CAMPO_02             As Integer = 600
Private Const PROP_TOP_CAMPO_03             As Integer = 960
Private Const PROP_TOP_CAMPO_04             As Integer = 1320
Private Const PROP_TOP_CAMPO_05             As Integer = 1680
Private Const PROP_TOP_CAMPO_06             As Integer = 2040
Private Const PROP_TOP_CAMPO_07             As Integer = 2400

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlPropriedades                     As MSXML2.DOMDocument40
Private xmlDOMRegistro                      As MSXML2.DOMDocument40
Private xmlDomFiltroApoio                   As MSXML2.DOMDocument40
Private lcControlesFiltro                   As New Collection

Private strFuncionalidade                   As String
Private arrSgSistCoVeicLega()               As Variant

Private blnPrimeiroActivate                 As Boolean
Private blnFiltraHora                       As Boolean
Private blnExibirTipoBackOffice             As Boolean

'Carregar dinamicamente apenas os combos utilizados no filtro

Private Sub flCarregarCombosUtilizados()

    If cboBancLiqu.Top < PROP_TOP_CAMPO_07 Then
        fgCarregarCombos cboBancLiqu, gxmlCombosFiltro, "Empresa", "CO_EMPR", "NO_REDU_EMPR", True
    End If
    
    If cboGrupoVeicLegal.Top < PROP_TOP_CAMPO_07 Then
        fgCarregarCombos cboGrupoVeicLegal, gxmlCombosFiltro, "GrupoVeiculoLegal", "CO_GRUP_VEIC_LEGA", "NO_GRUP_VEIC_LEGA", True
    End If
    
    If cboTipoBackoffice.Top < PROP_TOP_CAMPO_07 Then
        fgCarregarCombos cboTipoBackoffice, gxmlCombosFiltro, "TipoBackOffice", "TP_BKOF", "DE_BKOF", True
    End If

    Exit Sub

ErrorHandler:
    Call mdiSBR.uctLogErros.MostrarErros(Err, Me.Name & " - flCarregarCombosUtilizados", Me.Caption)

End Sub

' Aciona a última pesquisa efetuada para o objeto chamador do filtro.

Public Sub fgCarregarPesquisaAnterior()

On Error GoTo ErrorHandler
    
    If xmlDOMRegistro.xml <> vbNullString Then
        Call flAplicarFiltro(False)
    End If

    Exit Sub
    
ErrorHandler:
    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmFiltro - flCarregarPesquisaAnterior")
    
End Sub

' Formata data e hora para pesquisa no Oracle.

Private Function flConverteDataHoraOracle(ByVal strDataHora As String) As String

Dim strAno                                  As String
Dim strMes                                  As String
Dim strDia                                  As String
Dim strHora                                 As String
Dim strMinuto                               As String
Dim strSegundo                              As String
Dim strDataHoraConvertida                   As String

On Error GoTo ErrorHandler

    strAno = Format(Year(strDataHora), "0000")
    strMes = Format(Month(strDataHora), "00")
    strDia = Format(Day(strDataHora), "00")
    strHora = Format(Hour(strDataHora), "00")
    strMinuto = Format(Minute(strDataHora), "00")
    strSegundo = Format(Second(strDataHora), "00")
    
    strDataHoraConvertida = strAno & strMes & strDia & strHora & strMinuto & strSegundo
    
    flConverteDataHoraOracle = "TO_DATE('" & strDataHoraConvertida & "','YYYYMMDDHH24MISS')"

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flConverteDataHoraOracle", 0

End Function

' Formata data para pesquisa no Oracle.

Private Function flConverteDataOracle(ByVal strData As String) As String

Dim strAno                                  As String
Dim strMes                                  As String
Dim strDia                                  As String
Dim strDataConvertida                       As String

On Error GoTo ErrorHandler

    strAno = Format(Year(strData), "0000")
    strMes = Format(Month(strData), "00")
    strDia = Format(Day(strData), "00")
    
    strDataConvertida = strAno & strMes & strDia
    
    flConverteDataOracle = "TO_DATE('" & strDataConvertida & "','YYYYMMDD')"

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flConverteDataOracle", 0

End Function

' Aciona a aplicação do filtro selecionado pelo usuário.

Private Sub flAplicarFiltro(ByVal pblnNovaPesquisa As Boolean)

Dim strTituloTableCombo                     As String
Dim strDocFiltros                           As String
Dim xmlDomFiltros                           As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler
    
    If xmlDOMRegistro.xml = vbNullString Or pblnNovaPesquisa Then
        Set xmlDomFiltros = flInterfaceToXml(strTituloTableCombo)
        Call flGravarSettingsRegistry(xmlDomFiltros, strTituloTableCombo)
        
        If Not xmlDomFiltroApoio Is Nothing Then
            If xmlDomFiltroApoio.xml <> vbNullString Then
                If Not xmlDomFiltroApoio.selectSingleNode("Repeat_Apoio/Grupo_Sistema") Is Nothing Then
                    Call fgAppendXML(xmlDomFiltros, "Repeat_Filtros", _
                                                    xmlDomFiltroApoio.selectSingleNode("Repeat_Apoio/Grupo_Sistema").xml)
                End If
                
                If Not xmlDomFiltroApoio.selectSingleNode("Repeat_Apoio/Grupo_Data") Is Nothing Then
                    Call fgAppendXML(xmlDomFiltros, "Repeat_Filtros", _
                                                    xmlDomFiltroApoio.selectSingleNode("Repeat_Apoio/Grupo_Data").xml)
                End If
            End If
        End If
        
        strDocFiltros = xmlDomFiltros.xml
    Else
        strDocFiltros = xmlDOMRegistro.selectSingleNode("//Registry/Repeat_Filtros").xml
        strTituloTableCombo = xmlDOMRegistro.selectSingleNode("//Registry/TituloTableCombo").Text
    End If

    Me.Hide
    DoEvents
    
    If pblnNovaPesquisa Then
        RaiseEvent AplicarFiltro(strDocFiltros, strTituloTableCombo)
    End If
                                           
    Set xmlDomFiltros = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlDomFiltros = Nothing
   
    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmFiltro - flAplicarFiltro")

End Sub

' Converte os dados informados na tela em formato XML.

Private Function flInterfaceToXml(ByRef strTituloTableCombo As String) As MSXML2.DOMDocument40

Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim xmlDomControle                          As MSXML2.DOMDocument40
Dim xmlDomAux                               As MSXML2.DOMDocument40

Dim strDocFiltros                           As String
Dim intIndCombo                             As Integer

Dim strSiglaSistema                         As String
Dim strCodVeiculoLegal                      As String

    On Error GoTo ErrorHandler
    
    strTituloTableCombo = vbNullString
            
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    
    Set xmlDomControle = CreateObject("MSXML2.DOMDocument.4.0")
    
    Select Case TipoFiltro
        Case enumTipoFiltroA6.frmSubReservaResumo
            
            If cboGrupoVeicLegal.ListIndex > -1 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboGrupoVeicLegal.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
                
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                                 "VeiculoLegal", Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(1))
            End If
            
            If blnExibirTipoBackOffice Then
                If cboTipoBackoffice.ListIndex >= 0 Then
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BackOfficePerfilGeral", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_BackOfficePerfilGeral", _
                                                     "BackOfficePerfilGeral", IIf(cboTipoBackoffice.ListIndex = 0, 0, fgObterCodigoCombo(Me.cboTipoBackoffice.Text)))
                End If
            End If
                       
        Case enumTipoFiltroA6.frmSubReservaConsultaAberturaFechamento
            
            If cboGrupoVeicLegal.ListIndex > -1 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboGrupoVeicLegal.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
                
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                                 "VeiculoLegal", Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(1))
            End If
            
            If blnExibirTipoBackOffice Then
                If cboTipoBackoffice.ListIndex >= 0 Then
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BackOfficePerfilGeral", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_BackOfficePerfilGeral", _
                                                     "BackOfficePerfilGeral", IIf(cboTipoBackoffice.ListIndex = 0, 0, fgObterCodigoCombo(Me.cboTipoBackoffice.Text)))
                End If
            End If
                       
            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                    Case "Após"
                    
                         Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                         
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataIni", flConverteDataOracle(dtpInicio.Value))
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataFim", flConverteDataOracle("31/12/9999"))
                    
                    Case "Antes"
                    
                         Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                         
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataIni", flConverteDataOracle("01/01/1900"))
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataFim", flConverteDataOracle(dtpInicio.Value))
                    
                    Case "Entre"
                    
                         Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                         
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataIni", flConverteDataOracle(dtpInicio.Value))
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataFim", flConverteDataOracle(dtpFim.Value))
                    
                End Select
                
            End If
            
        Case enumTipoFiltroA6.frmControleRemessa
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            
            End If
            
            If cboSistema.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Sistema", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Sistema", _
                                                 "Sistema", fgObterCodigoCombo(Me.cboSistema.Text))
            End If
            
            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                    Case "Após"
                    
                         Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                         
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataIni", flConverteDataOracle(dtpInicio.Value))
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataFim", flConverteDataOracle("31/12/9999"))
                    
                    Case "Antes"
                    
                         Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                         
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataIni", flConverteDataOracle("01/01/1900"))
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataFim", flConverteDataOracle(dtpInicio.Value))
                    
                    Case "Entre"
                    
                         Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                         
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataIni", flConverteDataOracle(dtpInicio.Value))
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataFim", flConverteDataOracle(dtpFim.Value))
                    
                End Select
                
            End If
        
        Case enumTipoFiltroA6.frmSubReservaD0
            
            Set xmlDomFiltroApoio = CreateObject("MSXML2.DOMDocument.4.0")
            Call fgAppendNode(xmlDomFiltroApoio, "", "Repeat_Apoio", "")
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
                
            If cboGrupoVeicLegal.ListIndex > -1 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
                
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                                 "VeiculoLegal", Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(1))
            
                Call fgAppendNode(xmlDomFiltroApoio, "Repeat_Apoio", "Grupo_Sistema", "")
                Call fgAppendNode(xmlDomFiltroApoio, "Grupo_Sistema", _
                                                     "Sistema", Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(0))
            End If
            
            If Not IsNull(dtpInicio.Value) And dtpInicio.Enabled Then
                'Para este caso as datas de início e fim são iguais
                Call fgAppendNode(xmlDomFiltroApoio, "Repeat_Apoio", "Grupo_Data", "")
                Call fgAppendNode(xmlDomFiltroApoio, "Grupo_Data", _
                                                     "DataIni", flConverteDataOracle(dtpInicio.Value))
                Call fgAppendNode(xmlDomFiltroApoio, "Grupo_Data", _
                                                     "DataFim", flConverteDataOracle(dtpInicio.Value))
            End If
            
            If blnExibirTipoBackOffice Then
                If cboTipoBackoffice.ListIndex >= 0 Then
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BackOfficePerfilGeral", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_BackOfficePerfilGeral", _
                                                     "BackOfficePerfilGeral", IIf(cboTipoBackoffice.ListIndex = 0, 0, fgObterCodigoCombo(Me.cboTipoBackoffice.Text)))
                End If
            End If
                       
        Case enumTipoFiltroA6.frmRemessaRejeitada
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            If cboSistema.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Sistema", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Sistema", _
                                                 "Sistema", fgObterCodigoCombo(Me.cboSistema.Text))
            End If
        
            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                    Case "Após"
                    
                         Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                         
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataIni", flConverteDataOracle(dtpInicio.Value))
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataFim", flConverteDataOracle("31/12/9999"))
                    
                    Case "Antes"
                    
                         Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                         
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataIni", flConverteDataOracle("01/01/1900"))
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataFim", flConverteDataOracle(dtpInicio.Value))
                    
                    Case "Entre"
                    
                         Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                         
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataIni", flConverteDataOracle(dtpInicio.Value))
                         Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                          "DataFim", flConverteDataOracle(dtpFim.Value))
                    
                End Select
                
            End If
        
        Case enumTipoFiltroA6.frmCaixaFuturo
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
                
            If cboGrupoVeicLegal.ListIndex > -1 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
                
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                                 "VeiculoLegal", Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(1))
            End If
        
            If blnExibirTipoBackOffice Then
                If cboTipoBackoffice.ListIndex >= 0 Then
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BackOfficePerfilGeral", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_BackOfficePerfilGeral", _
                                                     "BackOfficePerfilGeral", IIf(cboTipoBackoffice.ListIndex = 0, 0, fgObterCodigoCombo(Me.cboTipoBackoffice.Text)))
                End If
            End If
                       
        Case enumTipoFiltroA6.frmSubReservaAbertura
            
            If cboGrupoVeicLegal.ListIndex > -1 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboGrupoVeicLegal.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
                
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                                 "VeiculoLegal", Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(1))
            End If
    
        Case enumTipoFiltroA6.frmSubReservaFechamento
            
            If cboGrupoVeicLegal.ListIndex > -1 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboGrupoVeicLegal.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
                
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                                 "VeiculoLegal", Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(1))
            End If
        
            Select Case tlbData.Buttons(1).Caption
            
                Case "Após"
                
                     Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                     
                     Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                      "DataIni", flConverteDataOracle(dtpInicio.Value))
                     Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                      "DataFim", flConverteDataOracle("31/12/9999"))
                
                Case "Antes"
                
                     Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                     
                     Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                      "DataIni", flConverteDataOracle("01/01/1900"))
                     Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                      "DataFim", flConverteDataOracle(dtpInicio.Value))
                
                Case "Entre"
                
                     Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                     
                     Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                      "DataIni", flConverteDataOracle(dtpInicio.Value))
                     Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                      "DataFim", flConverteDataOracle(dtpFim.Value))
                
            End Select
    End Select
                                           
    Set flInterfaceToXml = xmlDomFiltros
                                           
    Set xmlDomFiltros = Nothing
    Set xmlDomControle = Nothing
    Set xmlDomAux = Nothing
    
    Exit Function

ErrorHandler:
    Set xmlDomFiltros = Nothing
    Set xmlDomControle = Nothing
    Set xmlDomAux = Nothing
    
    fgRaiseError App.EXEName, "frmFiltro", "flInterfaceToXml", 0

End Function

' Atribui o conteúdo extraído do Registry do Windows, aos campos da tela.

Private Sub flAplicarSettingsRegistry()
    
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strRegistry                             As String
Dim intIndControles                         As Integer
    
On Error GoTo ErrorHandler
    
    If xmlDOMRegistro.xml = vbNullString Then Exit Sub
    
    intIndControles = 0
    For Each objDomNode In xmlDOMRegistro.documentElement.selectNodes("//Registry/Grupo_ControleFiltro/*")
        
        intIndControles = intIndControles + 1
        If intIndControles > lcControlesFiltro.Count Then Exit For
        
        If TypeName(lcControlesFiltro(intIndControles)) = "ComboBox" Then
            If lcControlesFiltro(intIndControles).ListCount > 0 Then
                If lcControlesFiltro(intIndControles).ListCount > Val("0" & objDomNode.Text) Then
                    lcControlesFiltro(intIndControles).ListIndex = IIf(objDomNode.Text = vbNullString Or Val(objDomNode.Text) < 0, -1, objDomNode.Text)
                Else
                    lcControlesFiltro(intIndControles).ListIndex = -1
                End If
                
            'Foi acrescentada a verificação abaixo para forçar o carregamento do combo de Veículo Legal,
            'já que este, por um problema de performance, será carregado apenas no dropdown, e não mais no
            'click do Grupo de Veículo Legal
            ElseIf lcControlesFiltro(intIndControles).Name = "cboVeiculoLegal" Then
                If objDomNode.Text <> vbNullString And objDomNode.Text <> "-1" Then
                    Call cboVeiculoLegal_DropDown
                    If Val(objDomNode.Text) <= lcControlesFiltro(intIndControles).ListCount - 1 Then
                        If lcControlesFiltro(intIndControles).ListCount > Val("0" & objDomNode.Text) Then
                            lcControlesFiltro(intIndControles).ListIndex = IIf(objDomNode.Text = vbNullString Or Val(objDomNode.Text) = 0, -1, objDomNode.Text)
                        Else
                            lcControlesFiltro(intIndControles).ListIndex = -1
                        End If
                    End If
                End If
            End If
        ElseIf TypeName(lcControlesFiltro(intIndControles)) = "TextBox" Then
            lcControlesFiltro(intIndControles).Text = objDomNode.Text
        ElseIf TypeName(lcControlesFiltro(intIndControles)) = "DTPicker" Then
            If IsDate(objDomNode.Text) Then
                lcControlesFiltro(intIndControles).Value = fgValidarMinDateDTPicker(lcControlesFiltro(intIndControles), fgDtXML_To_Date(objDomNode.Text))
            Else
                lcControlesFiltro(intIndControles).Value = fgDataHoraServidor(DataAux)
            End If
        ElseIf TypeName(lcControlesFiltro(intIndControles)) = "Toolbar" Then
            lcControlesFiltro(intIndControles).Buttons(1).Image = Val(Split(objDomNode.Text, "|")(1))
            If lcControlesFiltro(intIndControles).Name = "tlbData" Then
                blnFiltraHora = lcControlesFiltro(intIndControles).Buttons(1).Image <> 0
                flConfiguratlbData Split(objDomNode.Text, "|")(0), tlbData, dtpFim
            End If
        ElseIf TypeName(lcControlesFiltro(intIndControles)) = "Number" Then
            lcControlesFiltro(intIndControles).Valor = fgVlrXml_To_Decimal(objDomNode.Text)
        End If
    
    Next
    
    Set objDomNode = Nothing

    Exit Sub
    
ErrorHandler:
    Set objDomNode = Nothing
    
    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmFiltro - flAplicarSettingsRegistry")

End Sub

' Grava o conteúdo dos campos da tela no Registry do Windows.

Private Sub flGravarSettingsRegistry(ByVal objDOMFiltro As MSXML2.DOMDocument40, _
                                     ByVal strTituloTableCombo As String)

Dim objDomGrupoControle                     As MSXML2.DOMDocument40
Dim objControleFiltro                       As Object
    
    On Error GoTo ErrorHandler
    
    Set xmlDOMRegistro = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlDOMRegistro, "", "Registry", "")

    Set objDomGrupoControle = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(objDomGrupoControle, "", "Grupo_ControleFiltro", "")

    For Each objControleFiltro In lcControlesFiltro
        
        If TypeName(objControleFiltro) = "ComboBox" Then
            If objControleFiltro.Name <> cboTipoBackoffice.Name Then
                Call fgAppendNode(objDomGrupoControle, "Grupo_ControleFiltro", "ConteudoControle", objControleFiltro.ListIndex)
            End If
        ElseIf TypeName(objControleFiltro) = "TextBox" Then
            Call fgAppendNode(objDomGrupoControle, "Grupo_ControleFiltro", "ConteudoControle", objControleFiltro.Text)
        ElseIf TypeName(objControleFiltro) = "DTPicker" Then
            'Se o filtro estiver relacionado à rotina de D0,
            'não é necessário gravar as configurações de DATA no Registry
            If strFuncionalidade <> "frmFiltro_SubReservaD0" Then
                Call fgAppendNode(objDomGrupoControle, "Grupo_ControleFiltro", "ConteudoControle", fgDt_To_Xml(objControleFiltro.Value))
            End If
        ElseIf TypeName(objControleFiltro) = "Toolbar" Then
            Call fgAppendNode(objDomGrupoControle, "Grupo_ControleFiltro", "ConteudoControle", objControleFiltro.Buttons(1).Caption & "|" & objControleFiltro.Buttons(1).Image)
        ElseIf TypeName(objControleFiltro) = "Number" Then
            Call fgAppendNode(objDomGrupoControle, "Grupo_ControleFiltro", "ConteudoControle", fgVlr_To_Xml(objControleFiltro.Valor))
        End If
    
    Next
    
    Call fgAppendXML(xmlDOMRegistro, "Registry", objDomGrupoControle.xml)
    Call fgAppendXML(xmlDOMRegistro, "Registry", objDOMFiltro.xml)
    Call fgAppendNode(xmlDOMRegistro, "Registry", "TituloTableCombo", strTituloTableCombo)
    
    Set objDomGrupoControle = Nothing
    
    Call SaveSetting("A6SBR", "Form Filtro\" & FormOwner.Name, "Settings", xmlDOMRegistro.xml)
    
    Exit Sub
    
ErrorHandler:
    Set objDomGrupoControle = Nothing

    Call mdiSBR.uctLogErros.MostrarErros(Err, "frmFiltro - flGravarSettingsRegistry")

End Sub

' Configura o layout do formulário de filtro, de acordo com o objeto chamador.

Private Sub flConfiguraLayoutForm()

Dim colControles                            As New Collection
Dim objControle                             As Object

    On Error GoTo ErrorHandler

    With Me
        Set colControles = New Collection
        colControles.Add .lblBancLiqu
        colControles.Add .lblSistema
        colControles.Add .lblData
        colControles.Add .lblGrupoVeicLegal
        colControles.Add .lblVeiculoLegal
        colControles.Add .lblTipoBackoffice
        colControles.Add .lblDescSituacaoCaixa
        
        For Each objControle In colControles
            objControle.Top = PROP_TOP_CAMPO_07 * 3
        Next
        
        Select Case TipoFiltro
            Case enumTipoFiltroA6.frmSubReservaResumo
                
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_01
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_02
                
                If blnExibirTipoBackOffice Then
                    .lblTipoBackoffice.Top = PROP_TOP_CAMPO_03
                    .fraFiltro.Height = PROP_TOP_CAMPO_04
                Else
                    .fraFiltro.Height = PROP_TOP_CAMPO_03
                End If
                
            Case enumTipoFiltroA6.frmControleRemessa, _
                 enumTipoFiltroA6.frmRemessaRejeitada
                
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblSistema.Top = PROP_TOP_CAMPO_02
                .lblData.Top = PROP_TOP_CAMPO_03
                
                .fraFiltro.Height = PROP_TOP_CAMPO_04
                
            Case enumTipoFiltroA6.frmSubReservaD0
                
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblData.Top = PROP_TOP_CAMPO_02
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_03
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_04
                
                If blnExibirTipoBackOffice Then
                    .lblTipoBackoffice.Top = PROP_TOP_CAMPO_05
                    .lblDescSituacaoCaixa.Top = PROP_TOP_CAMPO_06
                    .fraFiltro.Height = PROP_TOP_CAMPO_07
                Else
                    .lblDescSituacaoCaixa.Top = PROP_TOP_CAMPO_05
                    .fraFiltro.Height = PROP_TOP_CAMPO_06
                End If
                
            Case enumTipoFiltroA6.frmCaixaFuturo
                
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_02
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_03
                
                If blnExibirTipoBackOffice Then
                    .lblTipoBackoffice.Top = PROP_TOP_CAMPO_04
                    .fraFiltro.Height = PROP_TOP_CAMPO_05
                Else
                    .fraFiltro.Height = PROP_TOP_CAMPO_04
                End If
                
            Case enumTipoFiltroA6.frmSubReservaAbertura, _
                 enumTipoFiltroA6.frmSubReservaFechamento
                
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_01
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_02
                
                .fraFiltro.Height = PROP_TOP_CAMPO_03
                
            Case enumTipoFiltroA6.frmSubReservaConsultaAberturaFechamento
                
                .lblData.Top = PROP_TOP_CAMPO_01
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_02
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_03
                
                If blnExibirTipoBackOffice Then
                    .lblTipoBackoffice.Top = PROP_TOP_CAMPO_04
                    .fraFiltro.Height = PROP_TOP_CAMPO_05
                Else
                    .fraFiltro.Height = PROP_TOP_CAMPO_04
                End If
                
        End Select
        
        .cboBancLiqu.Top = .lblBancLiqu.Top - 60
        .cboSistema.Top = .lblSistema.Top - 60
        .dtpInicio.Top = .lblData.Top - 60
        .tlbData.Top = .lblData.Top - 60
        .cboGrupoVeicLegal.Top = .lblGrupoVeicLegal.Top - 60
        .cboVeiculoLegal.Top = .lblVeiculoLegal.Top - 60
        .txtVeiculoLegal.Top = .lblVeiculoLegal.Top - 60
        .cboTipoBackoffice.Top = .lblTipoBackoffice.Top - 60
        .lblSituacaoCaixa.Top = .lblDescSituacaoCaixa.Top - 60
    
        .dtpFim.Top = .lblData.Top - 60
        
        .tlbComandos.Top = .fraFiltro.Top + .fraFiltro.Height + 30
        .Height = (.Height - .ScaleHeight) + .tlbComandos.Top + .tlbComandos.Height
            
    End With

    Call flCarregarCombosUtilizados
    
    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, "frmFiltro", "flConfiguraLayoutForm", 0

End Sub

' Carrega configurações iniciais do formulário.

Private Sub flInit()

    cboSistema.AddItem "<-- Todos -->"
    cboSistema.ListIndex = 0
    cboSistema.Enabled = False
    
    cboVeiculoLegal.Enabled = False
    txtVeiculoLegal.Enabled = False
    
End Sub

' Valida campos antes da aplicação do filtro.

Private Function flValidarCampos() As Boolean

    flValidarCampos = False

    If cboGrupoVeicLegal.Top < PROP_TOP_CAMPO_07 And (cboGrupoVeicLegal.Text = vbNullString Or cboGrupoVeicLegal.ListIndex <= 0) Then
        frmMural.Display = "Selecione um Grupo de Veículo Legal."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Function
    End If
    
    flValidarCampos = True

End Function

Private Sub cboBancLiqu_Click()

On Error GoTo ErrorHandler

    If cboBancLiqu.ListIndex = 0 Then ' Todos
        cboSistema.Clear
        cboSistema.AddItem "<-- Todos -->"
        cboSistema.ListIndex = 0
        cboSistema.Enabled = False
    Else
        Call flCarregarSistema
    End If

Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - cboBancLiqu_Click"
End Sub

Private Sub cboGrupoVeicLegal_Click()

On Error GoTo ErrorHandler

    If cboGrupoVeicLegal.ListIndex >= 0 Then
        cboVeiculoLegal.Clear
        txtVeiculoLegal.Text = vbNullString
        txtVeiculoLegal.Enabled = True
        cboVeiculoLegal.Enabled = True
    End If

    'Verifica se o filtro é referente à rotina SubReservaD0
    If strFuncionalidade = "frmFiltro_SubReservaD0" Then
        dtpInicio.Enabled = False
        lblSituacaoCaixa.Caption = vbNullString
    End If

Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - cboVeiculoLegal_DropDown"

End Sub

Private Sub cboVeiculoLegal_Click()

#If EnableSoap = 1 Then
    Dim objCaixaSubReserva      As MSSOAPLib30.SoapClient30
#Else
    Dim objCaixaSubReserva      As A6MIU.clsCaixaSubReserva
#End If

Dim xmlRetorno                  As MSXML2.DOMDocument40
Dim intSituacaoCaixa            As enumEstadoCaixa
Dim strDataCaixa                As String
Dim strDataCaixaMax             As String
Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    If cboVeiculoLegal.ListIndex > 0 Then
        If UCase$(txtVeiculoLegal.Text) <> UCase$(Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(1)) Then
            txtVeiculoLegal.Text = Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(1)
        End If
    Else
        If Not Me.ActiveControl Is txtVeiculoLegal Then
            txtVeiculoLegal.Text = vbNullString
        End If
    End If
    
    'Verifica se o filtro é referente à rotina SubReservaD0...
    If strFuncionalidade = "frmFiltro_SubReservaD0" Then
        '...se sim, verifica se algum veículo legal foi selecionado...
        If cboVeiculoLegal.ListIndex > 0 Then
            '...se sim, pesquisa a posição do caixa para o veículo legal selecionado
            Set objCaixaSubReserva = fgCriarObjetoMIU("A6MIU.clsCaixaSubReserva")
            Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
            
            If xmlRetorno.loadXML( _
                objCaixaSubReserva.ObterPosicaoCaixaSubReserva(Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(1), _
                                                               Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(0), _
                                                               vbNullString, _
                                                               True, _
                                                               vntCodErro, _
                                                               vntMensagemErro)) Then
                
                If vntCodErro <> 0 Then
                    GoTo ErrorHandler
                End If
                
                intSituacaoCaixa = _
                    Val(xmlRetorno.documentElement.selectSingleNode("Grupo_PosicaoCaixaSubReserva/CO_SITU_CAIX_SUB_RESE_ATUAL").Text)
                
                strDataCaixa = _
                    xmlRetorno.documentElement.selectSingleNode("Grupo_PosicaoCaixaSubReserva/DT_CAIX_SUB_RESE_ATUAL").Text
                
                strDataCaixaMax = _
                    xmlRetorno.documentElement.selectSingleNode("Grupo_PosicaoCaixaSubReserva/DT_CAIX_SUB_RESE").Text
                
                With dtpInicio
                    .Enabled = True
                    .MinDate = DateSerial(1900, 1, 1)
                    .MaxDate = DateSerial(2099, 12, 31)
                    .Value = DateSerial(Mid$(strDataCaixa, 1, 4), _
                                        Mid$(strDataCaixa, 5, 2), _
                                        Mid$(strDataCaixa, 7, 2))
                    .Tag = strDataCaixa & ";" & intSituacaoCaixa
                    
                    'Configura o período de datas de acordo com a situação do caixa
                    Select Case intSituacaoCaixa
                        'Se ABERTO, Máxima Data = Data da Posição do Caixa
                        '           Mínima Data = Livre
                        Case enumEstadoCaixa.Aberto
                            .MinDate = DateSerial(1900, 1, 1)
                            .MaxDate = .Value
                        
                        'Se FECHADO, Máxima Data = 1º dia útil após a Data da Posição do Caixa
                        '            Mínima Data = Livre
                        Case enumEstadoCaixa.Fechado
                            .MinDate = DateSerial(1900, 1, 1)
                            .MaxDate = DateSerial(Mid$(strDataCaixaMax, 1, 4), _
                                                  Mid$(strDataCaixaMax, 5, 2), _
                                                  Mid$(strDataCaixaMax, 7, 2))
                                                  
                            'Verifica se a data atual é maior que a data da situação do caixa...
                            If fgDataHoraServidor(DataAux) > DateSerial(Mid$(strDataCaixa, 1, 4), _
                                                                        Mid$(strDataCaixa, 5, 2), _
                                                                        Mid$(strDataCaixa, 7, 2)) Then
                                '...se sim, sugere a posição DISPONÍVEL
                                .Value = .MaxDate
                                intSituacaoCaixa = enumEstadoCaixa.Disponivel
                            End If
                        
                        'Se DISPONÍVEL, Máxima Data e Mínima Data = Data da Posição do Caixa
                        Case enumEstadoCaixa.Disponivel
                            .MinDate = .Value
                            .MaxDate = .Value
                        
                    End Select
                End With
                
                lblSituacaoCaixa.Caption = fgDescricaoEstadoCaixa(intSituacaoCaixa)
            Else
                dtpInicio.Enabled = False
                lblSituacaoCaixa.Caption = vbNullString
            End If
            
            Set xmlRetorno = Nothing
            Set objCaixaSubReserva = Nothing
        Else
            dtpInicio.Enabled = False
            lblSituacaoCaixa.Caption = vbNullString
        End If
    End If
    
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:
    Set objCaixaSubReserva = Nothing
    Set xmlRetorno = Nothing
    
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - cboVeiculoLegal_Click"
End Sub

Private Sub cboVeiculoLegal_DropDown()

    On Error GoTo ErrorHandler

    If cboGrupoVeicLegal.ListIndex >= 0 And cboVeiculoLegal.ListCount = 0 Then
        fgCursor True
        Call fgLerCarregarVeiculoLegal(cboGrupoVeicLegal, cboVeiculoLegal, xmlPropriedades, arrSgSistCoVeicLega)
        fgCursor
    End If
    
    Exit Sub

ErrorHandler:
    fgCursor
    mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - cboVeiculoLegal_DropDown"

End Sub

Private Sub dtpFim_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub dtpInicio_Change()

#If EnableSoap = 1 Then
    Dim objCaixaSubReserva      As MSSOAPLib30.SoapClient30
#Else
    Dim objCaixaSubReserva      As A6MIU.clsCaixaSubReserva
#End If

Dim xmlRetorno                  As MSXML2.DOMDocument40
Dim strData                     As String
Dim intSituacaoCaixa            As Integer
Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    'Verifica se o filtro é referente à rotina SubReservaD0
    If strFuncionalidade = "frmFiltro_SubReservaD0" Then
        strData = Format$(dtpInicio.Year, "0000") & _
                  Format$(dtpInicio.Month, "00") & _
                  Format$(dtpInicio.Day, "00")
                 
        If strData = vbNullString Then
            Call fgCursor(False)
            
            Exit Sub
        End If
        
        'Verifica se a data informada é maior do que a data da situação do caixa,
        'armazenada na propriedade TAG...
        If DateSerial(Format$(dtpInicio.Year, "0000"), Format$(dtpInicio.Month, "00"), Format$(dtpInicio.Day, "00")) > _
           DateSerial(Mid$(dtpInicio.Tag, 1, 4), Mid$(dtpInicio.Tag, 5, 2), Mid$(dtpInicio.Tag, 7, 2)) Then
            '...se sim, sugere a posição DISPONÍVEL
            lblSituacaoCaixa.Caption = fgDescricaoEstadoCaixa(enumEstadoCaixa.Disponivel)
        
        '...se não, verifica se é igual...
        ElseIf CStr(dtpInicio.Value) = fgDtXML_To_Date(Split(dtpInicio.Tag, ";")(0)) Then
            '...se sim, sugere a posição ATUAL do CAIXA
            lblSituacaoCaixa.Caption = fgDescricaoEstadoCaixa(Val(Split(dtpInicio.Tag, ";")(1)))
            
        '...se não, então a data informada é menor...
        Else
            '...neste caso, pesquisa a posição do caixa para o veículo legal e data selecionados
            Set objCaixaSubReserva = fgCriarObjetoMIU("A6MIU.clsCaixaSubReserva")
            Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
            
            If xmlRetorno.loadXML( _
                objCaixaSubReserva.ObterPosicaoCaixaSubReserva(Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(1), _
                                                               Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(0), _
                                                               strData, _
                                                               False, _
                                                               vntCodErro, _
                                                               vntMensagemErro)) Then
                If vntCodErro <> 0 Then
                    GoTo ErrorHandler
                End If

                intSituacaoCaixa = _
                    Val(xmlRetorno.documentElement.selectSingleNode("Grupo_PosicaoCaixaSubReserva/CO_SITU_CAIX_SUB_RESE_ATUAL").Text)
                
                lblSituacaoCaixa.Caption = fgDescricaoEstadoCaixa(intSituacaoCaixa)
            Else
                lblSituacaoCaixa.Caption = "Inexistente"
            End If
            
            Set objCaixaSubReserva = Nothing
            Set xmlRetorno = Nothing
        End If
    End If
    
    Call fgCursor(False)
    
    Exit Sub
    
ErrorHandler:
    Set objCaixaSubReserva = Nothing
    Set xmlRetorno = Nothing
    
    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmFiltro - dtpInicio_Change"

End Sub

Private Sub dtpInicio_KeyPress(KeyAscii As Integer)
    
    KeyAscii = 0

End Sub

Private Sub Form_Activate()

On Error GoTo ErrorHandler

    tlbData.Buttons("Comparacao").Image = 1
    blnFiltraHora = True

    If blnPrimeiroActivate Then
        fgCursor True
        Call flInit
        blnPrimeiroActivate = False
        fgCursor
    End If
    
    Call flAplicarSettingsRegistry
    
    Exit Sub

ErrorHandler:
    blnPrimeiroActivate = False
    mdiSBR.uctLogErros.MostrarErros Err, "frmFiltro - Form_Activate"
    
End Sub

' Carrega Registry do Windows referente ao A6.

Private Sub flCarregarRegistro()

Dim strRegistry                             As String

On Error GoTo ErrorHandler

    Set xmlDOMRegistro = CreateObject("MSXML2.DOMDocument.4.0")
    
    strRegistry = GetSetting("A6SBR", "Form Filtro\" & FormOwner.Name, "Settings")
    If strRegistry <> vbNullString Then
        If Not xmlDOMRegistro.loadXML(strRegistry) Then
            Call fgErroLoadXML(xmlDOMRegistro, App.EXEName, "frmFiltro", "flCarregarRegistro")
        Else
            If Not xmlDOMRegistro.selectSingleNode("//Grupo_BackOfficePerfilGeral") Is Nothing Then
                Call fgRemoveNode(xmlDOMRegistro, "Grupo_BackOfficePerfilGeral")
            End If
        End If
    End If

Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, "frmFiltro", "flCarregarRegistro", 0

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 13
            tlbComandos_ButtonClick tlbComandos.Buttons("aplicar")
        Case 27
            tlbComandos_ButtonClick tlbComandos.Buttons("cancelar")
    End Select

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCursor True
    flCarregarRegistro
    blnPrimeiroActivate = True

    'ReDim arrSgSistCoVeicLega(0)

    lblSituacaoCaixa.Caption = vbNullString

    Select Case TipoFiltro
    Case enumTipoFiltroA6.frmSubReservaResumo
        lcControlesFiltro.Add Me.cboGrupoVeicLegal
        lcControlesFiltro.Add Me.cboVeiculoLegal
        lcControlesFiltro.Add Me.cboTipoBackoffice
        strFuncionalidade = "frmFiltro_SubReservaResumo"
        
    Case enumTipoFiltroA6.frmControleRemessa
        lcControlesFiltro.Add Me.cboBancLiqu
        lcControlesFiltro.Add Me.cboSistema
        lcControlesFiltro.Add Me.tlbData
        lcControlesFiltro.Add Me.dtpInicio
        lcControlesFiltro.Add Me.dtpFim
        strFuncionalidade = "frmFiltro_ControleRemessa"
    
    Case enumTipoFiltroA6.frmSubReservaD0
        lcControlesFiltro.Add Me.cboBancLiqu
        lcControlesFiltro.Add Me.cboGrupoVeicLegal
        lcControlesFiltro.Add Me.cboVeiculoLegal
        lcControlesFiltro.Add Me.cboTipoBackoffice
        strFuncionalidade = "frmFiltro_SubReservaD0"
        
    Case enumTipoFiltroA6.frmRemessaRejeitada
        lcControlesFiltro.Add Me.cboBancLiqu
        lcControlesFiltro.Add Me.cboSistema
        lcControlesFiltro.Add Me.tlbData
        lcControlesFiltro.Add Me.dtpInicio
        lcControlesFiltro.Add Me.dtpFim

        strFuncionalidade = "frmFiltro_RemessaRejeitada"
        
    Case enumTipoFiltroA6.frmCaixaFuturo
        lcControlesFiltro.Add Me.cboBancLiqu
        lcControlesFiltro.Add Me.cboGrupoVeicLegal
        lcControlesFiltro.Add Me.cboVeiculoLegal
        lcControlesFiltro.Add Me.cboTipoBackoffice
        strFuncionalidade = "frmFiltro_CaixaFuturo"

    Case enumTipoFiltroA6.frmSubReservaAbertura
        lcControlesFiltro.Add Me.cboGrupoVeicLegal
        lcControlesFiltro.Add Me.cboVeiculoLegal
        strFuncionalidade = "frmFiltro_SubReservaAbertura"
        
    Case enumTipoFiltroA6.frmSubReservaFechamento
        lcControlesFiltro.Add Me.cboGrupoVeicLegal
        lcControlesFiltro.Add Me.cboVeiculoLegal
        strFuncionalidade = "frmFiltro_SubReservaFechamento"

    Case enumTipoFiltroA6.frmSubReservaConsultaAberturaFechamento
        lcControlesFiltro.Add Me.cboGrupoVeicLegal
        lcControlesFiltro.Add Me.cboVeiculoLegal
        lcControlesFiltro.Add Me.tlbData
        lcControlesFiltro.Add Me.dtpInicio
        lcControlesFiltro.Add Me.dtpFim
        lcControlesFiltro.Add Me.cboTipoBackoffice
        strFuncionalidade = "frmFiltro_SubReservaConsultaAberturaFechamento"

    End Select
    
    Set xmlPropriedades = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlPropriedades, "", "Grupo_Propriedades", "")
    Call fgAppendAttribute(xmlPropriedades, "Grupo_Propriedades", "Objeto", "")
    Call fgAppendAttribute(xmlPropriedades, "Grupo_Propriedades", "Operacao", "")
    
    Me.dtpInicio.Value = Date
    Me.dtpFim.Value = Date
    
    Call flConfiguraLayoutForm
    
    fgCursor
    
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmFiltro - Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlPropriedades = Nothing
    Set xmlDOMRegistro = Nothing
    Set xmlMapaNavegacao = Nothing

End Sub

Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    fgCursor True, False
    
    Select Case Button.Key
    Case "cancelar"
        Me.Hide
    Case "aplicar"
        If flValidarCampos Then
            Call flAplicarFiltro(True)
        End If
    End Select

    fgCursor

Exit Sub
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmFiltro - tlbComandos_ButtonClick"
End Sub

Private Sub tlbData_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    If Button.Image = 1 Then
        Button.Image = 0
        blnFiltraHora = False
    Else
        blnFiltraHora = True
        Button.Image = 1
    End If

Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - tlbData_ButtonClick"
End Sub

Private Sub tlbData_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
On Error GoTo ErrorHandler

    Select Case ButtonMenu.Text
        Case "Após"
            dtpFim.Visible = False
            ButtonMenu.Text = tlbData.Buttons(1).Caption
            tlbData.Buttons(1).Caption = "Após"
            
        Case "Antes"
            dtpFim.Visible = False
            ButtonMenu.Text = tlbData.Buttons(1).Caption
            tlbData.Buttons(1).Caption = "Antes"
        
        Case "Entre"
            dtpFim.Visible = True
            ButtonMenu.Text = tlbData.Buttons(1).Caption
            tlbData.Buttons(1).Caption = "Entre"
    
    End Select

Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - tlbData_ButtonMenuClick"

End Sub

' Carrega sistemas a partir da empresa selecionada.

Private Sub flCarregarSistema()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim xmlDomSistema       As MSXML2.DOMDocument40
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant
 
On Error GoTo ErrorHandler

    If cboSistema.Top > PROP_TOP_CAMPO_07 Then Exit Sub

    If cboBancLiqu.ListIndex > 0 Then
        
        Set xmlMapaNavegacao = Nothing
    
        Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    
        Call fgAppendNode(xmlMapaNavegacao, vbNullString, "Repeat_Leitura", vbNullString)
        Call fgAppendNode(xmlMapaNavegacao, "Repeat_Leitura", "Grupo_Leitura", vbNullString)
        Call fgAppendAttribute(xmlMapaNavegacao, "Grupo_Leitura", "Operacao", "LerTodos")
        Call fgAppendAttribute(xmlMapaNavegacao, "Grupo_Leitura", "Objeto", "A6A7A8.clsSistema")
        Call fgAppendNode(xmlMapaNavegacao, "Grupo_Leitura", "TP_VIGE", "S")
        Call fgAppendNode(xmlMapaNavegacao, "Grupo_Leitura", "TP_SEGR", "S")
        Call fgAppendNode(xmlMapaNavegacao, "Grupo_Leitura", "CO_EMPR", fgObterCodigoCombo(cboBancLiqu.Text))

        Set xmlDomSistema = CreateObject("MSXML2.DOMDocument.4.0")
        Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
        If Not xmlDomSistema.loadXML(objMIU.Executar(xmlMapaNavegacao.xml, vntCodErro, vntMensagemErro)) Then
        
            If vntCodErro <> 0 Then
                GoTo ErrorHandler
            End If

            cboSistema.Clear
            cboSistema.Enabled = False
            Exit Sub
        End If
        Set objMIU = Nothing
        
        Call fgCarregarCombos(cboSistema, xmlDomSistema, "Sistema", "SG_SIST", "NO_SIST", True)
        
        cboSistema.Enabled = True
        
    Else
        cboSistema.Clear
        cboSistema.Enabled = False
    End If

Exit Sub
ErrorHandler:

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarSistema", 0

End Sub

' Configura labels de data de acordo com o objeto chamador.

Private Sub flConfiguratlbData(ByVal pstrCaption As String, _
                               ByRef tlbControl As Object, _
                               ByRef dtpDataFim As Object)
        
On Error GoTo ErrorHandler

    With tlbControl
        
        Select Case pstrCaption
            
            Case "Após"
                dtpDataFim.Visible = False
                .Buttons(1).ButtonMenus(1).Text = "Antes"
                .Buttons(1).ButtonMenus(2).Text = "Entre"
                .Buttons(1).Caption = "Após"
                
            Case "Antes"
                dtpDataFim.Visible = False
                .Buttons(1).ButtonMenus(1).Text = "Após"
                .Buttons(1).ButtonMenus(2).Text = "Entre"
                .Buttons(1).Caption = "Antes"
            
            Case "Entre"
                dtpDataFim.Visible = True
                .Buttons(1).ButtonMenus(1).Text = "Após"
                .Buttons(1).ButtonMenus(2).Text = "Antes"
                .Buttons(1).Caption = "Entre"
                
        End Select
    
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flConfiguratlbData", 0

End Sub

' Busca código do veículo legal dentro do combo carregado.

Private Function flBuscaVeiculoLegal(ByVal pstrCodigo As String) As Boolean

Dim intCont                                 As Integer

On Error GoTo ErrorHandler

    If cboVeiculoLegal.ListCount <= 1 Then
        flBuscaVeiculoLegal = False
        Exit Function
    End If

    For intCont = 1 To UBound(arrSgSistCoVeicLega)
        If UCase$(Split(arrSgSistCoVeicLega(intCont), "k_")(1)) = UCase$(pstrCodigo) Then
            If cboVeiculoLegal.ListIndex <> intCont Then
                cboVeiculoLegal.ListIndex = intCont
            End If
            flBuscaVeiculoLegal = True
            Exit Function

        End If
    Next intCont
    
    flBuscaVeiculoLegal = False

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flBuscaVeiculoLegal", 0
End Function

Private Sub txtVeiculoLegal_Change()

On Error GoTo ErrorHandler

    If cboVeiculoLegal.ListCount > 0 Then
        If Not flBuscaVeiculoLegal(txtVeiculoLegal.Text) Then
            cboVeiculoLegal.ListIndex = 0
        End If
    Else
        Call cboVeiculoLegal_DropDown
    End If

Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - txtVeiculoLegal_Change"
End Sub

Private Sub txtVeiculoLegal_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub
