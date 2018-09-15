VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTesteConectividade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Teste de Conectividade"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.UpDown udTimer 
      Height          =   315
      Left            =   4740
      TabIndex        =   7
      Top             =   4440
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtTimer"
      BuddyDispid     =   196609
      OrigLeft        =   4860
      OrigTop         =   4470
      OrigRight       =   5100
      OrigBottom      =   4815
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
   End
   Begin VB.TextBox txtTimer 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   4350
      TabIndex        =   6
      Text            =   "10"
      Top             =   4440
      Width           =   390
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   60000
      Left            =   5250
      Top             =   4380
   End
   Begin VB.ComboBox cboEmissor 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4050
      Width           =   4860
   End
   Begin VB.ComboBox cboDestinatario 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4050
      Width           =   5115
   End
   Begin MSComctlLib.ListView lstConectividade 
      Height          =   3750
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   6615
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Emissor"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Destino"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data/Hora Envio"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Data/Hora R1"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Situação"
         Object.Width           =   4233
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbFuncoes 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   4860
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   635
      ButtonWidth     =   2937
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Enviar Mensagem"
            Key             =   "Enviar"
            ImageKey        =   "Testar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Atualizar            "
            Key             =   "Atualizar"
            ImageKey        =   "Atualizar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair                     "
            Key             =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   6000
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTesteConectividade.frx":0000
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTesteConectividade.frx":0112
            Key             =   "Testar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTesteConectividade.frx":0564
            Key             =   "Camara"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTesteConectividade.frx":09B6
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTesteConectividade.frx":0CD0
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTesteConectividade.frx":1022
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTesteConectividade.frx":1134
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTesteConectividade.frx":144E
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTesteConectividade.frx":1768
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTesteConectividade.frx":1A82
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      Caption         =   "Intervalo para Refresh automático da tela (em minutos) :"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   4500
      Width           =   3945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Emissor"
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
      TabIndex        =   4
      Top             =   3810
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Destinatário"
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
      Left            =   5160
      TabIndex        =   2
      Top             =   3810
      Width           =   1035
   End
End
Attribute VB_Name = "frmTesteConectividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Forulário responsável pelo envio e monitoramento das mensagens de teste de conectividade GEN0001

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Sub Form_Load()

Dim intIndexCombo                           As Integer
    
    On Error GoTo ErrorHandler
        
    Me.Icon = mdiBUS.Icon
    fgCenterMe Me
    
    Me.Show
    Call flInicializar
        
    DoEvents
    fgCursor True
    Call flCarregaLista
    fgCursor
    
    For intIndexCombo = 0 To cboEmissor.ListCount - 1
        If Left$(cboEmissor.List(intIndexCombo), 8) = "90400888" Then
            cboEmissor.ListIndex = intIndexCombo
        End If
    Next
    
    cboEmissor.Enabled = False
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmTesteConectividade - Form_Load"

End Sub

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A7Miu.clsMIU
#End If

Dim xmlNode             As MSXML2.IXMLDOMNode
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A7Miu.clsMIU")
    
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.BUS, "frmTesteConectividade", vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmTesteConectividade", "flInicializar")
    End If
    
    cboDestinatario.Clear
    cboDestinatario.AddItem "<Todos>"
    
    For Each xmlNode In xmlMapaNavegacao.selectNodes("//*[IN_TIPO_ISPB='2']")
         cboDestinatario.AddItem xmlNode.selectSingleNode("CO_ISPB").Text & " - " & xmlNode.selectSingleNode("NO_ISPB").Text
    Next
    
    cboEmissor.Clear
    cboEmissor.AddItem "<Todos>"
        
    For Each xmlNode In xmlMapaNavegacao.selectNodes("//*[CO_ISPB='61472676']")
         cboEmissor.AddItem xmlNode.selectSingleNode("CO_ISPB").Text & " - " & xmlNode.selectSingleNode("NO_ISPB").Text
    Next
        
    For Each xmlNode In xmlMapaNavegacao.selectNodes("//*[CO_ISPB='61411633']")
         cboEmissor.AddItem xmlNode.selectSingleNode("CO_ISPB").Text & " - " & xmlNode.selectSingleNode("NO_ISPB").Text
    Next
    
    For Each xmlNode In xmlMapaNavegacao.selectNodes("//*[CO_ISPB='33517640']")
         cboEmissor.AddItem xmlNode.selectSingleNode("CO_ISPB").Text & " - " & xmlNode.selectSingleNode("NO_ISPB").Text
    Next
    
    For Each xmlNode In xmlMapaNavegacao.selectNodes("//*[CO_ISPB='90400888']")
         cboEmissor.AddItem xmlNode.selectSingleNode("CO_ISPB").Text & " - " & xmlNode.selectSingleNode("NO_ISPB").Text
    Next
    
    Set objMIU = Nothing
    
    Exit Sub

ErrorHandler:
    
    Set xmlInit = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set xmlMapaNavegacao = Nothing

End Sub

Private Sub tlbFuncoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    
On Error GoTo ErrorHandler
    
    Select Case Button.Key
        Case "Enviar"
            
            If cboEmissor.ListIndex = -1 Then
                MsgBox "Selecione um emissor", vbInformation, Me.Caption
                Exit Sub
            End If
            
            If cboDestinatario.ListIndex = -1 Then
                MsgBox "Selecione um destinatário", vbInformation, Me.Caption
                Exit Sub
            End If
            
            fgCursor True
            Call flEnviarGEN0001
            fgCursor False
            
            MsgBox "Mensagem GEN0001 enviada."
            
            cboDestinatario.ListIndex = -1
            
        Case "Atualizar"
            
            fgCursor True
            Call flCarregaLista
            fgCursor False
            
        Case "Sair"
            
            Unload Me
    
    End Select
    
    Exit Sub

ErrorHandler:
    
    fgCursor False
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmTesteConectividade - tlbFuncoes_ButtonClick"

End Sub

'Processo de envio da mensagem GEN0001

Private Sub flEnviarGEN0001()

#If EnableSoap = 1 Then
    Dim objMonitoracao          As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracao          As A7Miu.clsMonitoracao
#End If

Dim xmlTesteConectividade       As MSXML2.DOMDocument40
Dim xmlAux                      As MSXML2.DOMDocument40
Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant

On Error GoTo ErrorHandler
    
    Set xmlTesteConectividade = CreateObject("MSXML2.DOMDocument.4.0")
    
    fgAppendNode xmlTesteConectividade, "", "TESTE_CONECTIVIDADE", ""
    
    If cboEmissor.Text = "<Todos>" Then
        
        If cboDestinatario.Text <> "<Todos>" Then
            
            Call fgAppendNode(xmlTesteConectividade, "TESTE_CONECTIVIDADE", "Grupo_TesteConectividade", "")
            Call fgAppendNode(xmlTesteConectividade, "Grupo_TesteConectividade", "CO_ISPB_EMIS", enumISPB.IspbSANTANDER)
            Call fgAppendNode(xmlTesteConectividade, "Grupo_TesteConectividade", "CO_EMPR", "523")
            Call fgAppendNode(xmlTesteConectividade, "Grupo_TesteConectividade", "CO_ISPB_DEST", fgObterCodigoCombo(cboDestinatario))
                
            Call fgAppendNode(xmlTesteConectividade, "TESTE_CONECTIVIDADE", "Grupo_TesteConectividade", "")
            Call fgAppendNode(xmlTesteConectividade, "Grupo_TesteConectividade", "CO_ISPB_EMIS", enumISPB.IspbBANESPA, "TESTE_CONECTIVIDADE")
            Call fgAppendNode(xmlTesteConectividade, "Grupo_TesteConectividade", "CO_EMPR", "701", "TESTE_CONECTIVIDADE")
            Call fgAppendNode(xmlTesteConectividade, "Grupo_TesteConectividade", "CO_ISPB_DEST", fgObterCodigoCombo(cboDestinatario), "TESTE_CONECTIVIDADE")
                
            Call fgAppendNode(xmlTesteConectividade, "TESTE_CONECTIVIDADE", "Grupo_TesteConectividade", "")
            Call fgAppendNode(xmlTesteConectividade, "Grupo_TesteConectividade", "CO_ISPB_EMIS", enumISPB.IspbBOZZANO, "TESTE_CONECTIVIDADE")
            Call fgAppendNode(xmlTesteConectividade, "Grupo_TesteConectividade", "CO_EMPR", "493", "TESTE_CONECTIVIDADE")
            Call fgAppendNode(xmlTesteConectividade, "Grupo_TesteConectividade", "CO_ISPB_DEST", fgObterCodigoCombo(cboDestinatario), "TESTE_CONECTIVIDADE")
                
            Call fgAppendNode(xmlTesteConectividade, "TESTE_CONECTIVIDADE", "Grupo_TesteConectividade", "")
            Call fgAppendNode(xmlTesteConectividade, "Grupo_TesteConectividade", "CO_ISPB_EMIS", enumISPB.IspbMERIDIONAL, "TESTE_CONECTIVIDADE")
            Call fgAppendNode(xmlTesteConectividade, "Grupo_TesteConectividade", "CO_EMPR", "558", "TESTE_CONECTIVIDADE")
            Call fgAppendNode(xmlTesteConectividade, "Grupo_TesteConectividade", "CO_ISPB_DEST", fgObterCodigoCombo(cboDestinatario), "TESTE_CONECTIVIDADE")
                
            fgAppendNode xmlTesteConectividade, "Grupo_TesteConectividade", "CO_ISPB_DEST", fgObterCodigoCombo(cboDestinatario)
        
        Else
            
            If cboDestinatario.Text = "<Todos>" Then
            
                For Each xmlNode In xmlMapaNavegacao.selectNodes("//*[IN_TIPO_ISPB='2']")
                    Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")
                    fgAppendNode xmlAux, "", "Grupo_TesteConectividade", ""
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_ISPB_EMIS", enumISPB.IspbSANTANDER
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_EMPR", "523"
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_ISPB_DEST", xmlNode.selectSingleNode("CO_ISPB").Text
                    fgAppendXML xmlTesteConectividade, "TESTE_CONECTIVIDADE", xmlAux.xml
                    Set xmlAux = Nothing
                Next
            
                For Each xmlNode In xmlMapaNavegacao.selectNodes("//*[IN_TIPO_ISPB='2']")
                    Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")
                    fgAppendNode xmlAux, "", "Grupo_TesteConectividade", ""
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_ISPB_EMIS", enumISPB.IspbBANESPA
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_EMPR", "701"
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_ISPB_DEST", xmlNode.selectSingleNode("CO_ISPB").Text
                    fgAppendXML xmlTesteConectividade, "TESTE_CONECTIVIDADE", xmlAux.xml
                    Set xmlAux = Nothing
                Next
            
                For Each xmlNode In xmlMapaNavegacao.selectNodes("//*[IN_TIPO_ISPB='2']")
                    Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")
                    fgAppendNode xmlAux, "", "Grupo_TesteConectividade", ""
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_ISPB_EMIS", enumISPB.IspbBOZZANO
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_EMPR", "493"
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_ISPB_DEST", xmlNode.selectSingleNode("CO_ISPB").Text
                    fgAppendXML xmlTesteConectividade, "TESTE_CONECTIVIDADE", xmlAux.xml
                    Set xmlAux = Nothing
                Next
            
                For Each xmlNode In xmlMapaNavegacao.selectNodes("//*[IN_TIPO_ISPB='2']")
                    Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")
                    fgAppendNode xmlAux, "", "Grupo_TesteConectividade", ""
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_ISPB_EMIS", enumISPB.IspbMERIDIONAL
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_EMPR", "558"
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_ISPB_DEST", xmlNode.selectSingleNode("CO_ISPB").Text
                    fgAppendXML xmlTesteConectividade, "TESTE_CONECTIVIDADE", xmlAux.xml
                    Set xmlAux = Nothing
                Next
            
            End If
        End If
    Else
        If cboDestinatario.Text <> "<Todos>" Then
            fgAppendNode xmlTesteConectividade, "TESTE_CONECTIVIDADE", "Grupo_TesteConectividade", ""
            fgAppendNode xmlTesteConectividade, "Grupo_TesteConectividade", "CO_ISPB_EMIS", fgObterCodigoCombo(cboEmissor)
            
            If fgObterCodigoCombo(cboEmissor) = enumISPB.IspbSANTANDER Then
                fgAppendNode xmlTesteConectividade, "Grupo_TesteConectividade", "CO_EMPR", "523"
            ElseIf fgObterCodigoCombo(cboEmissor) = enumISPB.IspbBANESPA Then
                fgAppendNode xmlTesteConectividade, "Grupo_TesteConectividade", "CO_EMPR", "701"
            ElseIf fgObterCodigoCombo(cboEmissor) = enumISPB.IspbBOZZANO Then
                fgAppendNode xmlTesteConectividade, "Grupo_TesteConectividade", "CO_EMPR", "493"
            ElseIf fgObterCodigoCombo(cboEmissor) = enumISPB.IspbMERIDIONAL Then
                fgAppendNode xmlTesteConectividade, "Grupo_TesteConectividade", "CO_EMPR", "558"
            End If
            
            fgAppendNode xmlTesteConectividade, "Grupo_TesteConectividade", "CO_ISPB_DEST", fgObterCodigoCombo(cboDestinatario)
        Else
            
            For Each xmlNode In xmlMapaNavegacao.selectNodes("//*[IN_TIPO_ISPB='2']")
                
                Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")
                
                fgAppendNode xmlAux, "", "Grupo_TesteConectividade", ""
                fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_ISPB_EMIS", fgObterCodigoCombo(cboEmissor)
                
                If fgObterCodigoCombo(cboEmissor) = enumISPB.IspbSANTANDER Then
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_EMPR", "523"
                ElseIf fgObterCodigoCombo(cboEmissor) = enumISPB.IspbBANESPA Then
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_EMPR", "701"
                ElseIf fgObterCodigoCombo(cboEmissor) = enumISPB.IspbBOZZANO Then
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_EMPR", "493"
                ElseIf fgObterCodigoCombo(cboEmissor) = enumISPB.IspbMERIDIONAL Then
                    fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_EMPR", "558"
                End If
        
                fgAppendNode xmlAux, "Grupo_TesteConectividade", "CO_ISPB_DEST", xmlNode.selectSingleNode("CO_ISPB").Text
                
                fgAppendXML xmlTesteConectividade, "TESTE_CONECTIVIDADE", xmlAux.xml
                
                Set xmlAux = Nothing
            Next
            
        End If
    End If
    
    Set objMonitoracao = fgCriarObjetoMIU("A7Miu.clsMonitoracao")
    Call objMonitoracao.TesteConectividade(xmlTesteConectividade.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMonitoracao = Nothing
    
    Set xmlTesteConectividade = Nothing
    Exit Sub

ErrorHandler:
    
    Set xmlTesteConectividade = Nothing
    Set objMonitoracao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flEnviarGEN0001", 0

End Sub

Private Sub tlbFuncoes_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
    tlbFuncoes.Buttons(1).ToolTipText = ButtonMenu.Text
    tlbFuncoes.Buttons(1).Tag = ButtonMenu.Key

End Sub

Private Sub flCarregaLista()

#If EnableSoap = 1 Then
    Dim objMonitoracao      As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracao      As A7Miu.clsMonitoracao
#End If

Dim xmlNode                 As MSXML2.IXMLDOMNode
Dim xmlHistGEN0001          As MSXML2.DOMDocument40
Dim strHistGEN0001          As String
Dim objListItem             As MSComctlLib.ListItem
Dim strISPBDestinatario     As String
Dim strISPBEmissor          As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    Set objMonitoracao = fgCriarObjetoMIU("A7Miu.clsMonitoracao")
    
    strHistGEN0001 = objMonitoracao.ObterHistoricoGEN0001(vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    lstConectividade.ListItems.Clear
            
    If strHistGEN0001 = vbNullString Then
        Set objMonitoracao = Nothing
        Exit Sub
    End If
                
    Set xmlHistGEN0001 = CreateObject("MSXML2.DOMDocument.4.0")
                
    xmlHistGEN0001.loadXML strHistGEN0001
            
    For Each xmlNode In xmlHistGEN0001.selectNodes("//REPET_GEN0001/*")
        
        strISPBDestinatario = xmlNode.selectSingleNode("CO_ISPB_DEST").Text
        strISPBEmissor = xmlNode.selectSingleNode("CO_ISPB_EMIS").Text
        
        If Not xmlMapaNavegacao.selectSingleNode("//Grupo_Instituicao_ISPB[CO_ISPB='" & strISPBDestinatario & "']/NO_ISPB") Is Nothing Then
            strISPBDestinatario = strISPBDestinatario & " - " & xmlMapaNavegacao.selectSingleNode("//Grupo_Instituicao_ISPB[CO_ISPB='" & strISPBDestinatario & "']/NO_ISPB").Text
        End If
        
        If Not xmlMapaNavegacao.selectSingleNode("//Grupo_Instituicao_ISPB[CO_ISPB='" & strISPBEmissor & "']/NO_ISPB") Is Nothing Then
            strISPBEmissor = strISPBEmissor & " - " & xmlMapaNavegacao.selectSingleNode("//Grupo_Instituicao_ISPB[CO_ISPB='" & strISPBEmissor & "']/NO_ISPB").Text
        End If
        
        Set objListItem = lstConectividade.ListItems.Add(, "", strISPBEmissor)
        objListItem.SubItems(1) = strISPBDestinatario
        objListItem.SubItems(2) = fgDtHrXML_To_Interface(xmlNode.selectSingleNode("DH_MESG").Text)
        
        If Not xmlNode.selectSingleNode("DH_MESG_R1") Is Nothing Then
            If xmlNode.selectSingleNode("DH_MESG_R1").Text <> "" Then
                objListItem.SubItems(3) = fgDtHrXML_To_Interface(xmlNode.selectSingleNode("DH_MESG_R1").Text)
            End If
        End If
        
        objListItem.SubItems(4) = xmlNode.selectSingleNode("DE_STAT_MESG").Text
        
        Select Case xmlNode.selectSingleNode("COR").Text
            
            Case "RED"
                objListItem.ForeColor = vbRed
                objListItem.Bold = True
                
                objListItem.ListSubItems.Item(1).ForeColor = vbRed
                objListItem.ListSubItems.Item(2).ForeColor = vbRed
                objListItem.ListSubItems.Item(3).ForeColor = vbRed
                objListItem.ListSubItems.Item(4).ForeColor = vbRed
                    
                objListItem.ListSubItems.Item(1).Bold = True
                objListItem.ListSubItems.Item(2).Bold = True
                objListItem.ListSubItems.Item(3).Bold = True
                objListItem.ListSubItems.Item(4).Bold = True
                    
            Case "BLUE"
                objListItem.ForeColor = vbBlue
                objListItem.ListSubItems.Item(1).ForeColor = vbBlue
                objListItem.ListSubItems.Item(2).ForeColor = vbBlue
                objListItem.ListSubItems.Item(3).ForeColor = vbBlue
                objListItem.ListSubItems.Item(4).ForeColor = vbBlue
            
        End Select
            
    Next
        
    Set objMonitoracao = Nothing
    
    Exit Sub

ErrorHandler:
    
    Set objMonitoracao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaLista", 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        
        fgCursor True
        flCarregaLista
        fgCursor
    End If
    
    Exit Sub
ErrorHandler:
    
    fgCursor
    
    mdiBUS.uctLogErros.MostrarErros Err, ("frmTesteConectividade - Form_KeyDown")

End Sub

Private Sub tmrRefresh_Timer()

On Error GoTo ErrorHandler

    If Not IsNumeric(txtTimer.Text) Then Exit Sub
    
    If CLng(txtTimer.Text) = 0 Then Exit Sub
    
    If fgVerificaJanelaVerificacao() Then Exit Sub
    
    fgCursor True

    intContMinutos = intContMinutos + 1
    
    If intContMinutos >= txtTimer.Text Then
        Call flCarregaLista
        intContMinutos = 0
    End If

    fgCursor False

    Exit Sub
ErrorHandler:
    
    fgCursor False
    
    mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - tmrRefresh_Timer"

End Sub
