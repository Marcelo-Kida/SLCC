VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmResultOperacaoLote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resultado da Operação em Lote"
   ClientHeight    =   6090
   ClientLeft      =   3420
   ClientTop       =   1935
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9120
   Begin SHDocVwCtl.WebBrowser wbResultado 
      Height          =   5475
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      ExtentX         =   15690
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
   Begin MSComctlLib.Toolbar tblBotoes 
      Height          =   330
      Left            =   8280
      TabIndex        =   1
      Top             =   5760
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
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   0
      Top             =   5520
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
            Picture         =   "frmResultOperacaoLote.frx":0000
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmResultOperacaoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário visualização do resultado de um processamento em lote.

'Estrutura do XML:  <Repeat_ControleErro>
'                       <Grupo_ControleErro> (1)
'                           <Operacao>...</Operacao><Status>...</Status>
'                       </Grupo_ControleErro>
'                       <Grupo_ControleErro> (N)
'                           <Operacao>...</Operacao><Status>...</Status>
'                       </Grupo_ControleErro>
'                   </Repeat_ControleErro>"
'
'                       ---    OU    ---
'
'                   <Repeat_Info>
'                       <Grupo_Info>        (1)
'                           <Mensagem>...</Mensagem>                <-- TAG não obrigatória
'                           <NumeroComando>...</NumeroComando>
'                       </Grupo_Info>
'                       <Grupo_Info>        (N)
'                           <Mensagem>...</Mensagem>                <-- TAG não obrigatória
'                           <NumeroComando>...</NumeroComando>
'                       </Grupo_Info>
'                   </Repeat_Info>

Option Explicit

Private blnApresentaInfo                    As Boolean
Private strResultado                        As String
Private strHTML                             As String
Public strDescricaoOperacao                 As String

Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    Me.Icon = mdiSBR.Icon
    
    fgCenterMe Me
    
    fgCursor True
    Call flAtualizaConteudoBrowser
    DoEvents
    
    Call flCarregarBrowser(strResultado)
    fgCursor
    
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "Form_Load", Me.Caption
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmResultOperacaoLote = Nothing
    
End Sub

Private Sub tblBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case gstrSair
            Unload Me
    
    End Select
    
End Sub

Public Property Let Resultado(ByVal pstrResultado As String)
    strResultado = pstrResultado
End Property

' Carrega browser para a apresentação do resultado do processamento em lote na tela.

Private Sub flCarregarBrowser(ByVal pstrResultado As String)

Dim xmlResultado                            As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim intContOK                               As Integer
Dim intContErroNegocioEspecifico            As Integer
Dim intContErroDetalhar                     As Integer
Dim strErroDetalhar                         As String
Dim blnCorVermelha                          As Boolean

On Error GoTo ErrorHandler

    Set xmlResultado = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlResultado.loadXML(pstrResultado)
    
    strHTML = "<html><body><center><table border=0>"
    
    If Not blnApresentaInfo Then
        'Captura o total de ITENS do lote enviado
        strHTML = strHTML & _
                "<tr BGColor=""#BBBBBB"">" & _
                "<td width=""80%""><font Color=White Size=3 Face=Verdana>" & "Total de itens processados no lote</font></td>" & _
                "<td><center><font Color=White Size=3 Face=Verdana>" & xmlResultado.documentElement.childNodes.length & "</font></center></td>" & _
                "</tr>"
        
        'Verifica o conteúdo do XML
        For Each objDomNode In xmlResultado.documentElement.childNodes
            'Operação OK
            If UCase(objDomNode.selectSingleNode("CodigoErro").Text) = 0 Then
                intContOK = intContOK + 1
                
            'Erros, neste caso, detalhar para o usuário
            Else
                'Verifica se o erro é crítico
                blnCorVermelha = objDomNode.selectSingleNode("TipoErro").Text = 2
                
                strErroDetalhar = strErroDetalhar & _
                        "<tr BGColor=""#CEDAEA"">" & _
                        "<td><font Color=Black Size=2 Face=Verdana>" & objDomNode.selectSingleNode("Operacao").Text & "</font></td>" & _
                        "<td><font Color=" & IIf(blnCorVermelha, "Red", "Black") & " Size=2 Face=Verdana>" & objDomNode.selectSingleNode("Status").Text & "</font></center></td>" & _
                        "</tr>"
                
                intContErroDetalhar = intContErroDetalhar + 1
          
            End If
        Next
        
        'Captura o total de ITENS confirmadOs OK
        If intContOK > 0 Then
            strHTML = strHTML & _
                    "<tr BGColor=""#BBBBBB"">" & _
                    "<td width=""80%""><font Color=White Size=3 Face=Verdana>Itens " & strDescricaoOperacao & " </font></td>" & _
                    "<td><center><font Color=White Size=3 Face=Verdana>" & intContOK & "</font></center></td>" & _
                    "</tr>"
        End If
        
        'Captura o total de ITENS com erro de negócio ESPECÍFICO
        If intContErroNegocioEspecifico > 0 Then
            strHTML = strHTML & _
                    "<tr BGColor=""#BBBBBB"">" & _
                    "<td width=""80%""><font Color=White Size=3 Face=Verdana>Itens processados por outro usuário</font></td>" & _
                    "<td><center><font Color=White Size=3 Face=Verdana>" & intContErroNegocioEspecifico & "</font></center></td>" & _
                    "</tr>"
        End If
        
        'Se o XML possuir erros apresenta-os detalhadamente
        '   Operação | Descrição do Erro
        If strErroDetalhar <> vbNullString Then
            strHTML = strHTML & _
                    "<tr></tr><tr></tr>"
            
            strHTML = strHTML & _
                    "<tr BGColor=""#BBBBBB"">" & _
                    "<td width=""80%""><font Color=White Size=3 Face=Verdana>Itens com erro</font></td>" & _
                    "<td><center><font Color=White Size=3 Face=Verdana>" & intContErroDetalhar & "</font></center></td>" & _
                    "</tr>"
            
            strHTML = strHTML & _
                    "</table><br><table border=0>"
                    
            strHTML = strHTML & _
                    "<tr BGColor=""#BBBBBB"">" & _
                    "<td width=""30%""><font Color=White Size=3 Face=Verdana>Veículo Legal</font></td>" & _
                    "<td width=""70%""><font Color=White Size=3 Face=Verdana>Descrição do Erro</font></center></td>" & _
                    "</tr>"
                    
            strHTML = strHTML & _
                    strErroDetalhar
        End If
    Else
        strHTML = strHTML & _
                "<tr BGColor=""#BBBBBB"">" & _
                "<td width=""80%""><font Color=White Size=3 Face=Verdana>" & "Mensagens para <b>Pagamento de Despesas</b> geradas através da Entrada Manual ou Outro Processo</font></td>" & _
                "<td><center><font Color=White Size=3 Face=Verdana>" & xmlResultado.documentElement.childNodes.length - xmlResultado.documentElement.selectNodes("//Mensagem").length & "</font></center></td>" & _
                "</tr>"
    
        If xmlResultado.documentElement.selectNodes("//Mensagem").length > 0 Then
            strHTML = strHTML & _
                    "<tr BGColor=""#BBBBBB"">" & _
                    "<td width=""80%""><font Color=White Size=3 Face=Verdana>" & "Mensagens para <b>Pagamento de Despesas</b> geradas e disponíveis para a <b>Liberação</b></font></td>" & _
                    "<td><center><font Color=White Size=3 Face=Verdana>" & xmlResultado.documentElement.selectNodes("//Mensagem").length & "</font></center></td>" & _
                    "</tr>"
            
            strHTML = strHTML & _
                    "</table><br><table border=0>"
            
            strHTML = strHTML & _
                    "<tr BGColor=""#BBBBBB"">" & _
                    "<td width=""50%""><font Color=White Size=3 Face=Verdana>Mensagem</font></td>" & _
                    "<td width=""50%""><center><font Color=White Size=3 Face=Verdana>Número Comando</font></center></td>" & _
                    "</tr>"
            
            For Each objDomNode In xmlResultado.documentElement.childNodes
                If Not objDomNode.selectSingleNode("Mensagem") Is Nothing Then
                    strHTML = strHTML & _
                            "<tr BGColor=""#CEDAEA"">" & _
                            "<td><font Color=Black Size=2 Face=Verdana>" & objDomNode.selectSingleNode("Mensagem").Text & "</font></td>" & _
                            "<td><font Color=Black Size=2 Face=Verdana>" & objDomNode.selectSingleNode("NumeroComando").Text & "</font></center></td>" & _
                            "</tr>"
                End If
            Next
        End If
    End If
    
    strHTML = strHTML & _
            "</table></center></body></html>"
    
    Set xmlResultado = Nothing
    
    Call flAtualizaConteudoBrowser(strHTML)

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flCarregarBrowser", 0

End Sub

Private Sub wbResultado_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    pDisp.Document.Body.innerHTML = strHTML
End Sub

' Refresh no conteúdo do browser.

Private Sub flAtualizaConteudoBrowser(Optional pstrConteudo As String = vbNullString)
    strHTML = pstrConteudo
    wbResultado.Navigate "about:blank"
End Sub

Public Property Let ApresentaInfo(ByVal pblnApresentaInfo As Boolean)
    blnApresentaInfo = pblnApresentaInfo
End Property
