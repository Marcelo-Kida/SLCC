VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBuscaNo 
   BorderStyle     =   0  'None
   Caption         =   "frmBuscaNo"
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBusca 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   300
      Width           =   3435
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaNo.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaNo.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaNo.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaNo.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaNo.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaNo.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaNo.frx":150C
            Key             =   "ItemElementar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaNo.frx":195E
            Key             =   "OpenReservaFuturo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaNo.frx":1A70
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaNo.frx":1B6A
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaNo.frx":1C64
            Key             =   "Leaf"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaNo.frx":1D5E
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaNo.frx":21B0
            Key             =   "Proximo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
      ButtonWidth     =   1931
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Anterior"
            Key             =   "Anterior"
            ImageKey        =   "Anterior"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Próximo"
            Key             =   "Proximo"
            ImageKey        =   "Proximo"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCriterio 
      Caption         =   "lblCriterio"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1335
   End
End
Attribute VB_Name = "frmBuscaNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário a navegação entre registros cadastrados,
' de determinada tabela.

Option Explicit

Public objTreeView                          As MSComctlLib.TreeView

Public Event BuscaEfetuada(ByRef objNode As MSComctlLib.Node)

Public Property Get Criterio() As String
    Criterio = lblCriterio.Caption
End Property

Public Property Let Criterio(ByVal NewValue As String)
    lblCriterio.Caption = NewValue
End Property

' Passa ao registro anterior.

Private Function flEncontraNoAnt(ByRef objNode As MSComctlLib.Node, _
                                  ByRef pstrBusca As String, _
                         Optional ByRef objNodeIgnorado As MSComctlLib.Node = Nothing, _
                         Optional ByVal pblnIgnoraChild As Boolean = False) As MSComctlLib.Node

Dim objNodeAux                              As MSComctlLib.Node

On Error GoTo ErrorHandler

    If objNode Is Nothing Then
        Set flEncontraNoAnt = Nothing
        Exit Function
    End If
    
    If Not objNodeIgnorado Is objNode Then
    
        If Not pblnIgnoraChild Then
            Set objNodeAux = flEncontraNoAnt(flEncontraUltimoNo(objNode.Child), pstrBusca, objNodeIgnorado)
            If Not objNodeAux Is Nothing Then
                Set flEncontraNoAnt = objNodeAux
                Exit Function
            End If
        End If

        If UCase$(pstrBusca) = UCase$(Left$(objNode.Text, Len(pstrBusca))) Then
            Set flEncontraNoAnt = objNode
            Exit Function
        End If
    End If
    
    Set objNodeAux = flEncontraNoAnt(objNode.Previous, pstrBusca, objNodeIgnorado)
    If Not objNodeAux Is Nothing Then
        Set flEncontraNoAnt = objNodeAux
        Exit Function
    End If
    
    Set flEncontraNoAnt = flEncontraNoAnt(objNode.Parent, pstrBusca, , True)
    
Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flEncontraNoAnt", 0
End Function

' Passa ao registro posterior.

Private Function flEncontraNoProx(ByRef objNode As MSComctlLib.Node, _
                                  ByVal pstrBusca As String, _
                         Optional ByRef objNodeIgnorado As MSComctlLib.Node = Nothing, _
                         Optional ByVal blnIgnoraChild As Boolean = False) As MSComctlLib.Node

Dim objNodeAux                              As MSComctlLib.Node

On Error GoTo ErrorHandler

    If objNode Is Nothing Then
        Set flEncontraNoProx = Nothing
        Exit Function
    End If
    
    If Not objNodeIgnorado Is objNode Then
        If UCase$(pstrBusca) = UCase$(Left$(objNode.Text, Len(pstrBusca))) Then
            Set flEncontraNoProx = objNode
            Exit Function
        End If
    End If
    
    If Not blnIgnoraChild Then
        Set objNodeAux = flEncontraNoProx(objNode.Child, pstrBusca, objNodeIgnorado)
        If Not objNodeAux Is Nothing Then
            Set flEncontraNoProx = objNodeAux
            Exit Function
        End If
    End If
    
    Set objNodeAux = flEncontraNoProx(objNode.Next, pstrBusca, objNodeIgnorado)
    If Not objNodeAux Is Nothing Then
        Set flEncontraNoProx = objNodeAux
        Exit Function
    End If
    
    Set flEncontraNoProx = flEncontraNoProx(objNode.Parent, pstrBusca, objNode.Parent, True)
    
Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flEncontraNoProx", 0
End Function

Private Sub Form_Activate()
On Error GoTo ErrorHandler

    flAlwaysOnTop Me, True
    txtBusca.SetFocus

Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - Form_Activate"
End Sub

' Passa ao registro anterior.

Public Function ProcuraNoAnt(ByVal pstrTexto As String) As MSComctlLib.Node

Dim objNode                                 As MSComctlLib.Node

On Error GoTo ErrorHandler

    If objTreeView.SelectedItem Is Nothing And objTreeView.Nodes.Count > 0 Then
        Set objNode = flEncontraNoAnt(flEncontraUltimoNo(flEncontraPrimeiroNo(objTreeView.Nodes(1))), pstrTexto)
    Else
        Set objNode = flEncontraNoAnt(objTreeView.SelectedItem, pstrTexto, objTreeView.SelectedItem)
        If objNode Is Nothing And objTreeView.Nodes.Count > 0 Then
            Set objNode = flEncontraNoAnt(flEncontraUltimoNo(flEncontraPrimeiroNo(objTreeView.Nodes(1))), pstrTexto)
        End If
    End If
    Set ProcuraNoAnt = objNode

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "ProcuraNoAnt", 0
End Function

' Passa ao primeiro registro.

Private Function flEncontraPrimeiroNo(ByRef objNode As MSComctlLib.Node) As MSComctlLib.Node

Dim objNodeAux                              As MSComctlLib.Node

On Error GoTo ErrorHandler

    If objNode Is Nothing Then
        Set flEncontraPrimeiroNo = Nothing
        Exit Function
    End If
    
    Set objNodeAux = flEncontraPrimeiroNo(objNode.Previous)
    If Not objNodeAux Is Nothing Then
        Set flEncontraPrimeiroNo = objNodeAux
        Exit Function
    End If

    Set objNodeAux = flEncontraPrimeiroNo(objNode.Parent)
    If Not objNodeAux Is Nothing Then
        Set flEncontraPrimeiroNo = objNodeAux
        Exit Function
    End If
    
    Set flEncontraPrimeiroNo = objNode

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flEncontraPrimeiroNo", 0
End Function

' Passa ao último registro.

Private Function flEncontraUltimoNo(ByRef objNode As MSComctlLib.Node) As MSComctlLib.Node

Dim objNodeAux                              As MSComctlLib.Node

On Error GoTo ErrorHandler

    If objNode Is Nothing Then
        Set flEncontraUltimoNo = Nothing
        Exit Function
    End If
    
    Set objNodeAux = flEncontraUltimoNo(objNode.Next)
    If Not objNodeAux Is Nothing Then
        Set flEncontraUltimoNo = objNodeAux
        Exit Function
    End If
    
    Set objNodeAux = flEncontraUltimoNo(objNode.Child)
    If Not objNodeAux Is Nothing Then
        Set flEncontraUltimoNo = objNodeAux
        Exit Function
    End If
    
    Set flEncontraUltimoNo = objNode

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flEncontraUltimoNo", 0
End Function

' Passa ao registro posterior.

Public Function ProcuraNoProx(ByVal pstrTexto As String) As MSComctlLib.Node

Dim objNode                                 As MSComctlLib.Node

On Error GoTo ErrorHandler

    If objTreeView.SelectedItem Is Nothing And objTreeView.Nodes.Count > 0 Then
        Set objNode = flEncontraNoProx(flEncontraPrimeiroNo(objTreeView.Nodes(1)), pstrTexto)
    Else
        Set objNode = flEncontraNoProx(objTreeView.SelectedItem, pstrTexto, objTreeView.SelectedItem)
        If objNode Is Nothing And objTreeView.Nodes.Count > 0 Then
            Set objNode = flEncontraNoProx(flEncontraPrimeiroNo(objTreeView.Nodes(1)), pstrTexto)
        End If
    End If
    Set ProcuraNoProx = objNode

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "ProcuraNoProx", 0
End Function

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim objNode                                 As MSComctlLib.Node

On Error GoTo ErrorHandler

    Select Case Button.Key
        Case "Proximo"
            Set objNode = ProcuraNoProx(RTrim$(txtBusca.Text))
            If Not objNode Is Nothing Then
                objNode.EnsureVisible
                objNode.Selected = True
            End If
            RaiseEvent BuscaEfetuada(objNode)
        Case "Anterior"
            Set objNode = ProcuraNoAnt(RTrim$(txtBusca.Text))
            If Not objNode Is Nothing Then
                objNode.EnsureVisible
                objNode.Selected = True
            End If
            RaiseEvent BuscaEfetuada(objNode)
        Case "Sair"
            Me.Hide
    End Select

Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - tlbCadastro_ButtonClick"
End Sub
