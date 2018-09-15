VERSION 5.00
Begin VB.UserControl ctlMenu 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1815
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   495
   ScaleWidth      =   1815
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu mnuSubReservaAbertura 
      Caption         =   "mnuSubReservaAbertura"
      Begin VB.Menu mnuMarcarTodas 
         Caption         =   "Marcar Todas"
      End
      Begin VB.Menu mnuDesmarcarTodas 
         Caption         =   "Desmarcar Todas"
      End
      Begin VB.Menu mnuSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuValorD1 
         Caption         =   "Valor em D-1"
      End
      Begin VB.Menu mnuValorInformado 
         Caption         =   "Valor Informado"
      End
      Begin VB.Menu mnuSaldoConta 
         Caption         =   "Saldo Conta"
      End
   End
   Begin VB.Menu mnuSubReservaFechamento 
      Caption         =   "mnuSubReservaFechamento"
      Begin VB.Menu mnuSubReservaFechamentoMarcarTodas 
         Caption         =   "Marcar Todas"
      End
      Begin VB.Menu mnuSubReservaFechamentoDermarcarTodas 
         Caption         =   "Desmarcar Todas"
      End
   End
End
Attribute VB_Name = "ctlMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ClickMenu(ByVal Retorno As Long)

Public Sub ShowMenuCaixaSubReservaAbertura(Optional ByVal pbAprovar As Boolean = False)
    PopupMenu mnuSubReservaAbertura
End Sub

Public Sub ShowMenuCaixaSubReservaFechamento(Optional ByVal pbAprovar As Boolean = False)
    PopupMenu mnuSubReservaFechamento
End Sub

Private Sub mnuMarcarTodas_Click()
    RaiseEvent ClickMenu(enumTipoSelecao.MarcarTodas)
End Sub

Private Sub mnuSaldoConta_Click()
    RaiseEvent ClickMenu(enumTipoAbertura.SaldoConta)
End Sub

Private Sub mnuDesmarcarTodas_Click()
    RaiseEvent ClickMenu(enumTipoSelecao.DesmarcarTodas)
End Sub

Private Sub mnuValorD1_Click()
    RaiseEvent ClickMenu(enumTipoAbertura.ValorD1)
End Sub

Private Sub mnuValorInformado_Click()
    RaiseEvent ClickMenu(enumTipoAbertura.ValorInformado)
End Sub

Private Sub mnuSubReservaFechamentoDermarcarTodas_Click()
    RaiseEvent ClickMenu(enumTipoSelecao.DesmarcarTodas)
End Sub

Private Sub mnuSubReservaFechamentoMarcarTodas_Click()
    RaiseEvent ClickMenu(enumTipoSelecao.MarcarTodas)
End Sub

