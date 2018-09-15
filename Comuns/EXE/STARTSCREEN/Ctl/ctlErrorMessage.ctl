VERSION 5.00
Begin VB.UserControl ctlErrorMessage 
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   555
   ScaleWidth      =   600
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "ctlErrorMessage.ctx":0000
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "ctlErrorMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ErroGerado(ByRef ErrNumber As Long, ByRef ErrDescription As String, ByRef Cancel As Boolean)

Public Function MostrarErros(ByVal objErro As ErrObject)
Dim llErrNumber                             As Long
Dim lsErrDescription                        As String
Dim lbCancel                                As Boolean
    
    llErrNumber = objErro.Number - (vbObjectError + 513)
    lsErrDescription = objErro.Description
    
    RaiseEvent ErroGerado(llErrNumber, lsErrDescription, lbCancel)
        
    If Not lbCancel Then
        frmErros.ErrorNumber = llErrNumber
        frmErros.ErrorDescription = lsErrDescription
        frmErros.Show vbModal
    End If
    
End Function
