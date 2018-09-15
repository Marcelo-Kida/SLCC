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
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EFB2E7B0311"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"User Control"
Option Explicit

Private blnFormLoaded                       As Boolean

Public Event ErroGerado(ByRef ErrNumber As Long, ByRef ErrDescription As String, ByRef Cancel As Boolean)

Public Function MostrarErros(ByVal objErro As ErrObject, _
                             ByVal pstrOrigem As String, _
                    Optional ByVal pstrTitulo As String = "Informações sobre o Erro") As Variant

Dim lngErrNumber                             As Long
Dim strErrDescription                        As String
Dim strErrSource                             As String
Dim blnCancel                                As Boolean
    
    fgCursor False
    
    strErrDescription = objErro.Description
    lngErrNumber = fgObterCodigoDeErroDeNegocioXMLErro(Err.Description)
    strErrSource = Err.Source
    
    RaiseEvent ErroGerado(lngErrNumber, strErrDescription, blnCancel)
    
    If Not blnCancel And Not blnFormLoaded Then
        blnFormLoaded = True
        frmErros.ErrorNumber = lngErrNumber
        frmErros.ErrorDescription = strErrDescription
        frmErros.ErrorSouce = App.EXEName & " - " & pstrOrigem
        frmErros.Caption = pstrTitulo
        frmErros.Show vbModal
    End If
    
    blnFormLoaded = False
    
End Function

Private Sub UserControl_Initialize()
    blnFormLoaded = False
End Sub
