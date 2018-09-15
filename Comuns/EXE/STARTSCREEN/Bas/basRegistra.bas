Attribute VB_Name = "basRegistra"
Option Explicit

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

Public Sub fgRegistraComponentes()

Dim strCLIREG32                              As String
Dim strArquivo                                As String

On Error GoTo ErrorHandler

    strCLIREG32 = App.Path & "\CliReg32.Exe"
    
    If gblnRegistraTLB Then
            
        strArquivo = App.Path & "\A6A7A8Miu"
        Shell strCLIREG32 & " """ & strArquivo & ".VBR"" -t """ & strArquivo & ".TLB"" -d -nologo -q -s " & gstrSource & " -l"
    
    End If
    
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    
    Call fgRaiseError("A6A7A8StartScreen", "basRegistra", "fgRegistraComponentes Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

Public Sub fgDesregistraComponentes()

Dim strCLIREG32                              As String
Dim strArquivo                               As String

On Error GoTo ErrorHandler
    
    strCLIREG32 = App.Path & "\CliReg32.Exe"
    
    'Caso tenha registrado os componentes, fuma tudo
    If gblnRegistraTLB Then
        
        strArquivo = App.Path & "\A6A7A8Miu"
        Shell strCLIREG32 & " """ & strArquivo & ".VBR"" -t """ & strArquivo & ".TLB"" -u -d -nologo -q -l"
    
    End If
    
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "modRegistra", "fgDesregistraComponentes Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
        
End Sub
