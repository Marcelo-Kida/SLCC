VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSharedPropMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Objeto responsável pela leitura e gravação de dados no Shared Property Memory do COM+

Option Explicit

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

'--------------------------------------------------------------------------------
'Ler os dados gravados no Shared Property Memory do COM+
'--------------------------------------------------------------------------------

Public Function GetSPMProperty(ByVal pstrSPMGroupName As String, _
                               ByVal pstrSPMPropertyName As String) As Variant

Dim objSpmManager                           As SharedPropertyGroupManager
Dim objSpmGroup                             As SharedPropertyGroup
Dim objSpmProperty                          As SharedProperty
Dim blnGroupExists                          As Boolean
Dim blnPropertyExists                       As Boolean

On Error GoTo ErrorHandler
    
'    SetSPMProperty pstrSPMGroupName, pstrSPMPropertyName, ""
    
    Set objSpmManager = CreateObject("MTxSpm.SharedPropertyGroupManager.1")

    Set objSpmGroup = objSpmManager.CreatePropertyGroup(Trim(pstrSPMGroupName), 0, Process, blnGroupExists)
    
    If blnGroupExists = False Then
        Set objSpmGroup = objSpmManager.CreatePropertyGroup(Trim(pstrSPMGroupName), 0, Process, blnGroupExists)
    End If
    
    Set objSpmProperty = objSpmGroup.CreateProperty(Trim(pstrSPMPropertyName), blnPropertyExists)

    If blnPropertyExists = False Then
        Set objSpmProperty = objSpmGroup.CreateProperty(Trim(pstrSPMPropertyName), blnPropertyExists)
        GetSPMProperty = vbNullString
    Else
        
        If objSpmProperty.Value = 0 Or objSpmProperty.Value = vbNullString Then
            GetSPMProperty = vbNullString
        Else
            GetSPMProperty = objSpmProperty.Value
        End If
    End If

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GetSPMProperty Function", lngCodigoErroNegocio, intNumeroSequencialErro, "Grupo:" & pstrSPMGroupName & "-Propriedade:" & pstrSPMPropertyName)

End Function

'--------------------------------------------------------------------------------
'Bravar os dados no Shared Property Memory do COM+
'--------------------------------------------------------------------------------

Public Sub SetSPMProperty(ByVal pstrSPMGroupName As String, _
                          ByVal pstrSPMPropertyName As String, _
                          ByRef vntValue As Variant)

Dim objSpmManager                           As SharedPropertyGroupManager
Dim objSpmGroup                             As SharedPropertyGroup
Dim objSpmProperty                          As SharedProperty
Dim blnGroupExists                          As Boolean
Dim blnPropertyExists                       As Boolean

On Error GoTo ErrorHandler

    Set objSpmManager = CreateObject("MTxSpm.SharedPropertyGroupManager.1")
    
    Set objSpmGroup = objSpmManager.CreatePropertyGroup(Trim(pstrSPMGroupName), 0, Process, blnGroupExists)
    
    Set objSpmProperty = objSpmGroup.CreateProperty(Trim(pstrSPMPropertyName), blnPropertyExists)

    objSpmProperty.Value = vntValue

    Exit Sub

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "SetSPMProperty Function", lngCodigoErroNegocio, intNumeroSequencialErro, "Grupo:" & pstrSPMGroupName & "-Propriedade:" & pstrSPMPropertyName)

End Sub



