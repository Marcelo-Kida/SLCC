VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsTab"
Attribute VB_Ext_KEY = "Member0" ,"clsTab"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private llCodigoErroNegocio                 As Long
Private liNumeroSequencialErro              As Integer

Private mcol                                As Collection

Public Function Add(Optional Index As Long, _
                    Optional Key As String, _
                    Optional Caption As String, _
                    Optional Description As String, _
                    Optional Visible As Boolean, _
                    Optional Enabled As Boolean) As clsTab
    
On Error GoTo ErrorHandler
    
    'create a new object
    Dim objNewMember As clsTab
    Set objNewMember = New clsTab


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.Index = Index
    objNewMember.Caption = Caption
    objNewMember.Description = Description
    objNewMember.Visible = Visible
    objNewMember.Enabled = Enabled
    
    If Len(Key) = 0 Then
        mcol.Add objNewMember
    Else
        mcol.Add objNewMember, Key
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

    Exit Function
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colTabs", "Add Function", llCodigoErroNegocio, liNumeroSequencialErro)

End Function

Public Property Get Item(vntIndexKey As Variant) As clsTab
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrorHandler
Dim objTab                                  As clsTab
    
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    
    If IsNumeric(vntIndexKey) Then
        Set Item = mcol(vntIndexKey)
    Else
        For Each objTab In Me
            If objTab.Key = vntIndexKey Then
                Set Item = objTab
                Exit For
            End If
        Next
    End If
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colTabs", "Item Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Count() As Long
On Error GoTo ErrorHandler

    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mcol.Count
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colTabs", "Count Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Sub Remove(vntIndexKey As Variant)
On Error GoTo ErrorHandler
    
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    mcol.Remove vntIndexKey
    
    Exit Sub
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colTabs", "Remove Sub", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
On Error GoTo ErrorHandler

    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mcol.[_NewEnum]
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colTabs", "NewEnum Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Private Sub Class_Initialize()
On Error GoTo ErrorHandler

    'creates the collection when this class is created
    Set mcol = New Collection
    
    Exit Sub
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colTabs", "Class_Initialize Sub", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrorHandler

    'destroys collection when this class is terminated
    Set mcol = Nothing
    
    Exit Sub
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colTabs", "Class_Terminate Sub", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Sub
