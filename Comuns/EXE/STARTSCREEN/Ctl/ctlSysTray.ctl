VERSION 5.00
Begin VB.UserControl ctlSysTray 
   CanGetFocus     =   0   'False
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   390
   ClipControls    =   0   'False
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   MouseIcon       =   "ctlSysTray.ctx":0000
   Picture         =   "ctlSysTray.ctx":030A
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   26
   ToolboxBitmap   =   "ctlSysTray.ctx":2434
End
Attribute VB_Name = "ctlSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'-------------------------------------------------------
' Control Property Globals...
'-------------------------------------------------------
Private gblnInTray                          As Boolean
Private glngTrayId                          As Long
Private gstrTrayTip                         As String
Private glngTrayHwnd                        As Long
Private gTrayIcon                           As StdPicture
Private gblnAddedToTray                     As Boolean

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Const MAX_SIZE = 390

Private Const blnDefInTray = False
Private Const strDefTrayTip = "SLCC Control." & vbNullChar

Private Const strInTray = "InTray"
Private Const strTrayIcon = "TrayIcon"
Private Const strTrayTip = "TrayTip"

'-------------------------------------------------------
' Control Events...
'-------------------------------------------------------
Public Event MouseMove(Id As Long)
Public Event MouseDown(Button As Integer, Id As Long)
Public Event MouseUp(Button As Integer, Id As Long)
Public Event MouseDblClick(Button As Integer, Id As Long)

Private Sub UserControl_Initialize()

On Error GoTo ErrorHandler

    gblnInTray = blnDefInTray                             ' Set global InTray default
    gblnAddedToTray = False                            ' Set default state
    glngTrayId = 0                                     ' Set global TrayId default
    glngTrayHwnd = hwnd                                ' Set and keep HWND of user control

    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "UserControl_Initialize Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub UserControl_InitProperties()

On Error GoTo ErrorHandler
    
    InTray = blnDefInTray                              ' Init InTray Property
    TrayTip = strDefTrayTip                            ' Init TrayTip Property
    Set TrayIcon = Picture                          ' Init TrayIcon property
    
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "UserControl_InitProperties Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

Private Sub UserControl_Paint()

Dim edge                                    As RECT ' Rectangle edge of control

On Error GoTo ErrorHandler
    
    edge.Left = 0                                   ' Set rect edges to outer
    edge.Top = 0                                    ' - most position in pixels
    edge.Bottom = ScaleHeight                       '
    edge.Right = ScaleWidth                         '
    DrawEdge hDC, edge, BDR_RAISEDOUTER, BF_RECT Or BF_SOFT ' Draw Edge...
    
    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "UserControl_Paint Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error GoTo ErrorHandler

    ' Read in the properties that have been saved into the PropertyBag...
    With PropBag
        InTray = .ReadProperty(strInTray, blnDefInTray)       ' Get InTray
        Set TrayIcon = .ReadProperty(strTrayIcon, Picture) ' Get TrayIcon
        TrayTip = .ReadProperty(strTrayTip, strDefTrayTip)    ' Get TrayTip
    End With

    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "UserControl_ReadProperties Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error GoTo ErrorHandler

    With PropBag
        .WriteProperty strInTray, gblnInTray                 ' Save InTray to propertybag
        .WriteProperty strTrayIcon, gTrayIcon             ' Save TrayIcon to propertybag
        .WriteProperty strTrayTip, gstrTrayTip               ' Save TrayTip to propertybag
    End With

    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "UserControl_WriteProperties Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    
    Height = MAX_SIZE                   ' Prevent Control from being resized...
    Width = MAX_SIZE
    
    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "UserControl_Resize Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

Private Sub UserControl_Terminate()
On Error GoTo ErrorHandler
    
    If InTray Then                      ' If TrayIcon is visible
        InTray = False                  ' Cleanup and unplug it.
    End If
    
    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "UserControl_Terminate Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

Public Property Set TrayIcon(Icon As StdPicture)

Dim Tray                                    As NOTIFYICONDATA   ' Notify Icon Data structure
Dim lngRc                                   As Long             ' API return code

On Error GoTo ErrorHandler

    
    If Not (Icon Is Nothing) Then                       ' If icon is valid...
        If (Icon.Type = vbPicTypeIcon) Then             ' Use ONLY if it is an icon
            If gblnAddedToTray Then                        ' Modify tray only if it is in use.
                Tray.uID = glngTrayId                      ' Unique ID for each HWND and callback message.
                Tray.hwnd = glngTrayHwnd                   ' HWND receiving messages.
                Tray.hIcon = Icon.Handle                ' Tray icon.
                Tray.uFlags = NIF_ICON                  ' Set flags for valid data items
                Tray.cbSize = Len(Tray)                 ' Size of struct.
                
                lngRc = Shell_NotifyIcon(NIM_MODIFY, Tray) ' Send data to Sys Tray.
            End If
    
            Set gTrayIcon = Icon                        ' Save Icon to global
            Set Picture = Icon                          ' Show user change in control as well(gratuitous)
            PropertyChanged strTrayIcon                   ' Notify control that property has changed.
        End If
    End If


    Exit Property
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "TrayIcon Property", lngCodigoErroNegocio, intNumeroSequencialErro)

End Property

Public Property Get TrayIcon() As StdPicture
On Error GoTo ErrorHandler

    Set TrayIcon = gTrayIcon                        ' Return Icon value
    
    Exit Property
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "TrayIcon Property", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Property

Public Property Let TrayTip(Tip As String)
Attribute TrayTip.VB_ProcData.VB_Invoke_PropertyPut = ";Misc"
Attribute TrayTip.VB_UserMemId = -517

Dim Tray                                    As NOTIFYICONDATA   ' Notify Icon Data structure
Dim lngRc                                      As Long             ' API Return code

On Error GoTo ErrorHandler
   
    If gblnAddedToTray Then                            ' if TrayIcon is in taskbar
        Tray.uID = glngTrayId                          ' Unique ID for each HWND and callback message.
        Tray.hwnd = glngTrayHwnd                       ' HWND receiving messages.
        Tray.szTip = Tip & vbNullChar               ' Tray tool tip
        Tray.uFlags = NIF_TIP                       ' Set flags for valid data items
        Tray.cbSize = Len(Tray)                     ' Size of struct.
        
        lngRc = Shell_NotifyIcon(NIM_MODIFY, Tray)     ' Send data to Sys Tray.
    End If
    
    gstrTrayTip = Tip                                  ' Save Tip
    PropertyChanged strTrayTip                        ' Notify control that property has changed

    Exit Property
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "TrayTip Property", lngCodigoErroNegocio, intNumeroSequencialErro)

End Property

Public Property Get TrayTip() As String
On Error GoTo ErrorHandler

    TrayTip = gstrTrayTip                              ' Return Global Tip...
    
    Exit Property
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "TrayTip Property", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Property

Public Property Let InTray(Show As Boolean)
Attribute InTray.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"

Dim lngClassAddr                               As Long ' Address pointer to Control Instance

On Error GoTo ErrorHandler
    
    If (Show <> gblnInTray) Then                       ' Modify ONLY if state is changing!
        If Show Then                                ' If adding Icon to system tray...
            If Ambient.UserMode Then                ' If in RunMode and not in IDE...
                 ' SubClass Controls window proc.
                PrevWndProc = SetWindowLong(glngTrayHwnd, GWL_WNDPROC, AddressOf SubWndProc)
                
                ' Get address to user control object
                'CopyMemory lngClassAddr, UserControl, 4&
                
                ' Save address to the USERDATA of the control's window struct.
                ' this will be used to get an object reference to the control
                ' from an HWND in the callback.
                SetWindowLong glngTrayHwnd, GWL_USERDATA, ObjPtr(Me) 'lngClassAddr
                
                AddIcon glngTrayHwnd, glngTrayId, TrayTip, TrayIcon ' Add TrayIcon to System Tray...
                gblnAddedToTray = True                 ' Save state of control used in teardown procedure
            End If
        Else                                        ' If removing Icon from system tray
            If gblnAddedToTray Then                    ' If Added to system tray then remove...
                DeleteIcon glngTrayHwnd, glngTrayId       ' Remove icon from system tray
                
                ' Un SubClass controls window proc.
                SetWindowLong glngTrayHwnd, GWL_WNDPROC, PrevWndProc
                gblnAddedToTray = False                ' Maintain the state for teardown purposes
            End If
        End If
        
        gblnInTray = Show                              ' Update global variable
        PropertyChanged strInTray                     ' Notify control that property has changed
    End If

    Exit Property
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "InTray Property", lngCodigoErroNegocio, intNumeroSequencialErro)

End Property

Public Property Get InTray() As Boolean
On Error GoTo ErrorHandler

    InTray = gblnInTray                                ' Return global property
    
    Exit Property
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "InTray Property", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Property

Private Sub AddIcon(hwnd As Long, Id As Long, Tip As String, Icon As StdPicture)

Dim Tray                                    As NOTIFYICONDATA   ' Notify Icon Data structure
Dim lngFlags                                As Long             ' Tray action flag
Dim lngRc                                   As Long             ' API return code

On Error GoTo ErrorHandler

    
    Tray.uID = Id                                   ' Unique ID for each HWND and callback message.
    Tray.hwnd = hwnd                                ' HWND receiving messages.
    
    If Not (Icon Is Nothing) Then                   ' Validate Icon picture
        Tray.hIcon = Icon.Handle                    ' Tray icon.
        Tray.uFlags = Tray.uFlags Or NIF_ICON       ' Set ICON flag to validate data item
        Set gTrayIcon = Icon                        ' Save icon
    End If
    
    If (Tip <> "") Then                             ' Validate Tip text
        Tray.szTip = Tip & vbNullChar               ' Tray tool tip
        Tray.uFlags = Tray.uFlags Or NIF_TIP        ' Set TIP flag to validate data item
        gstrTrayTip = Tip                              ' Save tool tip
    End If
    
    Tray.uCallbackMessage = TRAY_CALLBACK           ' Set user defined message
    Tray.uFlags = Tray.uFlags Or NIF_MESSAGE        ' Set flags for valid data item
    Tray.cbSize = Len(Tray)                         ' Size of struct.
    
    lngRc = Shell_NotifyIcon(NIM_ADD, Tray)            ' Send data to Sys Tray.

    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "AddIcon Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Public Sub DeleteIcon(hwnd As Long, Id As Long)

Dim Tray                                    As NOTIFYICONDATA   ' Notify Icon Data structure
Dim lngRc                                   As Long             ' API return code

On Error GoTo ErrorHandler
    
    Tray.uID = Id                                   ' Unique ID for each HWND and callback message.
    Tray.hwnd = hwnd                                ' HWNDreceiving messages.
    Tray.uFlags = 0&                                ' Set flags for valid data items
    Tray.cbSize = Len(Tray)                         ' Size of struct.
    
    DoEvents
    lngRc = Shell_NotifyIcon(NIM_DELETE, Tray)         ' Send delete message.
    DoEvents
    
    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlSysTray", "DeleteIcon Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Friend Sub SendEvent(MouseEvent As Long, Id As Long)
    
    Select Case MouseEvent                          ' Dispatch mouse events to control
        Case WM_MOUSEMOVE
            RaiseEvent MouseMove(Id)
        Case WM_LBUTTONDOWN
            RaiseEvent MouseDown(vbLeftButton, Id)
        Case WM_LBUTTONUP
            RaiseEvent MouseUp(vbLeftButton, Id)
        Case WM_LBUTTONDBLCLK
            RaiseEvent MouseDblClick(vbLeftButton, Id)
        Case WM_RBUTTONDOWN
            RaiseEvent MouseDown(vbRightButton, Id)
        Case WM_RBUTTONUP
            RaiseEvent MouseUp(vbRightButton, Id)
        Case WM_RBUTTONDBLCLK
            RaiseEvent MouseDblClick(vbRightButton, Id)
    End Select

End Sub
