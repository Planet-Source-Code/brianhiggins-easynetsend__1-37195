Attribute VB_Name = "modSystemTray"
Public Const GW_HWNDPREV = 3
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Declare Function NetMessageBufferSend Lib "Netapi32.dll" (ByVal sServerName$, ByVal sMsgName$, ByVal sFromName$, ByVal sMessageText$, ByVal lBufferLength&) As Long
Declare Function NetWkstaGetInfo Lib "Netapi32.dll" (ByVal sServerName$, ByVal lLevel&, vBuffer As Any) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (dest As Any, vSrc As Any, ByVal lSize&)
Public Declare Sub lstrcpyW Lib "kernel32" (vDest As Any, ByVal sSrc As Any)

Public nid As NOTIFYICONDATA

Public Type NOTIFYICONDATA
'user defined type required by Shell_NotifyIcon API call
   cbSize As Long
   hwnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
 End Type
 
Type WKSTA_INFO_100
    wki100_platform_id As Long
    wki100_computername As Long
    wki100_langroup As Long
    wki100_ver_major As Long
    wki100_ver_minor As Long
End Type

'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public UserName As String

Sub ActivatePrevInstance()

Dim OldTitle As String
Dim PrevHndl As Long
Dim result As Long
    
    'Save the title of the application.
    OldTitle = App.Title
    
    'Rename the title of this application so FindWindow
    'will not find this application instance.
    App.Title = "unwanted instance"
    'Attempt to get window handle using VB4 class name.
    PrevHndl = FindWindow("ThunderRTMain", OldTitle)
    'Check for no success.
    If PrevHndl = 0 Then
        'Attempt to get window handle using VB5 class name.
        PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
    End If
    'Check if found
    If PrevHndl = 0 Then
        'Attempt to get window handle using VB6 class name
        PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
    End If
    'Check if found
    If PrevHndl = 0 Then
        'No previous instance found.
        Exit Sub
    Else
        'Get handle to previous window.
        PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
        'Restore the program.
        result = OpenIcon(PrevHndl)
        'Activate the application.
        result = SetForegroundWindow(PrevHndl)
        'End the application.
        End
    End If
    
End Sub

Public Function GetLocalSystemName()

    Const csProcName As String = "GetLocalSystemName"
    Dim lReturnCode As Long
    Dim bBuffer(512) As Byte
    Dim i As Integer
    Dim twkstaInfo100 As WKSTA_INFO_100, lwkstaInfo100 As Long
    Dim lwkstaInfo100StructPtr As Long
    Dim sLocalName As String
    
    lReturnCode = NetWkstaGetInfo("", 100, lwkstaInfo100)
 
    lwkstaInfo100StructPtr = lwkstaInfo100
                 
    If lReturnCode = 0 Then
                 
        RtlMoveMemory twkstaInfo100, ByVal lwkstaInfo100StructPtr, Len(twkstaInfo100)
         
        lstrcpyW bBuffer(0), twkstaInfo100.wki100_computername

        i = 0
        Do While bBuffer(i) <> 0
            sLocalName = sLocalName & Chr(bBuffer(i))
            i = i + 2
        Loop
            
        GetLocalSystemName = sLocalName
         
    End If
    
End Function

