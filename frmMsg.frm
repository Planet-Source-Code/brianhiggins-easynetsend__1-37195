VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Easy Net Send"
   ClientHeight    =   2385
   ClientLeft      =   4365
   ClientTop       =   3435
   ClientWidth     =   6135
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6135
   Begin VB.CommandButton cmdHide 
      Cancel          =   -1  'True
      Height          =   195
      Left            =   6360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   135
   End
   Begin VB.ComboBox cboUserName 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Enter the computer or username you want to send the message to"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   5655
   End
   Begin VB.TextBox txtMsg 
      Height          =   1215
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Shift+Enter = Carrage Return,  Enter = Send"
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblInfo 
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Message:"
      Height          =   195
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   690
   End
   Begin VB.Label lblCompName 
      AutoSize        =   -1  'True
      Caption         =   "Computer Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1185
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'*  Licensing:                                                                          *
'*      this is freeware, you may use it for anypurpose in any enviroment you want.     *
'*      if used in a busness enviroment, you are REQUIRED to send me an E-Mail          *
'*      informing me of the use of my code/application(modified or unmodified).         *
'*      you may modify anything you want, as long as you do not redistrubute any        *
'*      modified code (aka, you can modify for personal use, but you are not authorized *
'*      to redistrubute any portion of this program without permission from me.         *
'*      for permission to redistrubute along with anything else, E-Mail me at:          *
'*      higgman16@yahoo.com                                                             *
'*      I do ask that you E-Mail me and let me know what you think of the program       *
'*      and if you continue to use it for more then 10 days, I would like to know.      *
'*  Limitations:                                                                        *
'*      Does not send to usernames, only computer names                                 *
'*      ONLY will run on NT3.51, NT4, 2K, XP, .NET (and other future NT based OS's      *
'*      in this product family)                                                         *
'*      AKA does NOT work on Win95, Win98, WinME there is no error handeling or         *
'*      version checking for this so when it crashes on an unsupported OS, don't        *
'*      E-Mail me telling me my code sucks                                              *
'*                                                                                      *
'*  Comments:                                                                           *
'*      If you have any ideas, sugestions or questions or if you know how to send       *
'*      to a username, feel free to E-Mail me at higgman16@yahoo.com I check that       *
'*      account once a week normally.                                                   *
'*      There is very little error handling in this code, I just didnt really see       *
'*      much need and I've wrote it in my spare time at work, so it hasn't been         *
'*      much of a priority.                                                             *
'*      I know there are a few things in here that aren't coded the nicest, I was       *
'*      going for funcionality and not proper coding practices in all cases             *
'*      all the API calls and associated componets I found online and claim no          *
'*      responsibility for them.                                                        *
'****************************************************************************************

Option Explicit
Dim Names() As String

Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_BAD_NETPATH As Long = 53
Private Const ERROR_INVALID_PARAMETER As Long = 87
Private Const ERROR_NOT_SUPPORTED As Long = 50
Private Const ERROR_INVALID_NAME As Long = 123
Private Const NERR_BASE As Long = 2100
Private Const NERR_SUCCESS As Long = 0
Private Const NERR_NetworkError As Long = (NERR_BASE + 36)
Private Const NERR_NameNotFound As Long = (NERR_BASE + 173)
Private Const NERR_UseNotFound As Long = (NERR_BASE + 150)
Private Const MAX_COMPUTERNAME As Long = 15
Private Const VER_PLATFORM_WIN32s As Long = 0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32_NT As Long = 2

Private Sub cboUserName_Change()
lblInfo = ""
End Sub

Private Sub cmdHide_Click()
'this command button must be kept visible in order for it to register as the default cancel button
'(when you hit escape) so I had to put the button off the edge of the visible form to work properly
Me.WindowState = vbMinimized
End Sub

Private Sub cmdSend_Click()
Dim Succuss As Integer
Dim uniToName As String
Dim uniMsg As String
Dim length As Integer
If cboUserName.Text = "" Or txtMsg = "" Then
    lblInfo = "Something is Blank..."
Else
    'Dim cmdLine As String
    lblInfo = "Sending..."
    
    'these 2 lines are not used anymore... they previously called a command window, not anymore
    'cmdLine = "net send " & Trim(cboUserName.Text) & " " & txtMsg
    'Call Shell(cmdLine, vbMinimizedNoFocus)
    
    uniToName = StrConv(Trim(cboUserName.Text), vbUnicode)
    uniMsg = StrConv(txtMsg, vbUnicode)
    length = Len(uniMsg)
    lblInfo = "Sending..."
    Me.MousePointer = vbHourglass
    
    Succuss = NetMessageBufferSend("", _
        uniToName, UserName, uniMsg, length)
    
    If Succuss = 0 Then
        lblInfo = "Sent..."
        AddName
    Else
        lblInfo = "Error..." & vbNewLine & GetNetSendMessageStatus(CLng(Succuss))
    End If
    DoEvents
    Me.MousePointer = vbNormal
    txtMsg.SetFocus
    txtMsg.SelStart = 0
    txtMsg.SelLength = Len(txtMsg)

End If
End Sub

Private Function GetNetSendMessageStatus(nError As Long) As String
    
   Dim msg As String
   
   Select Case nError
     Case NERR_SUCCESS:            msg = "The message was successfully sent"
     Case NERR_NameNotFound:       msg = "Name not found"
     Case NERR_NetworkError:       msg = "General network error occurred"
     Case NERR_UseNotFound:        msg = "Network connection not found"
     Case ERROR_ACCESS_DENIED:     msg = "Access to computer denied"
     Case ERROR_BAD_NETPATH:       msg = "Sent From server name not found."
     Case ERROR_INVALID_PARAMETER: msg = "Invalid parameter(s) specified."
     Case ERROR_NOT_SUPPORTED:     msg = "Network request not supported."
     Case ERROR_INVALID_NAME:      msg = "Illegal character or malformed name."
     Case Else:                    msg = "Unknown error executing command."
   End Select
   
   GetNetSendMessageStatus = msg
   
End Function

Private Sub Form_Activate()
If cboUserName.Text = "" Then
    cboUserName.SetFocus
End If

End Sub

Private Sub Form_Initialize()

If App.PrevInstance = True Then
    ActivatePrevInstance
End If

UserName = StrConv(GetLocalSystemName, vbUnicode)
ReDim Names(1 To 1)

Dim temp As String
On Error GoTo Done

Open App.Path & "\names.dat" For Input As #1
Do While EOF(1) = False
    Input #1, temp
    If temp <> "" Then
        If UBound(Names) >= 1 And Names(1) <> "" Then
            ReDim Preserve Names(1 To UBound(Names) + 1)
        End If
        Names(UBound(Names)) = temp
    End If
Loop
Done:

LoadList
Close #1
End Sub

Private Sub Form_Load()
'the form must be fully visible before calling Shell_NotifyIcon
Me.Show
Me.Refresh

    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Easy Net Send" & vbNullChar
    End With
    
Shell_NotifyIcon NIM_ADD, nid

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this procedure receives the callbacks from the System Tray icon.
    Dim result As Long
    Dim msg As Long
     'the value of X will vary depending upon the scalemode setting
     If Me.ScaleMode = vbPixels Then
      msg = X
     Else
      msg = X / Screen.TwipsPerPixelX
     End If
     
     Select Case msg
      'Case WM_LBUTTONUP        '514 restore form window
        ' Me.WindowState = vbNormal
        ' Result = SetForegroundWindow(Me.hwnd)
        ' Me.Show
      Case WM_LBUTTONDBLCLK    '515 restore form window
        Me.WindowState = vbNormal
        result = SetForegroundWindow(Me.hwnd)
        Me.Show
        txtMsg.SetFocus
        txtMsg.SelStart = 0
        txtMsg.SelLength = Len(txtMsg)
      Case WM_RBUTTONUP        '517 display popup menu
        result = SetForegroundWindow(Me.hwnd)
        Me.PopupMenu Me.mPopupSys
     End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Response As Integer

Response = MsgBox("Are you sure you want to exit? you will no " _
& "longer have easy access to send messages!" & vbNewLine & "Choosing Cancel will hide the applcation." _
, vbQuestion + vbYesNoCancel, "Are you sure?")
    
If Response = 7 Then
    'User chose No, does nothing
    Cancel = 1
ElseIf Response = 2 Then
    'user chose Cancel, no exit and hides the form..
    Cancel = 1
    Me.WindowState = vbMinimized
End If
    'if none of the above gets run, the user choose Yes, program exits normally

End Sub

Private Sub Form_Resize()
 'this is necessary to assure that the minimized window is hidden
  If Me.WindowState = vbMinimized Then
      Me.Hide
  End If
End Sub

Private Sub Form_Terminate()
'extra code normally, but supposidly will run upon an abnormal program error when it
'gets unexpectaly closed...
Shell_NotifyIcon NIM_DELETE, nid

End Sub

Private Sub Form_Unload(Cancel As Integer)
 'this removes the icon from the system tray
 Shell_NotifyIcon NIM_DELETE, nid
Dim i As Integer
Open App.Path & "\names.dat" For Output As #1
For i = 1 To UBound(Names)
    Print #1, Names(i)
Next i
 
End Sub

Private Sub lblCompName_DblClick()
Dim temp As String
temp = InputBox("Enter New UserName, or leave blank to reset to Default. ", "Set Send Name", StrConv(UserName, vbFromUnicode))
If temp <> "" Then
    UserName = StrConv(UCase(Trim(temp)), vbUnicode)
Else
    UserName = StrConv(GetLocalSystemName, vbUnicode)
End If
End Sub

Private Sub mPopExit_Click()
 'called when user clicks the popup menu Exit command
 Unload Me
End Sub

Private Sub mPopRestore_Click()
  'called when the user clicks the popup menu Restore command
  Dim result As Long
  Me.WindowState = vbNormal
  result = SetForegroundWindow(Me.hwnd)
  Me.Show
  txtMsg.SetFocus
  txtMsg.SelStart = 0
  txtMsg.SelLength = Len(txtMsg)
End Sub

Private Sub AddName()
Dim i As Integer
Dim Found As Boolean
Found = False
For i = 1 To UBound(Names)
    If UCase(cboUserName.Text) = UCase(Names(i)) Then
        Found = True
        Exit For
    End If
Next i

If Not Found Then
    If UBound(Names) = 1 And Names(1) = "" Then
        Names(UBound(Names)) = UCase(cboUserName.Text)
    Else
        ReDim Preserve Names(1 To UBound(Names) + 1)
        Names(UBound(Names)) = UCase(cboUserName.Text)
    End If
    LoadList
End If

End Sub

Private Sub LoadList()
Dim i As Integer
cboUserName.Clear
For i = 1 To UBound(Names)
    cboUserName.AddItem Names(i)
    cboUserName.ListIndex = cboUserName.NewIndex
Next i
End Sub

Private Sub txtMsg_Change()
lblInfo = ""
End Sub
