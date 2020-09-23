VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Telnet Server"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton RecvBtn 
      Caption         =   "RecvBtn"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton AcceptBtn 
      Caption         =   "AcceptBtn"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton KillBtn 
      Caption         =   "Kill"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton StartBtn 
      Caption         =   "Start"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox ServerWindow 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   2055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////
'// Project name: Full Comunacation telnet server.
'// By : No()ne
'// Email: data_tune@hotmail.com
'//
'// Bassed around "Tconsole" using winsock.ocx
'//
'// http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=5089
'//
'//////////////////////////////
'// This is the bare bones of a telnet server.
'// I tried to demonstrates recieveing and sending
'// data to a host. Also demonstrates buffer
'// switching to control data input.
'//
'// The future is your's
'////////////////

Dim Start_up_Data As WSADataType


Private Sub Form_Load()
    
    ' Clear server window
    '
    ServerWindow.Text = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    X = WSACleanup()
    
End Sub

Private Sub KillBtn_Click()

    X = WSACleanup()
    ServerWindow.Text = ServerWindow.Text & vbCrLf & "** telnet Halted **"
    
    ' Button control
    '
    StartBtn.Enabled = True
    KillBtn.Enabled = False
    
End Sub

Private Sub StartBtn_Click()
    
    ' Button Control
    StartBtn.Enabled = False
    KillBtn.Enabled = True


    
    ' This starts up winsock.
    '
    X = WSACleanup()
    X = WSAStartup(&H101, Start_up_Data)
    
        If (X = SOCKET_ERROR) Then
            Exit Sub
        End If
    
    
    'Create our socket (Install a Phone)
    '
    Socket_Number = socket(AF_INET, SOCK_STREAM, 0)
    
        If (Socket_Number = SOCKET_ERROR) Then
            Exit Sub
        End If
    
    
    'Now we give our phone a phone number.
    '
    Socket_Buffer.sin_family = AF_INET
    Socket_Buffer.sin_port = htons(23)      '// Port 23 is telnet.
    Socket_Buffer.sin_addr = 0
    Socket_Buffer.sin_zero = String$(8, 0)
    
                        
    ' Binding is giving our sock a local name.
    '
    X = bind(Socket_Number, Socket_Buffer, sockaddr_size)
    
        If X <> 0 Then
            X = WSACleanup()
            Exit Sub
        End If
    
    
    ' Now plug the phone in and hope you have a
    ' friend who will call you ;)
    '
    X = listen(Socket_Number, 1)
    
    
    ' Perform asynchronous version of select()
    ' Tell the sock how to behave.
    ' FD_CONNECT and FD_ACCEPT
    ''
    X = WSAAsyncSelect(Socket_Number, AcceptBtn.hWnd, &H202, FD_CONNECT Or FD_ACCEPT)
    
    
    ' Tell me when the server was started.
    '
    Dim dte As Date
    dte = Now
    
    ServerWindow.Text = "Sever started at " & dte & vbCrLf
    
    
    ' Tell the server Window what port we are using.
    ' Complacated but shows the use of htons.
    '
    ServerWindow.Text = ServerWindow.Text _
                    & "Listening on port: " _
                    & htons(Socket_Buffer.sin_port) _
                    & vbCrLf
                    
    
    
End Sub


Private Sub AcceptBtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Mouse up is a steady resource for a loop.
    '
    ' This part waits for a connection, then it
    ' will hand it over to our loop for recieving
    ' and the handleing of the in coming data.
    '
    Read_Sock = accept(Socket_Number, Remote_Sock_Buffer, Len(Remote_Sock_Buffer))
    
    
    ' When it comes to sockets, Windows is a square
    ' peg in a round hole. WSA* are socket commands
    ' only found in windows. But this kinda binds
    ' a connection to another message loop (or window).
    '
    X = WSAAsyncSelect(Read_Sock, RecvBtn.hWnd, ByVal &H202, ByVal FD_READ Or FD_CLOSE)
    
    
    ' Now we set up our console window. When dealing with
    ' Security issues, logging is a must. (Time stamped)
    '
    Dim daAt As Date
    
        daAt = Now
    
    
    ' getpeeraddress() is a function from the helper .bas
    ' To explain every thing that makes it
    ' work, is beound the scope of this project.
    '
    Dim Remote_iP As String
    
        Remote_iP = GetPeerAddress(Read_Sock)
    
    
    ' Now we log it :)
    '
    ServerWindow = ServerWindow & "Connection attemp by: " & Remote_iP & vbCrLf
    ServerWindow = ServerWindow & "Time of connection  : " & daAt & vbCrLf
    
    Call sendHeaderz
    Call SendPrompt
    
End Sub

Private Sub RecvBtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next
'// This alows for the Socket Stay open

    Bytes = recv(Read_Sock, Read_Buffer, 1024, 0)
    
    If Bytes <> 0 Then
        
        ' If the Enter key was pressed, then send
        ' our command in the data buffer to
        ' DataControl function and clear out
        ' our buffers.
        '
        If Left$(Read_Buffer, Bytes) = vbCrLf Then
            
            '// testing message
            '
            newMessage = "You typed in: " + Data_buffer + vbCrLf
            X = SendIt(Read_Sock, newMessage)
            
            Call SendPrompt
            
            '//dosomething (Data_buffer)
            Read_Buffer = ""
            Data_buffer = ""
            Exit Sub
        
        Else
        
        
        ' Else if enter has NOT been pressed, then
        ' continue to build the command buffer.
        '
        Data_buffer = Data_buffer + Left$(Read_Buffer, Bytes)
        
        End If
        
    End If
End Sub


