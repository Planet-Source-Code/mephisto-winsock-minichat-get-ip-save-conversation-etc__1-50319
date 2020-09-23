VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MiniChat"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5160
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Get IP"
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   5760
      TabIndex        =   10
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Change port"
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save Conversation"
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox st 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   5895
   End
   Begin VB.TextBox t 
      Height          =   5055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "CHAT2.frx":0000
      Top             =   1200
      Width           =   7215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send"
      Height          =   255
      Left            =   6120
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Host"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "MINICHAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Click here and select your nick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Nlabel1 
      Caption         =   "NickName:"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim port As Single 'here we will store the port we are going to be using
                   'port should always be higher than 5000, because there is a chance that
                   'other applications can use the ports you specify...
Dim Nick As String 'Nick stores your Nickname
Dim Data As String 'This is more like a temporary thing, we will see its use in GetData procedure
Dim Data2() As String 'Data2 is an array because we will use Split function to get from Data all then
                      'info we need. This can be located at the DataArrival Sub of Winsock1

Private Sub Command1_Click()
'if the user clicks CONNECT button, we need to get the HOST IP adress...
HostIp = InputBox("Enter the host's computer name or ip address:" & vbCrLf & "(Be careful not to include any unnecessary spaces etc or error message will be generated.")
Winsock1.Connect HostIp, port 'We use this command to Connect to the Host [HostIP] on port [port]
End Sub

Private Sub Command2_Click()
'If user clicks HOST, assign local port
Winsock1.LocalPort = port
'Listen for incoming connections
Winsock1.Listen

'MsgBox "Host Successfully created"
End Sub

Private Sub Command3_Click()
'Ok this line may be confusing, this is the core of the program. This happens when you click Send.
'Every time you want to send a message to the host, you are in fact sending two things.
'1 nickname
'2 message
'and the next line combines these two, and adds a dummie between them. When host receives this text,
'it will use Split function to split the Nick from the message and print both on the window
'===========>this, you can see in the DataArrival Sub
Winsock1.SendData Nick & "~!@#" & st.Text
'the "~!@#" is just a dummie

'add your own message to the chat textbox
'vbCrlf tells VisualBasic to put the text behind it in a new line
t.Text = t.Text & vbCrLf & Nick & ": " & st.Text

'erase text in the field you typed message in
st.Text = ""
'Once you click on the button, you may want to type a message again, so the textbox gets focus
st.SetFocus
'This line makes sure that the scroll bar is always in the very down position. Try removing this
'command and see what happens when you flood the textbox. The Scroll bar would stay up...
t.SelStart = Len(t)
End Sub

Private Sub Command4_Click()
'These all are only temporary Dims, you dont need them anywhere else only in this Sub
Dim a As Integer, b As Integer
Dim strURL As String, strIP As String

'open the www.whatismyip.com URL
 strURL = Inet1.OpenURL("http://www.whatismyip.com/")
        
        'Locates the part before the IP
    a = InStr(1, strURL, "<TITLE>Your ip is ")
        'Locates the part after IP
    b = InStr(1, strURL, " WhatIsMyIP.com</TITLE>")
        'Gets the IP adress between the two above
    strIP = Mid(strURL, a + 18, b - (a + 18))
        'Print the IP in textbox
    txtIP.Text = strIP
End Sub

Private Sub Command5_Click()
'Save Conversation, very easy process

'All the Junk :) Just so the user knows what the TXT file is about if he opens it once...
'If you open a file for output, it will be created by default if it doesnt exist
Open "C:\Conversation.txt" For Output As #1
Print #1, "**********MiniChat Record of conversation**********"
Print #1, "This Conversation took place in"
Print #1, "Date:" & Date
Print #1, "Time: " & Time
Print #1, ""
'print all the stuff in the TXT file
Print #1, Form1.t.Text
'close file
Close #1
'inform user
MsgBox ("File saved in C:\Conversation.txt")
End Sub

Private Sub Command6_Click()
'this makes sure the application doesnt crash if the user types in some unexpected junk, it only beeps
'try removing this, and click change port, and type in ASDF, the application will crash
On Error Resume Next
'assign new port number
port = InputBox("Enter new port : (DEFAULT = 5432)")
'if user clicked Cancel port will be 0, so we need to set it back up to 5432
If port = 0 Then
port = 5432
ElseIf port < 1000 Then
'inform user, but dont change the port... if they want, they can
MsgBox "I do not recommend setting the port number so low. Port should be more then 5000."
End If
End Sub

Private Sub Form_Load()
'assign default port number, you may change this if you like... this is a safe port though
port = 5432
'the nick is the crap typed in the label1 in the beggining
Nick = Label1.Caption
'make sure there is no junk typed in the chat window
t.Text = ""
End Sub

Private Sub Label1_Click()
'WHen user clicks on his Nick, it should change. So we ask him for new nick
Nick = InputBox("Enter new NickName")

'You may add something like this if you wish to limit people on the lenght of their nicknames in letters

'If Len(Nick) > 10 Then
'Nick = "NoName"
'MsgBox "Too long Nick!"
'End If

'print the new nick
Label1.Caption = Nick
End Sub

Private Sub st_KeyUp(KeyCode As Integer, Shift As Integer)
'This occurs when the user pushes a button
'vbkeyreturn simply means enter
If KeyCode = vbKeyReturn Then
'do same as if the user clicked command3 , a.k.a. Send
Call Command3_Click
End If
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close  'Got to do this to make sure the Winsock control isn't already being used.
Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'if a data is received, this event occurs
'so what we do is we read in the incoming information into Data
Winsock1.GetData Data, vbString, bytesTotal
'here, we use Split function to separate the Nick from the message and store both in Data2
Data2() = Split(Data, "~!@#")
'at last, we print first print Nickname, then we add colon and then the message
t.Text = t.Text & vbCrLf & Data2(0) & ": " & Data2(1)
'to make sure that the scroll bar goes way down again...
t.SelStart = Len(t)
End Sub


