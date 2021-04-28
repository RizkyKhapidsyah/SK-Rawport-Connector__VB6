VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Raw Port Connector (Telnet)"
   ClientHeight    =   5250
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4560
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   12632256
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Website"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Picture         =   "Form1.frx":0080
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin MSWinsockLib.Winsock rdest 
      Left            =   8040
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock rsource 
      Left            =   7560
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   960
      Top             =   4440
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   4440
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   9
      Top             =   4905
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "10:21 AM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12/6/2001"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4057
            MinWidth        =   4057
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send Data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1440
      Top             =   4440
   End
   Begin RichTextLib.RichTextBox text3 
      Height          =   3255
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1320
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5741
      _Version        =   393217
      BackColor       =   14737632
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0FED
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   9000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Port"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Host"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuPortRe 
         Caption         =   "Port Redirection"
      End
   End
   Begin VB.Menu stats 
      Caption         =   "Connection Statistics"
      Begin VB.Menu state 
         Caption         =   "Winsock Connection State"
      End
      Begin VB.Menu recd 
         Caption         =   "Bytes Received"
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
      Begin VB.Menu abtAuthor 
         Caption         =   "About the Author"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myvar As String
Dim connected As Boolean
Dim bannerdata()
Dim totalbytes As Variant
Dim isincombo As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation _
      As String, ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
      

Private Sub abtAuthor_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub Command1_Click()

If Text1 <> "" And Text2 <> "" Then
    If Not Winsock1.State = sckConnected Then
    
        Winsock1.Close
        
        Winsock1.RemoteHost = Text1.Text
            If Text2.Text = "" Then
                Winsock1.RemotePort = 23
            Else
                Winsock1.RemotePort = CInt(Text2.Text)
            End If
        Winsock1.Connect
    Else
        MsgBox ("You are already connected, Disconnect to reconnect.")
    End If
Else
    MsgBox ("Please complete host information.")
End If

End Sub

Private Sub Command2_Click()
Call ShellExecute(0&, vbNullString, "http://randy.onet.net", _
vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Command3_Click()
Winsock1.Close
connected = False
End Sub

Private Sub Command4_Click()

If connected = False Then
    MsgBox ("You are not currently Connected, please connect.")
Else
    Winsock1.SendData RichTextBox1.Text & vbCrLf
    text3 = text3.Text & RichTextBox1.Text & vbCrLf
    RichTextBox1.Text = ""
    text3.SelStart = Len(text3)
End If

End Sub

Private Sub Form_Load()
Text1.TabIndex = 0
totalbytes = 0

End Sub

Private Sub mnuPortRe_Click()
vlocalport = InputBox("Enter the Local Port:")
vremoteport = InputBox("Enter the Remote Port:")
vremotehost = InputBox("Enter the Remote Host:")

If vlocalport = "" Then
    mboxincomplete
    Exit Sub
ElseIf vremoteport = "" Then
    mboxincomplete
    Exit Sub
ElseIf vremotehost = "" Then
    mboxincomplete
    Exit Sub
Else

rsource.LocalPort = vlocalport
rdest.RemotePort = vremoteport
rdest.RemoteHost = vremotehost
rsource.Listen
MsgBox ("Relaying Data.")
End If

End Sub

Private Sub Picture1_Click()
Call ShellExecute(0&, vbNullString, "http://randy.onet.net", _
vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub rdest_Close()
rsource.Close
End Sub

Private Sub rdest_DataArrival(ByVal bytesTotal As Long)
rdest.GetData indatabuffer
rsource.SendData indatabuffer
indatabuffer = ""
End Sub

Private Sub recd_Click()
MsgBox (totalbytes & " bytes received.")
End Sub

Private Sub rsource_Close()
rdest.Close
End Sub

Private Sub rsource_ConnectionRequest(ByVal requestID As Long)

relayinprogress = True

If rsource.State <> sckConnected Then rsource.Close

rsource.Accept requestID
rdest.Connect

End Sub

Private Sub rsource_DataArrival(ByVal bytesTotal As Long)
rsource.GetData outdatabuffer
rdest.SendData outdatabuffer
outdatabuffer = ""
End Sub

Private Sub state_Click()
MsgBox ("Connection State #: " & Winsock1.State)
End Sub

Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command1_Click
    RichTextBox1.SetFocus
End If
End Sub

Private Sub richtextbox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command4_Click
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(3) = "Connection State: " & Winsock1.State
End Sub

Private Sub Timer2_Timer()
StatusBar1.Panels(4) = "Connected: " & connected
End Sub

Private Sub Timer3_Timer()
If rsource.State = sckConnected Or rsource.State = sckListening Then
StatusBar1.Panels(5) = "Relaying"
Else
StatusBar1.Panels(5) = "Relay Closed"
End If
End Sub

Private Sub Winsock1_Close()
connected = False
text3.Text = "Disconnected from remote host."
totalbytes = 0
End Sub

Private Sub Winsock1_Connect()
text3.Text = "Connecting to " & Winsock1.RemoteHost & "..." & vbCrLf
connected = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
totalbytes = totalbytes + bytesTotal
Winsock1.GetData myvar
text3.Text = text3.Text & myvar
text3.SelStart = Len(text3)
'Winsock1.Close
'connected = False
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox ("An Error Has Occured, the Active Connection has been disconnected, Please Reconnect.")
If Not Winsock1.State = 4 Then Winsock1.Close
End Sub

Public Function mboxincomplete()
MsgBox ("Incomplete Information.  NOT relaying.")
End Function

