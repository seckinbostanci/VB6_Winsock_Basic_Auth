VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Authentication Example"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CheckBox chkSendKey 
      Caption         =   "Send the correct key"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "127.0.0.1"
      Top             =   3120
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   5520
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1414
   End
   Begin RichTextLib.RichTextBox rtbLog 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3625
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Created by Andrew Armstrong"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Server Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":0082
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()
    If cmdConnect.Caption = "&Connect" Then
        sckClient.Connect txtAddress.Text 'Connect to the specified address
        cmdConnect.Caption = "Dis&connect"
    Else
        sckClient.Close 'Close the connection
        cmdConnect.Caption = "&Connect"
        HasAuthenticated = False
        Log "Connection closed."
    End If
End Sub
    
Private Sub cmdQuit_Click()
    sckClient.Close
    Unload Me
    End
End Sub
    
Private Sub Form_Load()
MsgBox "Created by Andrew Armstrong. Contact me at: " & vbCrLf & _
 "andrewa@bigpond.net.au or ICQ# 14344635 !" & vbCrLf & _
 "Please contact me if you ever do use this in an application, I would be interested to know! :)", vbInformation, "Contact/About"
End Sub

Private Sub rtbLog_Change()
    rtbLog.SelStart = Len(rtbLog) 'Scroll to the bottem of the textbox so you see the most
    'recent event
    
End Sub
    
Private Sub sckClient_Close()
    Log "Connection Closed." 'Log the event
    sckClient.Close 'Just cleanup
    HasAuthenticated = False
    cmdConnect.Caption = "&Connect"
    
End Sub
    
Private Sub sckClient_Connect()
    Log "Connection Established"
End Sub
    
Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String, SplitData() As String, SplitRequest() As String
    Dim strTemp As String
    
    'On Error GoTo errClient
    
    'strData is our raw data, SplitData is the data separated by DATA_DELIMITER
    'The splitting of data prevents winsock from jamming strings together with one another
    
    sckClient.GetData strData
    LogRAW strData
    
    SplitData = Split(strData, DATA_DELIMITER)
    
    For i = 0 To UBound(SplitData) - 1 'Loop through all the data
        SplitRequest = Split(SplitData(i), "|") 'Split the sub-data using the pipe character (|)
        If HasAuthenticated = False Then 'We have an encrypted string, we must return its
        'decrypted version to be authenticated
        If chkSendKey.Value = vbChecked Then 'If we want to send the valid string back...
        
        SendData TDecrypt(SplitData(0)) & "|" 'Send the pipe char since it is our data delimiter
        'It is removed on the serverside so it does not interfere with our auth string
        HasAuthenticated = True
        Else 'Otherwise, send some custom text back!
        
        Do
            strTemp = InputBox("Enter some text to send back", "Enter an authentication string to send back.")
        Loop Until strTemp <> ""
        
        SendData strTemp
        HasAuthenticated = True
    End If
    
Else
    Select Case SplitRequest(0)
        Case "AUTHENTICATION"
            Select Case SplitRequest(1)
                Case "GRANTED"
                    Log "Authentication Accepted!"
                Case "DENIED"
                    Log "Authentication Invalid; Access Denied."
            End Select
        Case "REPLY"
            Log "Received Reply from server: " & SplitRequest(1)
    End Select
End If


Next i

errClient: 'Just exit the sub
End Sub
    
Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckClient_Close
    
End Sub
    
