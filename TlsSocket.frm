VERSION 5.00
Begin VB.Form TlsSocket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SvTLS Wss By: Maatooh"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SendMsg 
      Caption         =   "Send Msg"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Wmsg 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chat"
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3375
      Begin VB.TextBox Msgtxt 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton SvOn 
      Caption         =   "Listen"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "TlsSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=========================================================================
' TlsSocketWss Project by Maatooh1@gmail.com
' Thanks to wqweto@gmail.com and ur1986@foxmail.com without their participation this project could not be possible.
'=========================================================================
Const PEM_FILES As String = "certificate.pem|ca_bundle.pem|private.pem"
Const PEM_FILE As String = "certificateC.pem"
Private WithEvents Serversck    As cTlsSocket
Attribute Serversck.VB_VarHelpID = -1
Private LastID                  As Long
Private Requests                As Collection
Private Const SvAcceptProtocol As String = "HTTP/1.1 101 Switching Protocols" & vbCrLf & _
"Upgrade: websocket" & vbCrLf & _
"Connection: Upgrade" & vbCrLf & _
"Sec-WebSocket-Accept: "

Private Sub Form_Load()
    Set Requests = New Collection
End Sub

Private Sub Msgtxt_Change()
Msgtxt.SelStart = Len(Msgtxt)
End Sub

Private Sub SendMsg_Click()
Dim D() As Byte
Dim Client As cClientRequest
If Not Wmsg = "" Then
    Debug.Print "SEND: " & Wmsg
    D = mWSProtocol.PackString(Wmsg) 'packaged for ws
    For Each Client In Requests
    Client.Socket.SendArray D
    Next
    Msgtxt = Msgtxt & "Server: " & Wmsg & vbCrLf
    Wmsg = ""
End If
End Sub

Private Sub SvOn_Click()
    Set Serversck = New cTlsSocket
    'Use self-signed certificate, before opening from the browser in the url: ports and accept the non-secure connection, then go back to test wss.
    'If Not Serversck.InitServerTls() Then
    If Not Serversck.InitServerTls(App.Path & "\" & PEM_FILES) Then
        GoTo QH
    End If
    If Not Serversck.Create(SocketPort:=5880, SocketAddress:="retromyths.com") Then
        GoTo QH
    End If
    If Not Serversck.Listen() Then
        GoTo QH
    End If
QH:
End Sub

Private Sub Serversck_OnAccept()
    Dim oSocket As cTlsSocket
    Dim Oclient As cClientRequest
    
    If Not Serversck.Accept(oSocket, UseTls:=True) Then
        GoTo QH
    End If
    Debug.Print "New User"
    Set Oclient = New cClientRequest
    LastID = LastID + 1
    Oclient.ID = LastID
    Set Oclient.Socket = oSocket
    Set Oclient.Callback = Me
    Oclient.HasHandshake = False
    Requests.Add Oclient, "#" & Oclient.ID
QH:
End Sub

Public Sub ClientOnReceive(Client As cClientRequest)
    Dim Svdata() As Byte
    Dim buff() As Byte
    
    If Not Client.Socket.ReceiveArray(Svdata) Then
        GoTo QH
    End If

    If Client.HasHandshake = False Then
        Dim headers As String
        headers = StrConv(Svdata, vbUnicode)
        buff = mWSProtocol.Handshake(headers)
        Client.HasHandshake = True
        Client.Socket.SendArray buff
        Exit Sub
    End If
    
    Dim DF As DataFrame
    Dim str As String
    DF = mWSProtocol.AnalyzeHeader(Svdata)
    buff = mWSProtocol.PickData(Svdata, DF)  '获取反掩码后的数据
    str = mUTF8.ToUnicodeString(buff)     '字节组转字符串
    Debug.Print "RECV: " & str
    Msgtxt = Msgtxt & "ClientID " & Client.ID & ": " & str & vbCrLf
    'Debug.Print Client.ID, StrConv(Svdata, vbUnicode)
'----Websocket
'If Not InStr(DataUserString, "GET / HTTP/1.1") = 0 Then
'SvKey = Mid(DataUserString, InStr(DataUserString, "Sec-WebSocket-Key: ") + 19, InStr(DataUserString, "Sec-WebSocket-Extensions:") - InStr(DataUserString, "Sec-WebSocket-Key: ") - 21) & "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"
'SvkeyB() = StrConv(SvKey, vbFromUnicode)
'Client.Socket.SendText SvAcceptProtocol & HexConv.Hex2Base64(HexConv.HexString(StrConv(SHA1.SHA1(SvkeyB), vbUnicode))) & vbCrLf & vbCrLf
'Else
'DF = EnConv.AnalyzeHeader(Svdata)
'Msgtxt = Msgtxt & "ClientID " & Client.ID & ": " & StrConv(EnConv.PickData(Svdata, DF), vbUnicode) & vbCrLf
'End If
'----
QH:
End Sub

Public Sub ClientOnClose(Client As cClientRequest)
    Requests.Remove "#" & Client.ID
    Debug.Print Client.ID, "Disconnected"
End Sub

Public Sub ClientOnSend(Client As cClientRequest)
     '
End Sub
