Attribute VB_Name = "mWSProtocol"
Option Explicit
Option Compare Text
'==============================================================
'By:       悠悠然
'QQ:       2860898817
'E-mail:   ur1986@foxmail.com
'服务端及客户端Demo放Q群文件共享:369088586
'项目创建时间: 2015.04.06
'最后改动时间: 2017.12.12
'==============================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Enum OpcodeType
    opContin = 0    '连续消息片断
    opText = 1      '文本消息片断
    opBinary = 2    '二进制消息片断
                    '3 - 7 非控制帧保留
    opClose = 8     '连接关闭
    opPing = 9      '心跳检查的ping
    opPong = 10     '心跳检查的pong
                    '11-15 控制帧保留
End Enum
Public Type DataFrame
    FIN As Boolean      '0表示不是当前消息的最后一帧，后面还有消息,1表示这是当前消息的最后一帧；
    RSV1 As Boolean     '1位，若没有自定义协议,必须为0,否则必须断开.
    RSV2 As Boolean     '1位，若没有自定义协议,必须为0,否则必须断开.
    RSV3 As Boolean     '1位，若没有自定义协议,必须为0,否则必须断开.
    Opcode As OpcodeType    '4位操作码，定义有效负载数据，如果收到了一个未知的操作码，连接必须断开.
    MASK As Boolean     '1位，定义传输的数据是否有加掩码,如果有掩码则存放在MaskingKey
    MaskingKey(3) As Byte   '32位的掩码
    Payloadlen As Long  '传输数据的长度
    DataOffset As Long  '数据源起始位
End Type

'==============================================================
'握手部分,只有一个开放调用函数 Handshake(requestHeader As String) As Byte()
'==============================================================
Private Const MagicKey = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"
Private Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Public Function Handshake(requestHeader As String) As Byte()
    Dim clientKey As String
    clientKey = getHeaderValue(requestHeader, "Sec-WebSocket-Key:")
    Dim AcceptKey As String
    AcceptKey = getAcceptKey(clientKey)
    Dim response As String
    response = "HTTP/1.1 101 Web Socket Protocol Handshake" & vbCrLf
    response = response & "Upgrade: WebSocket" & vbCrLf
    response = response & "Connection: Upgrade" & vbCrLf
    response = response & "Sec-WebSocket-Accept: " & AcceptKey & vbCrLf
    response = response & "WebSocket-Origin: " & getHeaderValue(requestHeader, "Sec-WebSocket-Origin:") & vbCrLf
    response = response & "WebSocket-Location: " & getHeaderValue(requestHeader, "Host:") & vbCrLf
    'response = response & "WebSocket-Server: VB.Shunshisan" & vbCrLf
    response = response & vbCrLf
    'Debug.Print response
    Handshake = StrConv(response, vbFromUnicode)
End Function
Private Function getHeaderValue(str As String, pname As String) As String
    Dim i As Long, j As Long
    i = InStr(str, pname)
    If i > 0 Then
        j = InStr(i, str, vbCrLf)
        If j > 0 Then
            i = i + Len(pname)
            getHeaderValue = Trim(Mid(str, i, j - i))
        End If
    End If
End Function
Private Function getAcceptKey(key As String) As String
    Dim b() As Byte
    b = mSHA1.SHA1(StrConv(key & "258EAFA5-E914-47DA-95CA-C5AB0DC85B11", vbFromUnicode))
    getAcceptKey = EnBase64(b)
End Function
Private Function EnBase64(str() As Byte) As String
    On Error GoTo over
    Dim buf() As Byte, length As Long, mods As Long
    mods = (UBound(str) + 1) Mod 3
    length = UBound(str) + 1 - mods
    ReDim buf(length / 3 * 4 + IIf(mods <> 0, 4, 0) - 1)
    Dim i As Long
    For i = 0 To length - 1 Step 3
        buf(i / 3 * 4) = (str(i) And &HFC) / &H4
        buf(i / 3 * 4 + 1) = (str(i) And &H3) * &H10 + (str(i + 1) And &HF0) / &H10
        buf(i / 3 * 4 + 2) = (str(i + 1) And &HF) * &H4 + (str(i + 2) And &HC0) / &H40
        buf(i / 3 * 4 + 3) = str(i + 2) And &H3F
    Next
    If mods = 1 Then
        buf(length / 3 * 4) = (str(length) And &HFC) / &H4
        buf(length / 3 * 4 + 1) = (str(length) And &H3) * &H10
        buf(length / 3 * 4 + 2) = 64
        buf(length / 3 * 4 + 3) = 64
    ElseIf mods = 2 Then
        buf(length / 3 * 4) = (str(length) And &HFC) / &H4
        buf(length / 3 * 4 + 1) = (str(length) And &H3) * &H10 + (str(length + 1) And &HF0) / &H10
        buf(length / 3 * 4 + 2) = (str(length + 1) And &HF) * &H4
        buf(length / 3 * 4 + 3) = 64
    End If
    For i = 0 To UBound(buf)
        EnBase64 = EnBase64 + Mid(B64_CHAR_DICT, buf(i) + 1, 1)
    Next
over:
End Function
'==============================================================
'数据帧解析,返回帧结构
'==============================================================
Public Function AnalyzeHeader(byt() As Byte) As DataFrame
    Dim DF As DataFrame
    Dim l(3) As Byte
    DF.FIN = IIf((byt(0) And &H80) = &H80, True, False)
    DF.RSV1 = IIf((byt(0) And &H40) = &H40, True, False)
    DF.RSV2 = IIf((byt(0) And &H20) = &H20, True, False)
    DF.RSV3 = IIf((byt(0) And &H10) = &H10, True, False)
    DF.Opcode = byt(0) And &H7F
    DF.MASK = IIf((byt(1) And &H80) = &H80, True, False)
    Dim plen As Byte
    plen = byt(1) And &H7F
    If plen < 126 Then
        DF.Payloadlen = plen
        If DF.MASK Then
            CopyMemory DF.MaskingKey(0), byt(2), 4
            DF.DataOffset = 6
        Else
            DF.DataOffset = 2
        End If
    ElseIf plen = 126 Then
        l(0) = byt(3)
        l(1) = byt(2)
        CopyMemory DF.Payloadlen, l(0), 4
        If DF.MASK Then
            CopyMemory DF.MaskingKey(0), byt(4), 4
            DF.DataOffset = 8
        Else
            DF.DataOffset = 4
        End If
    ElseIf plen = 127 Then
        '这部分没有什么意义就不写了,因为VB没有64位的整型可供使用
        '所以对长度设定为-1,自己再判断
        If byt(2) <> 0 Or byt(3) <> 0 Or byt(4) <> 0 Or byt(5) <> 0 Then
            '超过32位
            DF.Payloadlen = -1
        Else
            l(0) = byt(9)
            l(1) = byt(8)
            l(2) = byt(7)
            l(3) = byt(6)
            CopyMemory DF.Payloadlen, l(0), 4
            If DF.Payloadlen <= 0 Then
                '超过有符号
                DF.Payloadlen = -1
            Else
                If DF.MASK Then
                    CopyMemory DF.MaskingKey(0), byt(10), 4
                    DF.DataOffset = 14
                Else
                    DF.DataOffset = 10
                End If
            End If
        End If
    End If
    AnalyzeHeader = DF
End Function
'==============================================================
'接收的数据处理,有掩码就反掩码
'PickDataV  方法是出于性能的考虑,用于有时数据只是为了接收,做一些逻辑判断,并不需要对数据块进行单独提炼
'PickData   不赘述了...
'==============================================================
Public Sub PickDataV(byt() As Byte, dataType As DataFrame)
    Dim lenLimit As Long
    lenLimit = dataType.DataOffset + dataType.Payloadlen - 1
    If dataType.MASK And lenLimit <= UBound(byt) Then
        Dim i As Long, j As Long
        For i = dataType.DataOffset To lenLimit
            byt(i) = byt(i) Xor dataType.MaskingKey(j)
            j = j + 1
            If j = 4 Then j = 0
        Next i
    End If
End Sub
Public Function PickData(byt() As Byte, dataType As DataFrame) As Byte()
    Dim b() As Byte
    PickDataV byt, dataType
    ReDim b(dataType.Payloadlen - 1)
    CopyMemory b(0), byt(dataType.DataOffset), dataType.Payloadlen
    PickData = b
End Function

'==============================================================
'发送的数据处理,该部分未联网测试,使用下面的方式测试验证
'Private Sub Command1_Click()
'    Dim str As String, b() As Byte, bs() As Byte
'    Dim DF As DataFrame
'    str = "abc123"
'    Showlog "组装前数据:" & str
'    b = mWSProtocol.PackMaskString(str):    Showlog "掩码后字节:" & BytesToHex(b)
'    DF = mWSProtocol.AnalyzeHeader(b):      Showlog "结构体偏移:" & DF.DataOffset & "  长度:" & DF.Payloadlen
'    bs = mWSProtocol.PickData(b, DF):       Showlog "还原后字节:" & BytesToHex(bs)
'    Showlog "还原后数据:" & StrConv(bs, vbUnicode)
'End Sub
'==============================================================
'无掩码数据的组装,用于服务端向客户端发送
'--------------------------------------------------------------
Public Function PackString(str As String, Optional dwOpcode As OpcodeType = opText) As Byte()
    Dim b() As Byte
    b = mUTF8.Encoding(str) '默认UTF8
    PackString = PackData(b, dwOpcode)
End Function
Public Function PackData(data() As Byte, Optional dwOpcode As OpcodeType = opText) As Byte()
    Dim length As Long
    Dim byt() As Byte
    length = UBound(data) + 1
    
    If length < 126 Then
        ReDim byt(length + 1)
        byt(1) = CByte(length)
        CopyMemory byt(2), data(0), length
    ElseIf length <= 65535 Then
        ReDim byt(length + 3)
        Dim l(1) As Byte
        byt(1) = &H7E
        CopyMemory l(0), length, 2
        byt(2) = l(1)
        byt(3) = l(0)
        CopyMemory byt(4), data(0), length
    'ElseIf length <= 999999999999999# Then
        '这么长不处理了...
        'VB6也没有这么大的整型
        '有需要就根据上面调整来写吧
    End If
    '------------------------------
    '关于下面的 byt(0) = &H80 Or dwOpcode 中，&H80 对应的是 DataFrame 结构中的FIN + RSV1 + RSV2 + RSV3
    'FIN 的中文解释是：指示这个是消息的最后片段，第一个片段可能也是最后的片段。
    '这里我不是很理解，可能是自定义分包用到吧，但貌似分包应该不是自己可控的。
    '------------------------------
    byt(0) = &H80 Or dwOpcode
    PackData = byt
End Function
'--------------------------------------------------------------
'有掩码数据的组装,用于替代客户端向服务端发送
'--------------------------------------------------------------
Public Function PackMaskString(str As String, Optional dwOpcode As OpcodeType = opText) As Byte()
    Dim b() As Byte
    b = mUTF8.Encoding(str) '默认UTF8
    PackMaskString = PackMaskData(b, dwOpcode)
End Function
Public Function PackMaskData(data() As Byte, Optional dwOpcode As OpcodeType = opText) As Byte()
    '对源数据做掩码处理
    Dim mKey(3) As Byte
    mKey(0) = 108: mKey(1) = 188: mKey(2) = 98: mKey(3) = 208 '掩码,你也可以自己定义
    Dim i As Long, j As Long
    For i = 0 To UBound(data)
        data(i) = data(i) Xor mKey(j)
        j = j + 1
        If j = 4 Then j = 0
    Next i
    '包装,和上面的无掩码包装PackData()大体相同
    Dim length As Long
    Dim byt() As Byte
    length = UBound(data) + 1
    If length < 126 Then
        ReDim byt(length + 5)
        byt(0) = &H80 Or dwOpcode '帧类型
        byt(1) = (CByte(length) Or &H80)
        CopyMemory byt(2), mKey(0), 4
        CopyMemory byt(6), data(0), length
    ElseIf length <= 65535 Then
        ReDim byt(length + 7)
        Dim l(1) As Byte
        byt(0) = &H80 Or dwOpcode '&H81 '同上注意
        byt(1) = &HFE '固定 掩码位+126
        CopyMemory l(0), length, 2
        byt(2) = l(1)
        byt(3) = l(0)
        CopyMemory byt(4), mKey(0), 4
        CopyMemory byt(8), data(0), length
    'ElseIf length <= 999999999999999# Then
        '这么长不处理了...有需要就根据上面调整来写吧
    End If
    PackMaskData = byt
End Function
'==============================================================
'控制帧相关,Ping、Pong、Close 用于服务端向客户端发送未经掩码的信号
'我用的0长度,其实是可以包含数据的,但是附带数据客户端处理又麻烦了

'使用举例: Winsock1.SendData mWSProtocol.PongFrame()

'* 如果有附带信息的需求,也可以用PackString或PackData,可选参数指定OpcodeType
'* 协议规定,附带的字符串消息,依然按照掩码规则,客户端发送掩码,服务端发送不掩码
'==============================================================
Public Function PingFrame(Optional msg As String = "", Optional UseMask As Boolean = False) As Byte()
    Dim b(1) As Byte
    b(0) = &H89
    b(1) = &H0
    PingFrame = b
    '发送一个包含"Hello"的Ping信号: 0x89 0x05 0x48 0x65 0x6c 0x6c 0x6f
End Function
Public Function PongFrame(Optional msg As String = "", Optional UseMask As Boolean = False) As Byte()
    Dim b(1) As Byte
    b(0) = &H8A
    b(1) = &H0
    PongFrame = b
    '发送一个包含"Hello"的Pong信号: 0x8A 0x05 0x48 0x65 0x6c 0x6c 0x6f
End Function
Public Function CloseFrame(Optional msg As String = "", Optional UseMask As Boolean = False) As Byte()
    Dim b(1) As Byte
    b(0) = &H88
    b(1) = &H0
    CloseFrame = b
    '发送一个包含"Close"的Pong信号: 0x8A 0x05 0x43 0x6c 0x6f 0x73 0x65
End Function
