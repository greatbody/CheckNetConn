Attribute VB_Name = "Module1"
'检查网络通不通
Option Explicit

Private Const IP_SUCCESS As Long = 0
Private Const IP_STATUS_BASE As Long = 11000
Private Const IP_BUF_TOO_SMALL As Long = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE As Long = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE As Long = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE As Long = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE As Long = (11000 + 5)
Private Const IP_NO_RESOURCES As Long = (11000 + 6)
Private Const IP_BAD_OPTION As Long = (11000 + 7)
Private Const IP_HW_ERROR As Long = (11000 + 8)
Private Const IP_PACKET_TOO_BIG As Long = (11000 + 9)
Private Const IP_REQ_TIMED_OUT As Long = (11000 + 10)
Private Const IP_BAD_REQ As Long = (11000 + 11)
Private Const IP_BAD_ROUTE As Long = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT As Long = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM As Long = (11000 + 14)
Private Const IP_PARAM_PROBLEM As Long = (11000 + 15)
Private Const IP_SOURCE_QUENCH As Long = (11000 + 16)
Private Const IP_OPTION_TOO_BIG As Long = (11000 + 17)
Private Const IP_BAD_DESTINATION As Long = (11000 + 18)
Private Const IP_ADDR_DELETED As Long = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE As Long = (11000 + 20)
Private Const IP_MTU_CHANGE As Long = (11000 + 21)
Private Const IP_UNLOAD As Long = (11000 + 22)
Private Const IP_ADDR_ADDED As Long = (11000 + 23)
Private Const IP_GENERAL_FAILURE As Long = (11000 + 50)
Private Const MAX_IP_STATUS As Long = (11000 + 50)
Private Const IP_PENDING As Long = (11000 + 255)
Private Const PING_TIMEOUT As Long = 500
Private Const WS_VERSION_REQD As Long = &H101
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const SOCKET_ERROR As Long = -1
Private Const INADDR_NONE As Long = &HFFFFFFFF
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128

Public PingTime As Long
Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Private Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Long
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function inet_addr Lib "wsock32" (ByVal s As String) As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
'Private Declare Function WSAGetLastError Lib "wsock32" () As Long
'Private Declare Function gethostname Lib "wsock32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
'Private Declare Function gethostbyname Lib "wsock32" (ByVal szHost As String) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nbytes As Long)
   
Private Function GetStatusCode(status As Long) As String
   On Error GoTo ErrLine
   Dim Msg As String
   GetStatusCode = ""
   Select Case status
            Case IP_SUCCESS:               Msg = "ip success"
            Case INADDR_NONE:              Msg = "inet_addr: bad IP format"
            Case IP_BUF_TOO_SMALL:         Msg = "ip buf too_small"
            Case IP_DEST_NET_UNREACHABLE:  Msg = "ip dest net unreachable"
            Case IP_DEST_HOST_UNREACHABLE: Msg = "ip dest host unreachable"
            Case IP_DEST_PROT_UNREACHABLE: Msg = "ip dest port unreachable"
            Case IP_DEST_PORT_UNREACHABLE: Msg = "ip dest port unreachable"
            Case IP_NO_RESOURCES:          Msg = "ip no resources"
            Case IP_BAD_OPTION:            Msg = "ip bad option"
            Case IP_HW_ERROR:              Msg = "ip hw_error"
            Case IP_PACKET_TOO_BIG:        Msg = "ip packet too_big"
            Case IP_REQ_TIMED_OUT:         Msg = "ip req timed out"
            Case IP_BAD_REQ:               Msg = "ip bad req"
            Case IP_BAD_ROUTE:             Msg = "ip bad route"
            Case IP_TTL_EXPIRED_TRANSIT:   Msg = "ip ttl expired transit"
            Case IP_TTL_EXPIRED_REASSEM:   Msg = "ip ttl expired reassem"
            Case IP_PARAM_PROBLEM:         Msg = "ip param_problem"
            Case IP_SOURCE_QUENCH:         Msg = "ip source quench"
            Case IP_OPTION_TOO_BIG:        Msg = "ip option too_big"
            Case IP_BAD_DESTINATION:       Msg = "ip bad destination"
            Case IP_ADDR_DELETED:          Msg = "ip addr deleted"
            Case IP_SPEC_MTU_CHANGE:       Msg = "ip spec mtu change"
            Case IP_MTU_CHANGE:            Msg = "ip mtu_change"
            Case IP_UNLOAD:                Msg = "ip unload"
            Case IP_ADDR_ADDED:            Msg = "ip addr added"
            Case IP_GENERAL_FAILURE:       Msg = "ip general failure"
            Case IP_PENDING:               Msg = "ip pending"
            Case PING_TIMEOUT:             Msg = "ping timeout"
            Case Else:                     Msg = "unknown msg returned"
   End Select
   GetStatusCode = Msg
   Exit Function
ErrLine:
End Function

Private Function Ping(sAddress As String, sDataToSend As String, ECHO As ICMP_ECHO_REPLY) As Long
   On Error GoTo ErrLine
   Dim hPort As Long
   Dim dwAddress As Long
   dwAddress = inet_addr(sAddress)
   If dwAddress <> INADDR_NONE Then
      hPort = IcmpCreateFile()
      If hPort Then
         Call IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, ECHO, Len(ECHO), PING_TIMEOUT)
         Ping = ECHO.status
         Call IcmpCloseHandle(hPort)
      End If
   Else
      Ping = INADDR_NONE
   End If
   Exit Function
ErrLine:
   Ping = INADDR_NONE
End Function

Public Function PingIP(ByVal szIp As String) As Boolean
On Error GoTo ErrLine
Dim WSAD As WSADATA
Dim ECHO As ICMP_ECHO_REPLY
Dim ret As Long
'Delay 150
PingIP = False
PingTime = Empty
If WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS Then
   ret = Ping(Trim(szIp), "tanaya", ECHO)
   PingTime = ECHO.RoundTripTime
   If InStr(1, GetStatusCode(ret), "success") <> 0 Then
      WSACleanup
      PingIP = True
      PingTime = ECHO.RoundTripTime
      Exit Function
    End If
End If
Exit Function
ErrLine:
End Function


