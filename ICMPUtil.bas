Attribute VB_Name = "ICMPUtil"
' --------------------------------
' Module      : ICMPUtil
' Description : Ping Implementation in Excel VBA
' --------------------------------
' (see also)
' - http://support.microsoft.com/kb/154512/ja
' - http://support.microsoft.com/kb/300197/ja
'

Option Explicit

' Update here when you want to change the default parameter
Private Const ICMP_ECHO_TIMEOUT      As Long = 500
Private Const ICMP_ECHO_SENDDATASIZE As Long = 32
Private Const DEFAULT_SUCCESSPATTERN As String = "$S [$D] SIZE=$Bbytes TTL=$T RTT=$Rms ($A)"
Private Const DEFAULT_FAILUREPATTERN As String = "$S [$D] ($A)"

' Do not change
Private Const WSADESCRIPTION_LEN         As Long = 256
Private Const WSASYS_STATUS_LEN          As Long = 128
Private Const INADDR_NONE                As Long = &HFFFFFFFF
Private Const INADDR_ANY                 As Long = &H0
Private Const IP_SUCCESS                 As Long = 11000
Private Const IP_BUF_TOO_SMALL           As Long = 11001
Private Const IP_DEST_NET_UNREACHABLE    As Long = 11002
Private Const IP_DEST_HOST_UNREACHABLE   As Long = 11003
Private Const IP_DEST_PROT_UNREACHABLE   As Long = 11004
Private Const IP_DEST_PORT_UNREACHABLE   As Long = 11005
Private Const IP_NO_RESOURCES            As Long = 11006
Private Const IP_BAD_OPTION              As Long = 11007
Private Const IP_HW_ERROR                As Long = 11008
Private Const IP_PACKET_TOO_BIG          As Long = 11009
Private Const IP_REQ_TIMED_OUT           As Long = 11010
Private Const IP_BAD_REQ                 As Long = 11011
Private Const IP_BAD_ROUTE               As Long = 11012
Private Const IP_TTL_EXPIRED_TRANSIT     As Long = 11013
Private Const IP_TTL_EXPIRED_REASSEM     As Long = 11014
Private Const IP_PARAM_PROBLEM           As Long = 11015
Private Const IP_SOURCE_QUENCH           As Long = 11016
Private Const IP_OPTION_TOO_BIG          As Long = 11017
Private Const IP_BAD_DESTINATION         As Long = 11018
Private Const IP_GENERAL_FAILURE         As Long = 11050
Private Const ERROR_INSUFFICIENT_BUFFER  As Long = 122
Private Const ERROR_INVALID_PARAMETER    As Long = 87
Private Const ERROR_NOT_ENOUGH_MEMORY    As Long = 8
Private Const ERROR_NOT_SUPPORTED        As Long = 50
Private Const LOCALERR_SOCKETINIT        As Long = &HFFFFFF01
Private Const LOCALERR_DESTADDR_NONE     As Long = &HFFFFFF02
Private Const LOCALERR_DESTADDR_ANY      As Long = &HFFFFFF03
Private Const LOCALERR_GETHOSTBYNAME     As Long = &HFFFFFF04

' szDescription is [WSADESCRIPTION_LEN + 1]
' szSystemStatus is [WSASYS_STATUS_LEN + 1]
Private Type WSADATA
 wVersion                 As Integer
 wHighVersion             As Integer
 szDescription(0 To WSADESCRIPTION_LEN) As Byte
 szSystemStatus(0 To WSASYS_STATUS_LEN) As Byte
 iMaxSockets              As Integer
 iMaxUdpDg                As Integer
 lpVendorInfo             As Long
End Type

Private Type HOSTENT
  h_name      As Long
  h_aliases   As Long
  h_addrtype  As Integer
  h_length    As Integer
  h_addr_list As Long
End Type

Private Type IP_OPTION_INFORMATION
  TTL         As Byte
  Tos         As Byte
  Flags       As Byte
  OptionsSize As Byte
  OptionsData As Long
End Type

Private Type ICMP_ECHO_REPLY
  Address(0 To 3) As Byte
  Status        As Long
  RoundTripTime As Long
  Datasize      As Integer
  Reserved      As Integer
  Data          As Long
  Options       As IP_OPTION_INFORMATION
End Type

Private Type REPLY_BUFFER
  EchoReply                    As ICMP_ECHO_REPLY
  Data(ICMP_ECHO_SENDDATASIZE) As Byte
End Type


Private Declare Function WSAStartup Lib "ws2_32" ( _
  ByVal wVersionRequired As Integer, _
  ByRef lpWSADATA As WSADATA _
) As Long

Private Declare Function WSACleanup Lib "ws2_32" ( _
) As Long

Private Declare Function inet_addr Lib "ws2_32" ( _
  ByVal Address As String _
) As Long

Private Declare Function gethostbyname Lib "ws2_32" ( _
  ByVal Name As String _
) As Long

Private Declare Function IcmpCreateFile Lib "icmp" ( _
) As Long

Private Declare Function IcmpCloseHandle Lib "icmp" ( _
  ByVal IcmpHandle As Long _
) As Boolean

Private Declare Function IcmpSendEcho Lib "icmp" ( _
  ByVal IcmpHandle As Long, _
  ByVal DestinationAddress As Long, _
  ByVal RequestData As String, _
  ByVal RequestSize As Integer, _
  ByVal RequestOptions As String, _
  ByRef ReplyBuffer As REPLY_BUFFER, _
  ByVal ReplySize As Long, _
  ByVal Timeout As Long _
) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
  Destination As Any, _
  ByVal Source As Any, _
  ByVal Length As Long _
)

' Function  : Msg
' Parameter : Status = <ICMP_ECHO_REPLY.Status or LocalErrorValue>
' Return    : Translated message as String
Private Function Msg( _
  ByVal Status As Long _
) As String

  Select Case Status
    Case 0:                         Msg = "Success"
    Case IP_SUCCESS:                Msg = "Success"
    Case IP_BUF_TOO_SMALL:          Msg = "Buffer Too Small"
    Case IP_DEST_NET_UNREACHABLE:   Msg = "Destination Net Unreachable"
    Case IP_DEST_HOST_UNREACHABLE:  Msg = "Destination Host Unreachable"
    Case IP_DEST_PROT_UNREACHABLE:  Msg = "Destination Protocol Unreachable"
    Case IP_DEST_PORT_UNREACHABLE:  Msg = "Destination Port Unreachable"
    Case IP_NO_RESOURCES:           Msg = "No Resources"
    Case IP_BAD_OPTION:             Msg = "Bad Option"
    Case IP_HW_ERROR:               Msg = "Hardware Error"
    Case IP_PACKET_TOO_BIG:         Msg = "Packet Too Big"
    Case IP_REQ_TIMED_OUT:          Msg = "Request Timed Out"
    Case IP_BAD_REQ:                Msg = "Bad Request"
    Case IP_BAD_ROUTE:              Msg = "Bad Route"
    Case IP_TTL_EXPIRED_TRANSIT:    Msg = "TimeToLive Expired Transit"
    Case IP_TTL_EXPIRED_REASSEM:    Msg = "TimeToLive Expired Reassembly"
    Case IP_PARAM_PROBLEM:          Msg = "Parameter Problem"
    Case IP_SOURCE_QUENCH:          Msg = "Source Quench"
    Case IP_OPTION_TOO_BIG:         Msg = "Option Too Big"
    Case IP_BAD_DESTINATION:        Msg = "Bad Destination"
    Case IP_GENERAL_FAILURE:        Msg = "General Failure"
    Case ERROR_INSUFFICIENT_BUFFER: Msg = "Insufficient Buffer"
    Case ERROR_INVALID_PARAMETER:   Msg = "Invalid Parameter"
    Case ERROR_NOT_ENOUGH_MEMORY:   Msg = "Not Enough memory"
    Case ERROR_NOT_SUPPORTED:       Msg = "Not Supported"
    Case LOCALERR_SOCKETINIT:       Msg = "Socket Initialize Failure"
    Case LOCALERR_DESTADDR_NONE:    Msg = "Destination Address None"
    Case LOCALERR_DESTADDR_ANY:     Msg = "Destination Address Any"
    Case LOCALERR_GETHOSTBYNAME:    Msg = "GetHostByName Error"
    Case Else:                      Msg = "UnknownStatus " & Status
  End Select

End Function

Private Function SocketsInitialize() As Long
  Dim pWSADATA As WSADATA
  SocketsInitialize = WSAStartup(&H2, pWSADATA)
End Function

Private Function SocketsCleanup() As Long
  SocketsCleanup = WSACleanup()
End Function

' Function  : ExecGetHostByName
' Parameter : Name = <HostName>
' Return    : Address for the Target, in the form of Long representation
'             ErrorCode is 0 (AddressList is Null)
Private Function ExecGetHostByName( _
  Name As String _
) As Long

  Dim Result As Long
  Dim HostEntry As HOSTENT
  Dim pHostent As Long
  Dim pAddress As Long
    
  pHostent = gethostbyname(Name)
  If pHostent <> 0 Then
    RtlMoveMemory HostEntry, pHostent, LenB(HostEntry)
    RtlMoveMemory pAddress, HostEntry.h_addr_list, 4
    RtlMoveMemory Result, pAddress, HostEntry.h_length
  End If

  ExecGetHostByName = Result

End Function

' Function  : ExecIcmpSendEcho
' Parameter : DestinationAddress = <IPAddress as Long>
' Return    : The buffer of ICMP_ECHO_REPLY structure
'             When IcmpSendEcho fails , the function puts "Err().LastDllError"
'             in ICMP_ECHO_REPLY.Status, to returns DllErrorCode.
Private Function ExecIcmpSendEcho( _
  ByVal DestinationAddress As Long _
) As ICMP_ECHO_REPLY

  Dim IcmpHandle As Long
  Dim ReturnCode As Long
  Dim SendData(ICMP_ECHO_SENDDATASIZE) As Byte
  Dim ReplyBuffer As REPLY_BUFFER
    
    IcmpHandle = IcmpCreateFile()
    ReturnCode = IcmpSendEcho( _
      IcmpHandle, _
      DestinationAddress, _
      SendData, _
      ICMP_ECHO_SENDDATASIZE, _
      0, _
      ReplyBuffer, _
      LenB(ReplyBuffer.EchoReply) + UBound(ReplyBuffer.Data), _
      ICMP_ECHO_TIMEOUT)
    
    If ReturnCode = 0 And ReplyBuffer.EchoReply.Status = 0 Then
      ReplyBuffer.EchoReply.Status = Err().LastDllError
    End If
    
    IcmpCloseHandle (IcmpHandle)
    ExecIcmpSendEcho = ReplyBuffer.EchoReply

End Function

' Function  : BuildResponseString
' Parameter : Pattern = <Character string with parameter>
'             Address = <IPAddress as String>
'             StatusMessage = <Ping ResultMessage of String>
'             StatusCode = <Ping ResultCode of String>
'             RoundTripTime = <RTT as String>
'             Datasize = <Datesize as String>
'             TimeToLive = <TTL as String>
' Return    : Replaced character string
'
' This function rewrites the character string "Pattern" according to
' the parameter.
' There are the following in the parameter.
' $S -> StatusMessage
' $C -> StatusCode
' $A -> IPAddress
' $B -> Datasize
' $T -> TimeToLive
' $R -> RoundTripTime
' $D -> Return NOW()
' $$,$# -> $
'
' (sample)
' Pattern : "$S [$D] SIZE=$Bbytes TTL=$T RTT=$Rms ($A)"
' Result  : SUCCESS [2009/09/21 14:38:39] SIZE=32bytes TTL=128 RTT=0ms (127.0.0.1)
'
Private Function BuildResponseString( _
  Pattern As String, _
  Address As String, _
  StatusMessage As String, _
  StatusCode As Long, _
  RoundTripTime As Long, _
  Datasize As Long, _
  TimeToLive As Long _
) As String

  Pattern = Replace(Pattern, "$$", "$#")
  Pattern = Replace(Pattern, "$S", StatusMessage)
  Pattern = Replace(Pattern, "$C", StatusCode)
  Pattern = Replace(Pattern, "$A", Address)
  Pattern = Replace(Pattern, "$B", Datasize)
  Pattern = Replace(Pattern, "$T", TimeToLive)
  Pattern = Replace(Pattern, "$R", RoundTripTime)
  Pattern = Replace(Pattern, "$D", Now())
  Pattern = Replace(Pattern, "$#", "$")
  BuildResponseString = Pattern

End Function

' Function  : ExecPing
' Parameter : Target = <Target HostName or Target IPAddress>
'             SuccessPattern = <Character string used when succeeding>
'             FailurePattern = <Character string used when failing>
' Return    : Character string of Pattern was rewritten by BuildResponseString Function
Private Function ExecPing( _
  Target As String, _
  SuccessPattern As String, _
  FailurePattern As String _
) As String

  Dim ValidTarget As Boolean
  Dim DestinationAddress As Long
  Dim Pattern As String
  Dim Address As String
  Dim StatusMessage As String
  Dim StatusCode As Long
  Dim RoundTripTime As Long
  Dim Datasize As Long
  Dim TimeToLive As Long
  Dim EchoReply As ICMP_ECHO_REPLY
  
  Pattern = FailurePattern
  ValidTarget = False
  
  If SocketsInitialize() <> 0 Then
    ' Socket Initialize Failure
    StatusCode = LOCALERR_SOCKETINIT
  Else
    ' If the Target is IPAddress-String,
    DestinationAddress = inet_addr(Target)
    If DestinationAddress = INADDR_NONE Then
      ' If the Target is Hostname-String,
      DestinationAddress = ExecGetHostByName(Target)
      If DestinationAddress = 0 Then
        ' Lookup Failure
        StatusCode = LOCALERR_GETHOSTBYNAME
      ElseIf DestinationAddress = INADDR_NONE Then
        ' Invalid Address
        StatusCode = LOCALERR_DESTADDR_NONE
      Else
        ValidTarget = True
      End If
    ElseIf DestinationAddress = INADDR_ANY Then
      ' Invalid Address
      StatusCode = LOCALERR_DESTADDR_ANY
    Else
      ValidTarget = True
    End If
    
    If ValidTarget Then
    
      EchoReply = ExecIcmpSendEcho(DestinationAddress)
      
      Address = _
        EchoReply.Address(0) & "." & _
        EchoReply.Address(1) & "." & _
        EchoReply.Address(2) & "." & _
        EchoReply.Address(3)
      StatusCode = EchoReply.Status
      
      If EchoReply.Status = 0 Then
        Pattern = SuccessPattern
        RoundTripTime = EchoReply.RoundTripTime
        Datasize = EchoReply.Datasize
        TimeToLive = EchoReply.Options.TTL
      End If
      
    End If

    SocketsCleanup
  
  End If

  StatusMessage = Msg(StatusCode)
  
  ExecPing = BuildResponseString( _
    Pattern, _
    Address, _
    StatusMessage, _
    StatusCode, _
    RoundTripTime, _
    Datasize, _
    TimeToLive _
  )

End Function

' Function  : Ping
' Parameter : Target = <Target HostName or Target IPAddress as String>
'             SuccessPattern (Optional) = <Character string used when succeeding>
'             FailurePattern (Optional) = <Character string used when failing>
' Return    : Rewritten pattern
Public Function Ping( _
  ByVal Target As String, _
  Optional ByVal SuccessPattern As String = DEFAULT_SUCCESSPATTERN, _
  Optional ByVal FailurePattern As String = DEFAULT_FAILUREPATTERN _
) As String

  Ping = ExecPing(Target, SuccessPattern, FailurePattern)
 
End Function


