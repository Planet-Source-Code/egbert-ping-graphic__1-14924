Attribute VB_Name = "PingPing"
Option Explicit
'*******************************************************************
'   PingVB
'
'   This application implements a TCP/IP network ping
'   using the ICMP.DLL provided as part of Windows 95 and
'   Windows NT.
'
'*******************************************************************

'   option information for network ping, we don't implement these here as this is
'   a simple sample (simon says).
Private Type ip_option_information
    TTl             As Byte     'Time To Live
    Tos             As Byte     'Type Of Service
    Flags           As Byte     'IP header flags
    OptionsSize     As Byte     'Size in bytes of options data
    OptionsData     As Long     'Pointer to options data
End Type

'   structure that is returned from the ping to give status and error information
Private Type icmp_echo_reply
    Address         As Long             'Replying address
    Status          As Long             'Reply IP_STATUS, values as defined above
    RoundTripTime   As Long             'RTT in milliseconds
    DataSize        As Integer          'Reply data size in bytes
    Reserved        As Integer          'Reserved for system use
    DataPointer     As Long             'Pointer to the reply data
    Options         As ip_option_information    'Reply options
    Data            As String * 250     'Reply data which should be a copy of the string sent, NULL terminated
                                        ' this field length should be large enough to contain the string sent
End Type

'   declares for function to be used from icmp.dll
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, _
                                                    ByVal DestinationAddress As Long, _
                                                    ByVal RequestData As String, _
                                                    ByVal RequestSize As Integer, _
                                                    RequestOptions As ip_option_information, _
                                                    ReplyBuffer As icmp_echo_reply, _
                                                    ByVal ReplySize As Long, _
                                                    ByVal TimeOut As Long) As Long

Private Const WSADESCRIPTION_LEN = 256
Private Const WSASYSSTATUS_LEN = 256
Private Const WSADESCRIPTION_LEN_1 = WSADESCRIPTION_LEN + 1
Private Const WSASYSSTATUS_LEN_1 = WSASYSSTATUS_LEN + 1
Private Const SOCKET_ERROR = -1

Type PingReturn
Discriptionn As String
Ping As Long
Errors As Boolean
End Type

Dim CurIp As Long
Dim CurIpDes As String

Private Type tagWSAData
        wVersion            As Integer
        wHighVersion        As Integer
        szDescription       As String * WSADESCRIPTION_LEN_1
        szSystemStatus      As String * WSASYSSTATUS_LEN_1
        iMaxSockets         As Integer
        iMaxUdpDg           As Integer
        lpVendorInfo        As String * 200
        End Type
Private Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequested As Integer, lpWSAData As tagWSAData) As Integer
Private Declare Function WSACleanup Lib "wsock32" () As Integer

'   btnPing
'
'   This routine is called when the button is clicked. The Ip address to be pinged
'   is taken from the text box and converted to a long value for the Icmp call
'
Function Pinger(TTl As Long, TimeOut As Long) As PingReturn
    Dim hFile       As Long             ' handle for the icmp port opened
    Dim lRet        As Long             ' hold return values as required
    Dim lIPAddress  As Long
    Dim strMessage  As String
    Dim pOptions    As ip_option_information
    Dim pReturn     As icmp_echo_reply
    Dim iVal        As Integer
    Dim lPingRet    As Long
    Dim pWsaData    As tagWSAData
    
    strMessage = "Echo this string of data"
    
    iVal = WSAStartup(&H101, pWsaData)
    
    '   convert the IP address to a long, lIPAddress will be zero
    '   if the function failed. Normally you wouldn't ping if the address
    '   was no good to start with but we don't mind seeing bad return status
    '   as that is what samples are all about
    lIPAddress = CurIp
    
    '   open up a file handle for doing the ping
    hFile = IcmpCreateFile()
    
    '   set the TTL from the text box, try values of 1 to 255
    pOptions.TTl = TTl
    
    '   Call the function that actually does the ping. It is a blocking call so we
    '   don't get control back until it completes.
    lRet = IcmpSendEcho(hFile, lIPAddress, strMessage, Len(strMessage), pOptions, pReturn, Len(pReturn), TimeOut)

    If lRet = 0 Then
        ' the ping failed for some reason, hopefully the error is in the return buffer
    Pinger.Discriptionn = "Faild to ping : ''" & CurIpDes & "'' error code : " & ErrorCodeToDes(pReturn.Status)
    Pinger.Errors = True
    Pinger.Ping = TimeOut
    Else
        ' the ping succeeded, .Status will be 0, .RoundTripTime is the time in ms for
        '   the ping to complete, .Data is the data returned (NULL terminated), .Address
        '   is the Ip address that actually replied, .DataSize is the size of the string in
        '   .Data
        If pReturn.Status <> 0 Then
            Pinger.Discriptionn = "Faild to ping : ''" & CurIpDes & "'' error description : " & ErrorCodeToDes(pReturn.Status)
            Pinger.Errors = True
            Pinger.Ping = TimeOut
        Else
            Pinger.Discriptionn = "Success to ping : ''" & CurIpDes & "'' completion time is " & pReturn.RoundTripTime & "ms."
            Pinger.Ping = pReturn.RoundTripTime
            End If
            
        If pReturn.RoundTripTime > TimeOut Then
        Pinger.Discriptionn = "Faild to ping : ''" & CurIpDes & "'' error code : " & ErrorCodeToDes(11010)
        Pinger.Errors = True
        Pinger.Ping = TimeOut
        End If
    End If
                        
    '   close the file handle that was used
    lRet = IcmpCloseHandle(hFile)
    
    iVal = WSACleanup()
    
End Function

'
'   ConvertIPAddressToLong
'
'   Converts a dotted IP address (eg: "123.234.2.45") to a long
'   integer for use in sending a ping. This routine converts
'   the string as required by an Intel system.
'
'   Essentially we take the 4 numbers, flip them around and make
'   a long by shifting all the parts into the correct byte. We
'   do it here by making a hex string and converting it to a long.
'   Not pretty but it works (most of the time<g>).
'
'   When we get in "a.b.c.d" what we want out is Val(&Hddccbbaa).
'

Function SetIp(strAddress As String)

    Dim strTemp             As String
    Dim lAddress            As Long
    Dim iValCount           As Integer
    Dim lDotValues(1 To 4)  As String
    
    ' set up the initial storage and counter
    strTemp = strAddress
    iValCount = 0
    
    ' keep going while we still have dots in the string
    While InStr(strTemp, ".") > 0
        iValCount = iValCount + 1   ' count the number
        lDotValues(iValCount) = Mid(strTemp, 1, InStr(strTemp, ".") - 1)    ' pick it off and convert it
        strTemp = Mid(strTemp, InStr(strTemp, ".") + 1) ' chop off the number and the dot
        Wend
        
    ' the string only has the last number in it now
    iValCount = iValCount + 1
    lDotValues(iValCount) = strTemp
    
    ' if we didn't get four pieces then the IP address is no good
    If iValCount <> 4 Then
        CurIp = 0
        Exit Function
        End If
        
    '   take the four value, hex them, pad to 2 digits, make a hex
    '   string and then convert the whole mess to a long for returning
    lAddress = Val("&H" & Right("00" & Hex(lDotValues(4)), 2) & _
                Right("00" & Hex(lDotValues(3)), 2) & _
                Right("00" & Hex(lDotValues(2)), 2) & _
                Right("00" & Hex(lDotValues(1)), 2))
                
    '   set the return value
    CurIp = lAddress
    CurIpDes = strAddress
End Function

Function ErrorCodeToDes(Code As Long) As String
Dim Ret As String
Select Case Code
Case 0
Ret = "Success"
Case 11001
Ret = "Buffer to small"
Case 11000
Ret = "Status base"
Case 11002
Ret = "Destination network unreachable"
Case 11003
Ret = "Destination host unreachable"
Case 11004
Ret = "Destination protocol unreachable"
Case 11005
Ret = "Destination port unreachable"
Case 11006
Ret = "No resources"
Case 11007
Ret = "Bad option"
Case 11008
Ret = "hw error"
Case 11009
Ret = "Packet to big"
Case 11010
Ret = "Requested IP timed out"
Case 11011
Ret = "Bad request"
Case 11012
Ret = "Bad route"
Case 11013
Ret = "Time to live expired transmition"
Case 11014
Ret = "Time ti live expired reassem"
Case 11015
Ret = "Param problems"
Case 11016
Ret = "Source quench"
Case 11017
Ret = "Option too big"
Case 11018
Ret = "Bad destination"
Case 11019
Ret = "Address deleted"
Case 11020
Ret = "Spec mtu changed"
Case 11021
Ret = "Mtu changed"
Case 11022
Ret = "Unloaded"
Case 11023
Ret = "Address added"
Case 11050
Ret = "General failure"
Case 11255
Ret = "Ip pending"
Case Else
Ret = "Unindentified"
End Select
ErrorCodeToDes = Ret
End Function
