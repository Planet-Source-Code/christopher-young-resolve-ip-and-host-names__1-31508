Attribute VB_Name = "mHost"

Public Const MIN_SOCKETS_REQD As Long = 1
Public Const WS_VERSION_REQD As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const SOCKET_ERROR As Long = -1
Public Const WSADESCRIPTION_LEN = 257
Public Const WSASYS_STATUS_LEN = 129
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type
Declare Function WSACleanup Lib "WSOCK32" () As Long
Declare Function WSAStartup Lib "WSOCK32" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
Declare Function gethostbyaddr Lib "wsock32.dll" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)


Public Function GetHostName(ByVal Address As Long) As String
    Dim lLength As Long, lRet As Long
    If Not SocketsInitialize() Then Exit Function
    lRet = gethostbyaddr(Address, 4, AF_INET)
    If lRet <> 0 Then
        CopyMemory lRet, ByVal lRet, 4
        lLength = lstrlenA(lRet)
        If lLength > 0 Then
            GetHostName = Space$(lLength)
            CopyMemory ByVal GetHostName, ByVal lRet, lLength
        End If
    Else
        GetHostName = "No Host Found"
    End If
    SocketsCleanup
End Function
Public Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function
Public Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function
Public Sub SocketsCleanup()
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
End Sub
Public Function SocketsInitialize() As Boolean
    Dim WSAD As WSAData
    Dim sLoByte As String
    Dim sHiByte As String
    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
        MsgBox "The 32-bit Windows Socket is not responding."
        SocketsInitialize = False
        Exit Function
    End If
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & CStr(MIN_SOCKETS_REQD) & " supported sockets."
        SocketsInitialize = False
        Exit Function
    End If
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        sHiByte = CStr(HiByte(WSAD.wVersion))
        sLoByte = CStr(LoByte(WSAD.wVersion))
        MsgBox "Sockets version " & sLoByte & "." & sHiByte & " is not supported by 32-bit Windows Sockets."
        SocketsInitialize = False
        Exit Function
    End If
    'must be OK, so lets do it
    SocketsInitialize = True
End Function
