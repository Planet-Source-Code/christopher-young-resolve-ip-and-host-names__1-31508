VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmResolveIP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resolve IP"
   ClientHeight    =   990
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3105
   Icon            =   "RESOLV~1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   615
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4013
            MinWidth        =   1341
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1376
            MinWidth        =   1323
            TextSave        =   "11:19 AM"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtHostName 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "Host Name"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdResolve 
      Caption         =   "Resolve IP"
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuOtherIP 
      Caption         =   "Other"
      Begin VB.Menu mnuResolveHost 
         Caption         =   "Resolve Host"
      End
      Begin VB.Menu mnuMultHostIP 
         Caption         =   "Resolve Multiple Hosts"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmResolveIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADESCRIPTION_LEN = 256
Private Const WSASYS_STATUS_LEN = 128

Private Type HOSTENT
   hName As Long
   hAliases As Long
   hAddrType As Integer
   hLength As Integer
   hAddrList As Long
End Type

Private Type WSAData
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADESCRIPTION_LEN) As Byte
   szSystemStatus(0 To WSASYS_STATUS_LEN) As Byte
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpszVendorInfo As Long
End Type

Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSADATA As WSAData) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)


Function HiByte(ByVal wParam As Integer)
   
   HiByte = wParam \ &H100 And &HFF&
   
End Function

Function LoByte(ByVal wParam As Integer)
   
   LoByte = wParam And &HFF&
   
End Function

Sub SocketsInitialize()
   
   Dim WSAD As WSAData
   Dim iReturn As Integer
   Dim sLowByte As String, sHighByte As String, sMsg As String
   
   iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
   
   If iReturn <> 0 Then
      MsgBox "Winsock.dll is not responding."
      End
   End If
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      sHighByte = Trim$(Str$(HiByte(WSAD.wVersion)))
      sLowByte = Trim$(Str$(LoByte(WSAD.wVersion)))
      sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
      sMsg = sMsg & " is not supported by winsock.dll "
      MsgBox sMsg
      End
   End If
   
   If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
      sMsg = "This application requires a minimum of "
      sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
      MsgBox sMsg
      End
   End If
   
End Sub

Sub SocketsCleanup()
   Dim lReturn As Long
   
   lReturn = WSACleanup()
   
   If lReturn <> 0 Then
      MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
      End
   End If
   
End Sub

Sub Form_Load()
   
   SocketsInitialize
   CenterForm frmResolveIP
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   SocketsCleanup
   
Unload frmResolveIP
Unload frmMultHost
Unload frmResolveHost
   
End Sub

Private Sub cmdResolve_Click()

StatusBar1.Panels(1).Text = "IP Address:"
   
   Dim hostent_addr As Long
   Dim host As HOSTENT
   Dim hostip_addr As Long
   Dim temp_ip_address() As Byte
   Dim i As Integer
   Dim ip_address As String
   
   hostent_addr = gethostbyname(txtHostName)
   
   If hostent_addr = 0 Then
      MsgBox "Can't resolve name."
      Exit Sub
   End If
   
   RtlMoveMemory host, hostent_addr, LenB(host)
   RtlMoveMemory hostip_addr, host.hAddrList, 4
   
   ReDim temp_ip_address(1 To host.hLength)
   RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength
   
   For i = 1 To host.hLength
      ip_address = ip_address & temp_ip_address(i) & "."
   Next
   ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
   
   StatusBar1.Panels(1).Text = "IP Address: " & ip_address
   
End Sub

Public Sub mnuAbout_Click()

frmAbout.Show

End Sub

Private Sub mnuMultHostIP_Click()

frmResolveIP.Hide
frmMultHost.Show

End Sub

Private Sub mnuResolveHost_Click()
frmResolveHost.Show
frmResolveIP.Hide
End Sub

Private Sub txtHostName_GotFocus()

   txtHostName.SelStart = 0
   txtHostName.SelLength = Len(txtHostName.Text)

End Sub

Private Sub txtHostName_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyReturn
        Call cmdResolve_Click
    Case vbKeyEscape
        txtHostName.Text = ""
    End Select
        
End Sub
