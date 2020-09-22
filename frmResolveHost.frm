VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmResolveHost 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resolve Host"
   ClientHeight    =   1080
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3780
   Icon            =   "frmResolveHost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar stbHost 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   705
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5204
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
   Begin VB.CommandButton cmdResolveHost 
      Caption         =   "Resolve Host Name"
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuOtherH 
      Caption         =   "Other"
      Begin VB.Menu mnuResolveIP 
         Caption         =   "Resolve IP"
      End
      Begin VB.Menu mnuMultHostH 
         Caption         =   "Resolve Multipule Hosts"
      End
   End
   Begin VB.Menu mnuAbout2 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmResolveHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

Private Sub cmdResolveHost_Click()

Dim lHostName As Long

stbHost.Panels(1).Text = "Host Name:"

lHostName = ConvertToLong(txtIP.Text)

stbHost.Panels(1).Text = GetHostName(lHostName)

End Sub



Private Sub Form_Load()

CenterForm frmResolveHost

End Sub

Private Sub Form_Unload(Cancel As Integer)

Unload frmResolveHost
frmResolveIP.Show

End Sub

Private Sub mnuAbout2_Click()

Call frmAbout.Show

End Sub

Private Sub mnuMultHostH_Click()

frmResolveHost.Hide
frmMultHost.Show

End Sub

Private Sub mnuResolveIP_Click()

frmResolveIP.Show
frmResolveHost.Hide

End Sub

Private Function ConvertToLong(IP As String) As Long

ConvertToLong = inet_addr(IP)

End Function

Private Sub txtIP_GotFocus()

   txtIP.SelStart = 0
   txtIP.SelLength = Len(txtIP.Text)

End Sub

Private Sub txtIP_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyReturn
        Call cmdResolveHost_Click
    Case vbKeyEscape
        txtIP.Text = "'"
End Select

End Sub
