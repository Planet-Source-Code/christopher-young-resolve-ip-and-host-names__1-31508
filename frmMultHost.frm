VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMultHost 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resolve Multiple Hosts"
   ClientHeight    =   3420
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4995
   Icon            =   "frmMultHost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Frame fraRange 
      Caption         =   "IP Range"
      Height          =   615
      Left            =   1440
      TabIndex        =   8
      Top             =   120
      Width           =   1695
      Begin VB.ComboBox cboNum1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cboNum2 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         Caption         =   "to"
         Height          =   195
         Left            =   780
         TabIndex        =   9
         Top             =   330
         Width           =   135
      End
   End
   Begin VB.Frame fraSubnet 
      Caption         =   "Subnet:"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1215
      Begin VB.TextBox txtSubnet 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar stbMult 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   3045
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7347
            MinWidth        =   1341
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1376
            MinWidth        =   1323
            TextSave        =   "11:22 AM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdGetMultHost 
      Caption         =   "Get Host Names"
      Height          =   285
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin MSComctlLib.ListView lstHosts 
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuOtherM 
      Caption         =   "Other"
      Begin VB.Menu mnuResolveIPM 
         Caption         =   "Resolve IP"
      End
      Begin VB.Menu mnuResolveHostM 
         Caption         =   "Resolve Host"
      End
   End
   Begin VB.Menu mnuAboutM 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMultHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Dim GMultHosts As Boolean


Private Sub cboNum1_Click()

Dim i As Integer
Dim n As Integer

n = cboNum1.Text

For i = n + 1 To 255

cboNum2.AddItem (i)

Next i
End Sub

Private Sub cmdGetMultHost_Click()

Dim lHosts As Long
Dim i As Integer
Dim n As Integer
Dim newline As ListItem
Dim IP As String

lstHosts.ListItems.Clear

Me.MousePointer = vbHourglass

GMultHosts = True

For i = cboNum1.Text To cboNum2.Text

If GMultHosts = True Then

    IP = txtSubnet.Text & "." & i
    
    stbMult.Panels(1).Text = "Resolving:  " & IP
    
    DoEvents
    
    lHosts = ConvertToLong(IP)
    DoEvents
    
    Set newline = lstHosts.ListItems.Add(, , IP)
        newline.SubItems(1) = GetHostName(lHosts)
    
    DoEvents
 
End If
 
Next i

stbMult.Panels(1).Text = "Finished"

Me.MousePointer = vbDefault

End Sub

Private Sub cmdStop_Click()

GMultHosts = False

End Sub

Private Sub Form_Load()

CenterForm frmMultHost

stbMult.Panels(1).Text = "Status:"

With lstHosts.ColumnHeaders
    .Add , , "IP Address"
    .Add , , "Host Name"
End With

lstHosts.View = lvwReport

With lstHosts
    .ColumnHeaders(1).Width = lstHosts.Width * 0.35
    .ColumnHeaders(2).Width = lstHosts.Width * 0.6
End With

txtSubnet.Text = ""

Dim i As Integer

For i = 1 To 255

cboNum1.AddItem (i), i - 1

Next i

End Sub
Private Function ConvertToLong(IP As String) As Long

ConvertToLong = inet_addr(IP)

End Function

Private Sub Form_Unload(Cancel As Integer)

Unload frmMultHost
frmResolveIP.Show

End Sub

Private Sub mnuAboutM_Click()

frmAbout.Show

End Sub

Private Sub mnuResolveHostM_Click()

frmMultHost.Hide
frmResolveHost.Show

End Sub

Private Sub mnuResolveIPM_Click()

frmMultHost.Hide
frmResolveIP.Show

End Sub

Private Sub txtSubnet_GotFocus()

txtSubnet.SelStart = 0
txtSubnet.SelLength = Len(txtSubnet.Text)

End Sub

Private Sub txtSubnet_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyEscape
        txtSubnet.Text = "'"

End Select

End Sub
