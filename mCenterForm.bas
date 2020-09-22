Attribute VB_Name = "mCenterForm"
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Sub CenterForm(frm As Form)
    
    frm.Left = Screen.TwipsPerPixelX * GetSystemMetrics(16) / 2 - frm.Width / 2
    frm.Top = Screen.TwipsPerPixelY * GetSystemMetrics(17) / 2 - frm.Height / 2

End Sub
