VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUserNotificationCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'All of these functions expect an HRESULT to be returned
'S_OK to continue, anything else cancels the notification

Implements IUserNotificationCallback

Private Sub IUserNotificationCallback_OnBalloonUserClick(pt As POINT)
Form1.List1.AddItem "Balloon click"
mCancel = True
End Sub
Private Sub IUserNotificationCallback_OnLeftClick(pt As POINT)
Form1.List1.AddItem "Left click"
Err.ReturnHResult = 1
mCancel = True
End Sub
Private Sub IUserNotificationCallback_OnContextMenu(pt As POINT)
    Form1.List1.AddItem "Context menu"
    Dim hMenu As LongPtr
    Dim idCmd As Long
    hMenu = CreatePopupMenu()
    Call AppendMenu(hMenu, 0, 100, "Hide icon")
    Call AppendMenu(hMenu, 0, 101, "Leave icon alone")
    idCmd = TrackPopupMenu(hMenu, TPM_LEFTBUTTON Or TPM_RIGHTBUTTON Or TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_HORIZONTAL Or TPM_RETURNCMD, pt.x, pt.y, 0, Form1.hWnd, 0)

    If idCmd = 100 Then Err.ReturnHResult = 1

    
End Sub
