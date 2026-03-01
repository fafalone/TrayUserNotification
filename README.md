# TrayUserNotification
IUserNotification2 demo updated for Win8-11

A while back I posted a [VB6 demo](https://www.vbforums.com/showthread.php?795153-VB6-Enhanced-Tray-Message-w-custom-ToolTip-icon-and-feedback-w-o-ShellNotifyIcon) for easier tray popup notifications with IUserNotification2:

![IUserNotification](https://github.com/user-attachments/assets/9a8743cd-f8b5-4b97-8b72-868b686ad314)

However if you run that code on Windows 8-11, you'll notice there's no popup, just an icon hidden in the extra icons overflow you can expand. This is because additional requirements have been imposed on these legacy notification methods (MS wants you to use Toast notifications; not impossible but extraordinarily difficult from VB6/tB). This project not only updates the original for twinBASIC/x64 using WinDevLib, but shows how to actually get the popups on new Windows versions. The biggest downside is it requires you to have a Start Menu shortcut that's been created with specific properties, namely using IPropertyStore to set `PKEY_AppUserModel_ID`, which must then match the string you register via the `SetCurrentProcessExplicitAppUserModelID` API, which must be called prior to window creation, so you'll need to start from `Sub Main()`.

But once you do those steps, you can get a nice modern popup:

<img width="734" height="354" alt="image" src="https://github.com/user-attachments/assets/af8be407-596b-462c-afb4-5bdb8d2cecb8" />

There's also new handling for dismissing the notification properly; previously even if you clicked it you'd end up with a ghost process. Now we set a flag for IQueryContinue to cancel it in response to a left click or balloon click.

This is the new startup code:
```vba
Sub Main()
    If App.IsInIDE Then
        MsgBox "Please run as exe", vbCritical Or vbOKOnly
        Exit Sub
    End If
    SetCurrentProcessExplicitAppUserModelID StrPtr(aumid)
    If CreateStartMenuEntry() Then
        Form1.Show
    End If
    
End Sub

Private Function CreateStartMenuEntry() As Boolean
    On Error GoTo fail
    Dim szPath As String
    szPath = Space$(MAX_PATH)
    SHGetFolderPath(0, CSIDL_PROGRAMS, 0, SHGFP_TYPE_CURRENT, szPath)
    Dim startMenuPath As String = Left$(szPath, InStr(szPath, Chr(0)) - 1) & "\" & appName & ".lnk"
    If PathFileExists(startMenuPath) Then
        Return True 'Already created
    End If
    
    Dim shellLink As IShellLinkW
    Dim persistFile As IPersistFile
    Dim propStore As IPropertyStore
    Dim pv As Variant
    
    ' Create the shortcut
    Set shellLink = New ShellLinkW
    shellLink.SetPath StrPtr(App.Path & "\" & App.EXEName & ".exe")
    shellLink.SetDescription StrPtr(appName)
    
    Set persistFile = shellLink
    Set propStore = shellLink
    
    ' Set the AUMID property
    Dim key As PROPERTYKEY = PKEY_AppUserModel_ID
    
    InitPropVariantFromString aumid, pv
    propStore.SetValue key, pv
    propStore.Commit
    PropVariantClear pv
    
    ' Save the shortcut
    persistFile.Save startMenuPath, CTRUE
    Return True
fail:
    MsgBox Err.Number & ", " & Err.Description
End Function
```


    
