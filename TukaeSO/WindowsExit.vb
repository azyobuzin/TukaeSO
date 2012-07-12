Namespace TukaeruModuleAndClass
    ''' <summary>Windowsを簡単に終了する方法を提供します。</summary>
    Public MustInherit Class WindowsExit
#Region "プライベート"
        Private Enum ExitWindows
            EWX_LOGOFF = &H0
            EWX_SHUTDOWN = &H1
            EWX_REBOOT = &H2
            EWX_POWEROFF = &H8
            EWX_RESTARTAPPS = &H40
            EWX_FORCE = &H4
            EWX_FORCEIFHUNG = &H10
        End Enum
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)> _
        Private Shared Function ExitWindowsEx(ByVal uFlags As ExitWindows, _
            ByVal dwReason As Integer) As Boolean
        End Function
        <System.Runtime.InteropServices.DllImport("kernel32.dll")> _
        Private Shared Function GetCurrentProcess() As IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError:=True)> _
        Private Shared Function OpenProcessToken(ByVal ProcessHandle As IntPtr, _
            ByVal DesiredAccess As Integer, _
            ByRef TokenHandle As IntPtr) As Boolean
        End Function

        <System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError:=True, _
            CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
        Private Shared Function LookupPrivilegeValue(ByVal lpSystemName As String, _
            ByVal lpName As String, _
            ByRef lpLuid As Long) As Boolean
        End Function

        <System.Runtime.InteropServices.StructLayout( _
            System.Runtime.InteropServices.LayoutKind.Sequential, Pack:=1)> _
        Private Structure TOKEN_PRIVILEGES
            Public PrivilegeCount As Integer
            Public Luid As Long
            Public Attributes As Integer
        End Structure

        <System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError:=True)> _
        Private Shared Function AdjustTokenPrivileges(ByVal TokenHandle As IntPtr, _
            ByVal DisableAllPrivileges As Boolean, _
            ByRef NewState As TOKEN_PRIVILEGES, _
            ByVal BufferLength As Integer, _
            ByVal PreviousState As IntPtr, _
            ByVal ReturnLength As IntPtr) As Boolean
        End Function
        'シャットダウンするためのセキュリティ特権を有効にする
        Private Shared Sub AdjustToken()
            Const TOKEN_ADJUST_PRIVILEGES As Integer = &H20
            Const TOKEN_QUERY As Integer = &H8
            Const SE_PRIVILEGE_ENABLED As Integer = &H2
            Const SE_SHUTDOWN_NAME As String = "SeShutdownPrivilege"

            If Environment.OSVersion.Platform <> PlatformID.Win32NT Then
                Return
            End If

            Dim procHandle As IntPtr = GetCurrentProcess()

            'トークンを取得する
            Dim tokenHandle As IntPtr
            OpenProcessToken(procHandle, _
                TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, tokenHandle)
            'LUIDを取得する
            Dim tp As New TOKEN_PRIVILEGES()
            tp.Attributes = SE_PRIVILEGE_ENABLED
            tp.PrivilegeCount = 1
            LookupPrivilegeValue(Nothing, SE_SHUTDOWN_NAME, tp.Luid)
            '特権を有効にする
            AdjustTokenPrivileges(tokenHandle, False, tp, 0, IntPtr.Zero, IntPtr.Zero)
        End Sub
#End Region
        ''' <summary>シャットダウンします。</summary>
        Public Shared Sub Shutdown()
            AdjustToken()
            ExitWindowsEx(ExitWindows.EWX_POWEROFF, 0)
        End Sub
        ''' <summary>ログオフします。</summary>
        Public Shared Sub LogOff()
            ExitWindowsEx(ExitWindows.EWX_LOGOFF, 0)
        End Sub
        ''' <summary>再起動します。</summary>
        Public Shared Sub Restart()
            AdjustToken()
            ExitWindowsEx(ExitWindows.EWX_REBOOT, 0)
        End Sub
        ''' <summary>スタンバイまたはスリープにします。</summary>
        Public Shared Sub Suspend()
            Application.SetSuspendState(PowerState.Suspend, False, False)
        End Sub
        ''' <summary>休止状態にします。</summary>
        Public Shared Sub Hibernate()
            Application.SetSuspendState(PowerState.Hibernate, False, False)
        End Sub
    End Class
End Namespace