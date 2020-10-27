Imports System.Runtime.InteropServices

Public Class Privilege
    Public Shared Message As String = String.Empty

    Public Enum SePrivilege As Integer
        'SeUnsolicitedInputPrivilege = 1
        SeCreateTokenPrivilege = 2
        SeAssignPrimaryTokenPrivilege
        SeLockMemoryPrivilege
        SeIncreaseQuotaPrivilege
        SeMachineAccountPrivilege
        SeTcbPrivilege
        SeSecurityPrivilege
        SeTakeOwnershipPrivilege
        SeLoadDriverPrivilege
        SeSystemProfilePrivilege
        SeSystemtimePrivilege
        SeProfileSingleProcessPrivilege
        SeIncreaseBasePriorityPrivilege
        SeCreatePagefilePrivilege
        SeCreatePermanentPrivilege
        SeBackupPrivilege
        SeRestorePrivilege
        SeShutdownPrivilege
        SeDebugPrivilege
        SeAuditPrivilege
        SeSystemEnvironmentPrivilege
        SeChangeNotifyPrivilege
        SeRemoteShutdownPrivilege
        SeUndockPrivilege
        SeSyncAgentPrivilege
        SeEnableDelegationPrivilege
        SeManageVolumePrivilege
        SeImpersonatePrivilege
        SeCreateGlobalPrivilege
        SeTrustedCredManAccessPrivilege
        SeRelabelPrivilege
        SeIncreaseWorkingSetPrivilege
        SeTimeZonePrivilege
        SeCreateSymbolicLinkPrivilege
        SeDelegateSessionUserImpersonatePrivilege
    End Enum

    <Flags>
    Private Enum PrivilegeAttributes As UInt32
        SE_PRIVILEGE_DISABLED = 0
        SE_PRIVILEGE_ENABLED_BY_DEFAULT = 1
        SE_PRIVILEGE_ENABLED = 2
        SE_PRIVILEGE_REMOVED = 4
        SE_PRIVILEGE_USED_FOR_ACCESS = &H80000000UI
    End Enum

    'typedef struct _LUID {
    '  DWORD LowPart;
    '  LONG  HighPart;
    '} LUID, *PLUID;
    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure LUID
        Public LowPart As UInt32
        Public HighPart As Int32
    End Structure

    'typedef struct _LUID_AND_ATTRIBUTES {
    '  LUID  Luid;
    '  DWORD Attributes;
    '} LUID_AND_ATTRIBUTES, *PLUID_AND_ATTRIBUTES;
    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure LUID_AND_ATTRIBUTES
        Public Luid As LUID
        Public Attributes As PrivilegeAttributes
    End Structure

    'typedef struct _TOKEN_PRIVILEGES {
    '  DWORD               PrivilegeCount;
    '  LUID_AND_ATTRIBUTES Privileges[ANYSIZE_ARRAY];
    '} TOKEN_PRIVILEGES, *PTOKEN_PRIVILEGES;
    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure TOKEN_PRIVILEGES
        Public PrivilegeCount As UInt32
        <MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst:=1, ArraySubType:=UnmanagedType.Struct)> Public Privileges As LUID_AND_ATTRIBUTES()
    End Structure

    'BOOL LookupPrivilegeValueW(
    '  LPCWSTR lpSystemName,
    '  LPCWSTR lpName,
    '  PLUID   lpLuid
    ');
    <DllImport("Advapi32.dll", EntryPoint:="LookupPrivilegeValueW", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function LookupPrivilegeValue(ByVal lpSystemName As String, ByVal lpName As String, ByVal lpLuid As IntPtr) As Boolean
    End Function

    'BOOL AdjustTokenPrivileges(
    '  HANDLE            TokenHandle,
    '  BOOL              DisableAllPrivileges,
    '  PTOKEN_PRIVILEGES NewState,
    '  DWORD             BufferLength,
    '  PTOKEN_PRIVILEGES PreviousState,
    '  PDWORD            ReturnLength
    ');
    <DllImport("Advapi32.dll", SetLastError:=True)> _
    Private Shared Function AdjustTokenPrivileges(ByVal TokenHandle As IntPtr, ByVal DisableAllPrivileges As Boolean, ByVal NewState As IntPtr, ByVal BufferLength As UInt32, ByVal PreviousState As IntPtr, ByRef ReturnLength As UInt32) As Boolean
    End Function

    'BOOL CloseHandle(
    '  HANDLE hObject
    ');
    <DllImport("Kernel32.dll", SetLastError:=True)> _
    Private Shared Function CloseHandle(ByVal hObject As IntPtr) As Boolean
    End Function

    'HANDLE GetCurrentProcess();
    <DllImport("Kernel32.dll", SetLastError:=True)> _
    Private Shared Function GetCurrentProcess() As IntPtr
    End Function

    'BOOL OpenProcessToken(
    '  HANDLE  ProcessHandle,
    '  DWORD   DesiredAccess,
    '  PHANDLE TokenHandle
    ');
    <DllImport("Advapi32.dll", SetLastError:=True)> _
    Private Shared Function OpenProcessToken(ByVal ProcessHandle As IntPtr, ByVal DesiredAccess As UInt32, ByRef TokenHandle As IntPtr) As Boolean
    End Function

    Public Shared Function EnablePrivilege(Privilege As SePrivilege) As Boolean
        '查找PrivilegeLuid
        Dim PrivilegeName As String = Privilege.ToString
        Dim PrivilegeLuid As LUID
        Dim PrivilegeLuidPtr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(PrivilegeLuid))
        If LookupPrivilegeValue(Nothing, PrivilegeName, PrivilegeLuidPtr) = False Then
            Message = "LookupPrivilegeValue:" & (New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error)).Message
            Marshal.FreeHGlobal(PrivilegeLuidPtr)
            Return False
        Else
            PrivilegeLuid = Marshal.PtrToStructure(PrivilegeLuidPtr, GetType(LUID))
            Marshal.FreeHGlobal(PrivilegeLuidPtr)
        End If

        '打开进程Token句柄
        Const TOKEN_QUERY As UInt32 = &H8UI
        Const TOKEN_ADJUST_PRIVILEGES As UInt32 = &H20UI
        Dim ProcessHandle As IntPtr = GetCurrentProcess()
        Dim TokenHandle As IntPtr
        If OpenProcessToken(ProcessHandle, TOKEN_ADJUST_PRIVILEGES + TOKEN_QUERY, TokenHandle) = False Then
            Message = "OpenProcessToken:" & (New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error)).Message
            CloseHandle(ProcessHandle)
            Return False
        Else
            CloseHandle(ProcessHandle)
        End If

        '构造Privilege参数
        Dim TokenPrivileges As New TOKEN_PRIVILEGES
        TokenPrivileges.PrivilegeCount = 1
        ReDim TokenPrivileges.Privileges(0)
        TokenPrivileges.Privileges(0).Luid = PrivilegeLuid
        TokenPrivileges.Privileges(0).Attributes = PrivilegeAttributes.SE_PRIVILEGE_ENABLED
        Dim TokenPrivilegesSize As Int32 = Marshal.SizeOf(TokenPrivileges)
        Dim TokenPrivilegesPtr As IntPtr = Marshal.AllocHGlobal(TokenPrivilegesSize)
        Marshal.StructureToPtr(TokenPrivileges, TokenPrivilegesPtr, True)

        '设置权限
        If AdjustTokenPrivileges(TokenHandle, False, TokenPrivilegesPtr, TokenPrivilegesSize, Nothing, Nothing) = False Then
            Message = "AdjustTokenPrivileges:" & (New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error)).Message
            CloseHandle(TokenHandle)
            Marshal.FreeHGlobal(TokenPrivilegesPtr)
            Return False
        Else
            Dim LastError As Int32 = Marshal.GetLastWin32Error
            Const ERROR_SUCCESS As Int32 = 0
            If LastError = ERROR_SUCCESS Then
                CloseHandle(TokenHandle)
                Marshal.FreeHGlobal(TokenPrivilegesPtr)
                Return True
            Else
                Message = "AdjustTokenPrivileges:" & (New System.ComponentModel.Win32Exception(LastError)).Message
                CloseHandle(TokenHandle)
                Marshal.FreeHGlobal(TokenPrivilegesPtr)
                Return False
            End If
        End If
    End Function
End Class
