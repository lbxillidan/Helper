Imports System.Runtime.InteropServices

Public Class Privilege
    Private Const ERROR_INSUFFICIENT_BUFFER As Integer = 122

    'BOOL CloseHandle(
    '  HANDLE hObject
    ');
    <DllImport("Kernel32.dll", SetLastError:=True)> _
    Private Shared Function CloseHandle(ByVal hObject As IntPtr) As Boolean
    End Function

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

        'SE_UNSOLICITED_INPUT_NAME = SeUnsolicitedInputPrivilege
        'SE_CREATE_TOKEN_NAME = SeCreateTokenPrivilege
        'SE_ASSIGNPRIMARYTOKEN_NAME = SeAssignPrimaryTokenPrivilege
        'SE_LOCK_MEMORY_NAME = SeLockMemoryPrivilege
        'SE_INCREASE_QUOTA_NAME = SeIncreaseQuotaPrivilege 
        'SE_MACHINE_ACCOUNT_NAME = SeMachineAccountPrivilege
        'SE_TCB_NAME = SeTcbPrivilege
        'SE_SECURITY_NAME = SeSecurityPrivilege
        'SE_TAKE_OWNERSHIP_NAME = SeTakeOwnershipPrivilege
        'SE_LOAD_DRIVER_NAME = SeLoadDriverPrivilege
        'SE_SYSTEM_PROFILE_NAME = SeSystemProfilePrivilege
        'SE_SYSTEMTIME_NAME = SeSystemtimePrivilege
        'SE_PROF_SINGLE_PROCESS_NAME = SeProfileSingleProcessPrivilege
        'SE_INC_BASE_PRIORITY_NAME = SeIncreaseBasePriorityPrivilege
        'SE_CREATE_PAGEFILE_NAME = SeCreatePagefilePrivilege
        'SE_CREATE_PERMANENT_NAME = SeCreatePermanentPrivilege
        'SE_BACKUP_NAME = SeBackupPrivilege
        'SE_RESTORE_NAME = SeRestorePrivilege
        'SE_SHUTDOWN_NAME = SeShutdownPrivilege
        'SE_DEBUG_NAME = SeDebugPrivilege
        'SE_AUDIT_NAME = SeAuditPrivilege
        'SE_SYSTEM_ENVIRONMENT_NAME = SeSystemEnvironmentPrivilege
        'SE_CHANGE_NOTIFY_NAME = SeChangeNotifyPrivilege
        'SE_REMOTE_SHUTDOWN_NAME = SeRemoteShutdownPrivilege
        'SE_UNDOCK_NAME = SeUndockPrivilege
        'SE_SYNC_AGENT_NAME = SeSyncAgentPrivilege
        'SE_ENABLE_DELEGATION_NAME = SeEnableDelegationPrivilege
        'SE_MANAGE_VOLUME_NAME = SeManageVolumePrivilege
        'SE_IMPERSONATE_NAME = SeImpersonatePrivilege
        'SE_CREATE_GLOBAL_NAME = SeCreateGlobalPrivilege
        'SE_TRUSTED_CREDMAN_ACCESS_NAME = SeTrustedCredManAccessPrivilege
        'SE_RELABEL_NAME = SeRelabelPrivilege
        'SE_INC_WORKING_SET_NAME = SeIncreaseWorkingSetPrivilege
        'SE_TIME_ZONE_NAME = SeTimeZonePrivilege
        'SE_CREATE_SYMBOLIC_LINK_NAME = SeCreateSymbolicLinkPrivilege
        'SE_DELEGATE_SESSION_USER_IMPERSONATE_NAME = SeDelegateSessionUserImpersonatePrivilege
    End Enum

    'BOOL LookupPrivilegeValueW(
    '  LPCWSTR lpSystemName,
    '  LPCWSTR lpName,
    '  PLUID   lpLuid
    ');
    <DllImport("Advapi32.dll", EntryPoint:="LookupPrivilegeValueW", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function LookupPrivilegeValue(ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As UInt64) As Boolean
    End Function

    'BOOL LookupPrivilegeNameW(
    '  LPCWSTR lpSystemName,
    '  PLUID   lpLuid,
    '  LPWSTR  lpName,
    '  LPDWORD cchName
    ');
    <DllImport("advapi32.dll", EntryPoint:="LookupPrivilegeNameW", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function LookupPrivilegeName(ByVal lpSystemName As String, ByVal lpLuid As IntPtr, ByVal lpName As System.Text.StringBuilder, ByRef cchName As UInt32) As Boolean
    End Function

    <Flags>
    Private Enum PrivilegeAttributes As UInt32
        SE_PRIVILEGE_DISABLED = 0
        SE_PRIVILEGE_ENABLED_BY_DEFAULT = 1
        SE_PRIVILEGE_ENABLED = 2
        SE_PRIVILEGE_REMOVED = 4
        SE_PRIVILEGE_USED_FOR_ACCESS = &H80000000UI
    End Enum

    'typedef struct _LUID_AND_ATTRIBUTES {
    '  LUID  Luid;
    '  DWORD Attributes;
    '} LUID_AND_ATTRIBUTES, *PLUID_AND_ATTRIBUTES;
    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure LUID_AND_ATTRIBUTES
        Public Luid As UInt64
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

    'HANDLE GetCurrentProcess();
    <DllImport("Kernel32.dll", SetLastError:=True)> _
    Private Shared Function GetCurrentProcess() As IntPtr
    End Function

    <Flags>
    Private Enum ACCESS_TYPE As UInt32
        DELETE = &H10000UI
        READ_CONTROL = &H20000UI
        WRITE_DAC = &H40000UI
        WRITE_OWNER = &H80000UI
        SYNCHRONIZE = &H100000UI

        STANDARD_RIGHTS_REQUIRED = DELETE + READ_CONTROL + WRITE_DAC + WRITE_OWNER
        STANDARD_RIGHTS_READ = READ_CONTROL
        STANDARD_RIGHTS_WRITE = READ_CONTROL
        STANDARD_RIGHTS_EXECUTE = READ_CONTROL
        STANDARD_RIGHTS_ALL = DELETE + READ_CONTROL + WRITE_DAC + WRITE_OWNER + SYNCHRONIZE

        SPECIFIC_RIGHTS_ALL = &HFFFFUI

        ACCESS_SYSTEM_SECURITY = &H1000000UI
        MAXIMUM_ALLOWED = &H2000000UI

        GENERIC_READ = &H80000000UI
        GENERIC_WRITE = &H40000000UI
        GENERIC_EXECUTE = &H20000000UI
        GENERIC_ALL = &H10000000UI
    End Enum

    <Flags>
    Private Enum TOKEN_ACCESS As UInt32
        TOKEN_ASSIGN_PRIMARY = &H1
        TOKEN_DUPLICATE = &H2
        TOKEN_IMPERSONATE = &H4
        TOKEN_QUERY = &H8
        TOKEN_QUERY_SOURCE = &H10
        TOKEN_ADJUST_PRIVILEGES = &H20
        TOKEN_ADJUST_GROUPS = &H40
        TOKEN_ADJUST_DEFAULT = &H80
        TOKEN_ADJUST_SESSIONID = &H100

        TOKEN_ALL_ACCESS = ACCESS_TYPE.STANDARD_RIGHTS_REQUIRED + TOKEN_ASSIGN_PRIMARY + TOKEN_DUPLICATE + TOKEN_IMPERSONATE + TOKEN_QUERY + TOKEN_QUERY_SOURCE + TOKEN_ADJUST_PRIVILEGES + TOKEN_ADJUST_GROUPS + TOKEN_ADJUST_DEFAULT + TOKEN_ADJUST_SESSIONID
        TOKEN_READ = ACCESS_TYPE.STANDARD_RIGHTS_READ + TOKEN_QUERY
        TOKEN_WRITE = ACCESS_TYPE.STANDARD_RIGHTS_WRITE + TOKEN_ADJUST_PRIVILEGES + TOKEN_ADJUST_GROUPS + TOKEN_ADJUST_DEFAULT
        TOKEN_EXECUTE = ACCESS_TYPE.STANDARD_RIGHTS_EXECUTE
    End Enum

    'BOOL OpenProcessToken(
    '  HANDLE  ProcessHandle,
    '  DWORD   DesiredAccess,
    '  PHANDLE TokenHandle
    ');
    <DllImport("Advapi32.dll", SetLastError:=True)> _
    Private Shared Function OpenProcessToken(ByVal ProcessHandle As IntPtr, ByVal DesiredAccess As UInt32, ByRef TokenHandle As IntPtr) As Boolean
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
    Private Shared Function AdjustTokenPrivileges(ByVal TokenHandle As IntPtr, ByVal DisableAllPrivileges As Boolean, ByVal NewState As TOKEN_PRIVILEGES, ByVal BufferLength As UInt32, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As UInt32) As Boolean
    End Function

    Private Const PRIVILEGE_SET_ALL_NECESSARY As UInt32 = 1

    'typedef struct _PRIVILEGE_SET {
    '  DWORD               PrivilegeCount;
    '  DWORD               Control;
    '  LUID_AND_ATTRIBUTES Privilege[ANYSIZE_ARRAY];
    '} PRIVILEGE_SET, *PPRIVILEGE_SET;
    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure PRIVILEGE_SET
        Public PrivilegeCount As UInt32
        Public Control As UInt32
        <MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst:=1, ArraySubType:=UnmanagedType.Struct)> Public Privilege As LUID_AND_ATTRIBUTES()
    End Structure

    'BOOL PrivilegeCheck(
    '  HANDLE         ClientToken,
    '  PPRIVILEGE_SET RequiredPrivileges,
    '  LPBOOL         pfResult
    ');
    <DllImport("Advapi32.dll", SetLastError:=True)> _
    Private Shared Function PrivilegeCheck(ByVal ClientToken As IntPtr, ByRef RequiredPrivileges As PRIVILEGE_SET, ByRef pfResult As Boolean) As Boolean
    End Function

    Private Enum TOKEN_INFORMATION_CLASS As UInt32
        TokenUser = 1
        TokenGroups
        TokenPrivileges
        TokenOwner
        TokenPrimaryGroup
        TokenDefaultDacl
        TokenSource
        TokenType
        TokenImpersonationLevel
        TokenStatistics
        TokenRestrictedSids
        TokenSessionId
        TokenGroupsAndPrivileges
        TokenSessionReference
        TokenSandBoxInert
        TokenAuditPolicy
        TokenOrigin
        TokenElevationType
        TokenLinkedToken
        TokenElevation
        TokenHasRestrictions
        TokenAccessInformation
        TokenVirtualizationAllowed
        TokenVirtualizationEnabled
        TokenIntegrityLevel
        TokenUIAccess
        TokenMandatoryPolicy
        TokenLogonSid
        TokenIsAppContainer
        TokenCapabilities
        TokenAppContainerSid
        TokenAppContainerNumber
        TokenUserClaimAttributes
        TokenDeviceClaimAttributes
        TokenRestrictedUserClaimAttributes
        TokenRestrictedDeviceClaimAttributes
        TokenDeviceGroups
        TokenRestrictedDeviceGroups
        TokenSecurityAttributes
        TokenIsRestricted
        MaxTokenInfoClass
    End Enum

    'BOOL GetTokenInformation(
    '  HANDLE                  TokenHandle,
    '  TOKEN_INFORMATION_CLASS TokenInformationClass,
    '  LPVOID                  TokenInformation,
    '  DWORD                   TokenInformationLength,
    '  PDWORD                  ReturnLength
    ');
    <DllImport("Advapi32.dll", SetLastError:=True)> _
    Private Shared Function GetTokenInformation(ByVal TokenHandle As IntPtr, ByVal TokenInformationClass As TOKEN_INFORMATION_CLASS, ByVal TokenInformation As IntPtr, ByVal TokenInformationLength As UInt32, ByRef ReturnLength As UInt32) As Boolean
    End Function

    Public Shared Message As String = String.Empty

    Private Shared Function GetPrivilegeLuid(PrivilegeName As String, ByRef PrivilegeLuid As UInt64) As Boolean
        If LookupPrivilegeValue(Nothing, PrivilegeName, PrivilegeLuid) = False Then
            Message = "LookupPrivilegeValue:" & (New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error)).Message
            Return False
        Else
            Return True
        End If
    End Function

    Private Shared Function GetPrivilegeName(PrivilegeLuid As UInt64, ByRef PrivilegeName As String) As Boolean
        Dim Result As Boolean = False

        Dim PrivilegeLuidPtr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(PrivilegeLuid))
        Marshal.StructureToPtr(PrivilegeLuid, PrivilegeLuidPtr, True)

        Dim PrivilegeNameSize As UInt32
        LookupPrivilegeName(Nothing, PrivilegeLuidPtr, Nothing, PrivilegeNameSize)
        Dim LastWin32Error As Integer = Marshal.GetLastWin32Error
        If LastWin32Error = ERROR_INSUFFICIENT_BUFFER Then
            Dim PrivilegeNameBuilder As New System.Text.StringBuilder()
            PrivilegeNameBuilder.EnsureCapacity(PrivilegeNameSize)
            LookupPrivilegeName(Nothing, PrivilegeLuidPtr, PrivilegeNameBuilder, PrivilegeNameSize)
            PrivilegeName = PrivilegeNameBuilder.ToString
            If String.IsNullOrEmpty(PrivilegeName) Then
                Result = False
            Else
                Result = True
            End If
        Else
            Message = "LookupPrivilegeName:" & (New System.ComponentModel.Win32Exception(LastWin32Error)).Message
            Result = False
        End If

        Marshal.FreeHGlobal(PrivilegeLuidPtr)

        Return Result
    End Function

    Private Shared Function GetTokenHandle(ByRef TokenHandle As IntPtr) As Boolean
        Dim Result As Boolean = False
        Dim ProcessHandle As IntPtr = GetCurrentProcess()
        If OpenProcessToken(ProcessHandle, TOKEN_ACCESS.TOKEN_ADJUST_PRIVILEGES + TOKEN_ACCESS.TOKEN_QUERY, TokenHandle) = False Then
            Message = "OpenProcessToken:" & (New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error)).Message
        Else
            Result = True
        End If

        CloseHandle(ProcessHandle)

        Return Result
    End Function

    Private Shared Function CheckPrivilege(TokenHandle As IntPtr, PrivilegeLuid As UInt64) As Boolean
        Dim PrivilegeSet As New PRIVILEGE_SET
        PrivilegeSet.PrivilegeCount = 1
        PrivilegeSet.Control = PRIVILEGE_SET_ALL_NECESSARY
        ReDim PrivilegeSet.Privilege(0)
        PrivilegeSet.Privilege(0).Luid = PrivilegeLuid
        PrivilegeSet.Privilege(0).Attributes = PrivilegeAttributes.SE_PRIVILEGE_DISABLED

        Dim CheckResult As Boolean = False
        If PrivilegeCheck(TokenHandle, PrivilegeSet, CheckResult) = False Then
            Message = "PrivilegeCheck:" & (New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error)).Message
            CloseHandle(TokenHandle)
            Return False
        Else
            CloseHandle(TokenHandle)
        End If

        If CheckResult = True AndAlso PrivilegeSet.Privilege(0).Attributes = PrivilegeAttributes.SE_PRIVILEGE_USED_FOR_ACCESS Then
            Return True
        Else
            Message = String.Format("CheckPrivilege:CheckResult={0},Privilege.Attributes={1}", CheckResult, PrivilegeSet.Privilege(0).Attributes.ToString)
            Return False
        End If
    End Function

    Public Shared Function CheckPrivilege(Privilege As SePrivilege) As Boolean
        Dim PrivilegeLuid As UInt64
        Dim TokenHandle As IntPtr

        If GetPrivilegeLuid(Privilege.ToString, PrivilegeLuid) = False Then
            Return False
        End If

        If GetTokenHandle(TokenHandle) = False Then
            Return False
        End If

        Return CheckPrivilege(TokenHandle, PrivilegeLuid)
    End Function

    Public Shared Function EnablePrivilege(Privilege As SePrivilege) As Boolean
        Dim PrivilegeLuid As UInt64
        Dim TokenHandle As IntPtr

        If GetPrivilegeLuid(Privilege.ToString, PrivilegeLuid) = False Then
            Return False
        End If

        If GetTokenHandle(TokenHandle) = False Then
            Return False
        End If

        Dim TokenPrivileges As New TOKEN_PRIVILEGES
        TokenPrivileges.PrivilegeCount = 1
        ReDim TokenPrivileges.Privileges(0)
        TokenPrivileges.Privileges(0).Luid = PrivilegeLuid
        TokenPrivileges.Privileges(0).Attributes = PrivilegeAttributes.SE_PRIVILEGE_ENABLED

        Dim TokenPrivilegesSize As UInt32 = Marshal.SizeOf(TokenPrivileges)
        If AdjustTokenPrivileges(TokenHandle, False, TokenPrivileges, TokenPrivilegesSize, Nothing, Nothing) = False Then
            Message = "AdjustTokenPrivileges:" & (New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error)).Message
            CloseHandle(TokenHandle)
            Return False
        End If

        Return CheckPrivilege(TokenHandle, PrivilegeLuid)
    End Function

    Public Shared Function QueryPrivileges(ByRef PrivilegeDict As Dictionary(Of String, Boolean)) As Boolean
        Dim TokenHandle As IntPtr
        If GetTokenHandle(TokenHandle) = False Then
            Return False
        End If

        Dim Result As Boolean = False

        Dim TokenPrivileges As New TOKEN_PRIVILEGES
        Dim TokenPrivilegesSize As Int32 = 0
        GetTokenInformation(TokenHandle, TOKEN_INFORMATION_CLASS.TokenPrivileges, Nothing, Nothing, TokenPrivilegesSize)
        If TokenPrivilegesSize > 0 Then
            Dim TokenPrivilegesPtr As IntPtr = Marshal.AllocHGlobal(TokenPrivilegesSize)
            GetTokenInformation(TokenHandle, TOKEN_INFORMATION_CLASS.TokenPrivileges, TokenPrivilegesPtr, TokenPrivilegesSize, TokenPrivilegesSize)
            TokenPrivileges.PrivilegeCount = Marshal.ReadInt32(TokenPrivilegesPtr)
            ReDim TokenPrivileges.Privileges(TokenPrivileges.PrivilegeCount - 1)
            Dim LuidAndAttributesSize As Long = Marshal.SizeOf(GetType(LUID_AND_ATTRIBUTES))
            For i As Integer = 0 To TokenPrivileges.PrivilegeCount - 1
                TokenPrivileges.Privileges(i) = Marshal.PtrToStructure(IntPtr.Add(TokenPrivilegesPtr, 4 + i * LuidAndAttributesSize), GetType(LUID_AND_ATTRIBUTES))
            Next
            Marshal.FreeHGlobal(TokenPrivilegesPtr)
        Else
            Message = "GetTokenInformation:" & (New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error)).Message
            Result = False
        End If

        CloseHandle(TokenHandle)

        Dim PrivilegeDictTemp As New Dictionary(Of String, Boolean)
        If TokenPrivileges.PrivilegeCount > 0 Then
            For i As Integer = 0 To TokenPrivileges.PrivilegeCount - 1
                Dim LuidAndAttributes As LUID_AND_ATTRIBUTES = TokenPrivileges.Privileges(i)
                Dim PrivilegeLuid As UInt64 = LuidAndAttributes.Luid
                Dim PrivilegeName As String = Nothing
                If GetPrivilegeName(PrivilegeLuid, PrivilegeName) = False Then
                    Exit For
                End If
                Dim PrivilegeEnabled As Boolean = LuidAndAttributes.Attributes.HasFlag(PrivilegeAttributes.SE_PRIVILEGE_ENABLED)
                PrivilegeDictTemp.Add(PrivilegeName, PrivilegeEnabled)
            Next
            If PrivilegeDictTemp.Count = TokenPrivileges.PrivilegeCount Then
                PrivilegeDict = PrivilegeDictTemp
                Result = True
            End If
        End If

        Return Result
    End Function

    Private Shared Function QueryPrivileges() As String
        Dim PrivilegeDict As Dictionary(Of String, Boolean) = Nothing
        If QueryPrivileges(PrivilegeDict) = True Then
            Dim PrivilegeBuilder As New System.Text.StringBuilder
            For Each Pair In PrivilegeDict
                PrivilegeBuilder.AppendFormat("{1}{0}{2}", vbTab, Pair.Key.PadRight(50, " "), If(Pair.Value, "1", "0"))
                PrivilegeBuilder.AppendLine()
            Next
            Return PrivilegeBuilder.ToString
        Else
            Return Message
        End If
    End Function

    Private Shared Sub Test()
        Debug.Print("--------------------------------------------------")
        For Each Privilege As SePrivilege In [Enum].GetValues(GetType(SePrivilege))
            Debug.Print(Privilege.ToString & ":" & CheckPrivilege(Privilege))
        Next
        Debug.Print("--------------------------------------------------")
        Debug.Print(QueryPrivileges())
        Debug.Print("--------------------------------------------------")
        Dim PrivilegeDict As Dictionary(Of String, Boolean) = Nothing
        If QueryPrivileges(PrivilegeDict) = True Then
            For Each Pair In PrivilegeDict
                If Pair.Value = False Then
                    EnablePrivilege([Enum].Parse(GetType(SePrivilege), Pair.Key))
                End If
            Next
        Else
            Debug.Print(Message)
        End If
        Debug.Print(QueryPrivileges())
    End Sub

End Class
