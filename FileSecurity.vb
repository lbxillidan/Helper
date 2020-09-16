Imports System.Runtime.InteropServices

Public Class FileSecurity
    Private Const ERROR_INSUFFICIENT_BUFFER As Integer = 122

    <Flags>
    Private Enum SECURITY_INFORMATION As UInt32
        OWNER_SECURITY_INFORMATION = &H1UI
        GROUP_SECURITY_INFORMATION = &H2UI
        DACL_SECURITY_INFORMATION = &H4UI
        SACL_SECURITY_INFORMATION = &H8UI
        LABEL_SECURITY_INFORMATION = &H10UI
        ATTRIBUTE_SECURITY_INFORMATION = &H20UI
        SCOPE_SECURITY_INFORMATION = &H10UI
        BACKUP_SECURITY_INFORMATION = &H10000UI

        PROTECTED_DACL_SECURITY_INFORMATION = &H80000000UI
        PROTECTED_SACL_SECURITY_INFORMATION = &H40000000UI
        UNPROTECTED_DACL_SECURITY_INFORMATION = &H20000000UI
        UNPROTECTED_SACL_SECURITY_INFORMATION = &H10000000UI

        OWNER_GROUP_DACL_SACL = OWNER_SECURITY_INFORMATION + GROUP_SECURITY_INFORMATION + DACL_SECURITY_INFORMATION + SACL_SECURITY_INFORMATION
    End Enum

    'BOOL GetFileSecurityA(
    '  LPCSTR               lpFileName,
    '  SECURITY_INFORMATION RequestedInformation,
    '  PSECURITY_DESCRIPTOR pSecurityDescriptor,
    '  DWORD                nLength,
    '  LPDWORD              lpnLengthNeeded
    ');
    <DllImport("Advapi32.dll", EntryPoint:="GetFileSecurityA", CharSet:=CharSet.Ansi, SetLastError:=True)>
    Private Shared Function GetFileSecurity(ByVal lpFileName As String, ByVal RequestedInformation As SECURITY_INFORMATION, ByVal pSecurityDescriptor As IntPtr, ByVal nLength As UInt32, ByRef lpnLengthNeeded As UInt32) As Boolean
    End Function

    'BOOL SetFileSecurityA(
    '  LPCSTR               lpFileName,
    '  SECURITY_INFORMATION SecurityInformation,
    '  PSECURITY_DESCRIPTOR pSecurityDescriptor
    ');
    <DllImport("Advapi32.dll", EntryPoint:="SetFileSecurityA", CharSet:=CharSet.Ansi, SetLastError:=True)>
    Private Shared Function SetFileSecurity(ByVal lpFileName As String, ByVal SecurityInformation As SECURITY_INFORMATION, ByVal pSecurityDescriptor As IntPtr) As Boolean
    End Function

    Public Shared Message As String = String.Empty
 
    Public Shared Function GetSecurityDescriptor(Path As String, ByRef SecurityDescriptor As Byte()) As Boolean
        Dim SecurityDescriptorSize As Integer = 1024
        Dim SecurityDescriptorPtr As IntPtr = Marshal.AllocHGlobal(SecurityDescriptorSize)
        If GetFileSecurity(Path, SECURITY_INFORMATION.OWNER_GROUP_DACL_SACL, SecurityDescriptorPtr, SecurityDescriptorSize, SecurityDescriptorSize) = False Then
            Dim LastWin32Error As Integer = Marshal.GetLastWin32Error
            If LastWin32Error = ERROR_INSUFFICIENT_BUFFER Then
                SecurityDescriptorPtr = Marshal.AllocHGlobal(SecurityDescriptorSize)
                GetFileSecurity(Path, SECURITY_INFORMATION.OWNER_GROUP_DACL_SACL, SecurityDescriptorPtr, SecurityDescriptorSize, SecurityDescriptorSize)
            Else
                Message = "GetFileSecurity:" & (New System.ComponentModel.Win32Exception(LastWin32Error)).Message
                Marshal.FreeHGlobal(SecurityDescriptorPtr)
                Return False
            End If
        End If
        ReDim SecurityDescriptor(SecurityDescriptorSize - 1)
        Marshal.Copy(SecurityDescriptorPtr, SecurityDescriptor, 0, SecurityDescriptorSize)
        Marshal.FreeHGlobal(SecurityDescriptorPtr)
        Return True
    End Function

    Public Shared Function SetSecurityDescriptor(Path As String, SecurityDescriptorBits As Byte()) As Boolean
        Dim SecurityDescriptorSize As Integer = SecurityDescriptorBits.Length
        Dim SecurityDescriptorPtr As IntPtr = Marshal.AllocHGlobal(SecurityDescriptorSize)
        Marshal.Copy(SecurityDescriptorBits, 0, SecurityDescriptorPtr, SecurityDescriptorSize)
        Dim Result As Boolean =SetFileSecurity(Path, SECURITY_INFORMATION.OWNER_GROUP_DACL_SACL, SecurityDescriptorPtr)
        If Result = False Then
            Message = "SetFileSecurity:" & (New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error)).Message
        End If
        Marshal.FreeHGlobal(SecurityDescriptorPtr)
        Return Result
    End Function

End Class
