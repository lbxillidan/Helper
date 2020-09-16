Imports System.Runtime.InteropServices

Public Module ComFileTimeExtension
    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToSysDateTime(ByRef ComFileTime As ComTypes.FILETIME) As Date
        Return Date.FromFileTime((CLng(ComFileTime.dwHighDateTime) << 32) + CLng(ComFileTime.dwLowDateTime))
    End Function
    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToComFileTime(ByRef SysDateTime As Date) As ComTypes.FILETIME
        Dim SysFileTime As UInt64 = SysDateTime.ToFileTime
        Dim ComFileTime As ComTypes.FILETIME
        ComFileTime.dwLowDateTime = SysFileTime << 32 >> 32
        ComFileTime.dwHighDateTime = SysFileTime >> 32
        Return ComFileTime
    End Function
End Module
