Public Class DosDateTimeHelper
    'https://docs.microsoft.com/zh-cn/windows/win32/api/winbase/nf-winbase-dosdatetimetofiletime
    'wFatDate
    'The MS-DOS date. The date is a packed value with the following format.
    'Bits	Description
    '0-4	Day of the month (1–31)
    '5-8	Month (1 = January, 2 = February, and so on)
    '9-15	Year offset from 1980 (add 1980 to get actual year)
    'wFatTime
    'The MS-DOS time. The time is a packed value with the following format.
    'Bits	Description
    '0-4	Second divided by 2
    '5-10	Minute (0–59)
    '11-15	Hour (0–23 on a 24-hour clock)

    Private Shared ReadOnly MarkYer As UInt16 = Convert.ToUInt16("1111111000000000", 2)
    Private Shared ReadOnly MarkMon As UInt16 = Convert.ToUInt16("0000000111100000", 2)
    Private Shared ReadOnly MarkDay As UInt16 = Convert.ToUInt16("0000000000011111", 2)
    Private Shared ReadOnly MarkHor As UInt16 = Convert.ToUInt16("1111100000000000", 2)
    Private Shared ReadOnly MarkMin As UInt16 = Convert.ToUInt16("0000011111100000", 2)
    Private Shared ReadOnly MarkSec As UInt16 = Convert.ToUInt16("0000000000011111", 2)

    Public Shared Function GetSystemDateTime(dosDate As UInt16, dosTime As UInt16) As System.DateTime
        Dim Yer As UInt16 = (dosDate And MarkYer) >> 9
        Dim Mon As UInt16 = (dosDate And MarkMon) >> 5
        Dim Day As UInt16 = (dosDate And MarkDay)
        Dim Hor As UInt16 = (dosTime And MarkHor) >> 11
        Dim Min As UInt16 = (dosTime And MarkMin) >> 5
        Dim Sec As UInt16 = (dosTime And MarkSec)
        Return New System.DateTime(1980 + Yer, Mon, Day, Hor, Min, Sec * 2)
    End Function

    Public Shared Sub SetSystemDateTime(sysDateTime As System.DateTime, ByRef dosDate As UInt16, ByRef dosTime As UInt16)
        Dim Yer As UInt16 = CUShort(sysDateTime.Year - 1980) << 9
        Dim Mon As UInt16 = CUShort(sysDateTime.Month) << 5
        Dim Day As UInt16 = CUShort(sysDateTime.Day)
        Dim Hor As UInt16 = CUShort(sysDateTime.Hour) << 11
        Dim Min As UInt16 = CUShort(sysDateTime.Minute) << 5
        Dim Sec As UInt16 = CUShort(sysDateTime.Second / 2)
        dosDate = Yer + Mon + day
        dosTime = Hor + Min + Sec
    End Sub
End Class
