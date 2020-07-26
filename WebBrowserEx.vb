Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class WebBrowserEx
    Inherits System.Windows.Forms.WebBrowser

    Sub New()
        MyBase.ScriptErrorsSuppressed = True
    End Sub

    ''' <summary>
    ''' 获取HtmlText方法
    ''' </summary>
    ''' <remarks></remarks>
    Public HtmlTextPolicy As Integer = 0

    Public ReadOnly Property HtmlText As String
        Get
            Dim Html As String = Nothing
            Select Case HtmlTextPolicy
                Case 0
                    Html = MyBase.DocumentText
                Case 1
                    If MyBase.Document IsNot Nothing Then
                        Dim HtmlEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding(MyBase.Document.Encoding)
                        Dim HtmlStream As New System.IO.MemoryStream
                        MyBase.DocumentStream.Position = 0
                        MyBase.DocumentStream.CopyTo(HtmlStream)
                        Html = HtmlEncoding.GetString(HtmlStream.ToArray)
                    End If
                Case 2
                    If MyBase.Document IsNot Nothing Then
                        Dim HtmlElements As HtmlElementCollection = MyBase.Document.GetElementsByTagName("HTML")
                        If HtmlElements.Count > 0 Then
                            Html = HtmlElements(0).OuterHtml
                        End If
                    End If
            End Select
            Return Html
        End Get
    End Property

    Private UrlCompleted As String
    Private Sub WebBrowser_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles MyBase.DocumentCompleted
        UrlCompleted = e.Url.ToString
        Debug.Print("----------------------------------------------------------------------------------------------------")
        Debug.Print(Date.Now.ToString("yyyy-MM-dd HH:mm:ss.ff ") & e.Url.ToString)
    End Sub

    <Flags>
    Private Enum CACHE_ENTRY_TYPE As UInt32
        STICKY_CACHE_ENTRY = &H4
        EDITED_CACHE_ENTRY = &H8
        TRACK_OFFLINE_CACHE_ENTRY = &H10
        TRACK_ONLINE_CACHE_ENTRY = &H20
        SPARSE_CACHE_ENTRY = &H10000

        NORMAL_CACHE_ENTRY = &H1
        COOKIE_CACHE_ENTRY = &H100000
        URLHISTORY_CACHE_ENTRY = &H200000
    End Enum

    'https://docs.microsoft.com/en-us/windows/win32/api/wininet/ns-wininet-internet_cache_entry_infow
    'typedef struct _INTERNET_CACHE_ENTRY_INFOW {
    '  DWORD    dwStructSize;
    '  LPWSTR   lpszSourceUrlName;
    '  LPWSTR   lpszLocalFileName;
    '  DWORD    CacheEntryType;
    '  DWORD    dwUseCount;
    '  DWORD    dwHitRate;
    '  DWORD    dwSizeLow;
    '  DWORD    dwSizeHigh;
    '  FILETIME LastModifiedTime;
    '  FILETIME ExpireTime;
    '  FILETIME LastAccessTime;
    '  FILETIME LastSyncTime;
    '  LPWSTR   lpHeaderInfo;
    '  DWORD    dwHeaderInfoSize;
    '  LPWSTR   lpszFileExtension;
    '  union {
    '    DWORD dwReserved;
    '    DWORD dwExemptDelta;
    '  };
    '} INTERNET_CACHE_ENTRY_INFOW, *LPINTERNET_CACHE_ENTRY_INFOW;
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure INTERNET_CACHE_ENTRY_INFO
        Public dwStructSize As UInt32
        Public lpszSourceUrlName As IntPtr
        Public lpszLocalFileName As IntPtr
        Public CacheEntryType As CACHE_ENTRY_TYPE
        Public dwUseCount As UInt32
        Public dwHitRate As UInt32
        Public dwSizeLow As UInt32
        Public dwSizeHigh As UInt32
        Public LastModifiedTime As ComTypes.FILETIME
        Public ExpireTime As ComTypes.FILETIME
        Public LastAccessTime As ComTypes.FILETIME
        Public LastSyncTime As ComTypes.FILETIME
        Public lpHeaderInfo As IntPtr
        Public dwHeaderInfoSize As UInt32
        Public lpszFileExtension As IntPtr
        Public dwExemptDelta As UInt32
    End Structure

#Region "枚举缓存"
    'https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-findfirsturlcacheentryw
    'void FindFirstUrlCacheEntryW(
    '  LPCWSTR                      lpszUrlSearchPattern,
    '  LPINTERNET_CACHE_ENTRY_INFOW lpFirstCacheEntryInfo,
    '  LPDWORD                      lpcbCacheEntryInfo
    ');
    <DllImport("wininet.dll", CharSet:=CharSet.Unicode, EntryPoint:="FindFirstUrlCacheEntryW", SetLastError:=True)>
    Private Shared Function FindFirstUrlCacheEntry(ByVal lpszUrlSearchPattern As String, ByVal lpFirstCacheEntryInfo As IntPtr, ByRef lpcbCacheEntryInfo As UInt32) As IntPtr
    End Function

    'https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-findnexturlcacheentryw
    'BOOLAPI FindNextUrlCacheEntryW(
    '  HANDLE                       hEnumHandle,
    '  LPINTERNET_CACHE_ENTRY_INFOW lpNextCacheEntryInfo,
    '  LPDWORD                      lpcbCacheEntryInfo
    ');
    <DllImport("wininet.dll", CharSet:=CharSet.Unicode, EntryPoint:="FindNextUrlCacheEntryW", SetLastError:=True)>
    Private Shared Function FindNextUrlCacheEntry(ByVal hEnumHandle As IntPtr, ByVal lpNextCacheEntryInfo As IntPtr, ByRef lpcbCacheEntryInfo As UInt32) As Boolean
    End Function

    'https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-findcloseurlcache
    'BOOLAPI FindCloseUrlCache(
    '  HANDLE hEnumHandle
    ');
    <DllImport("wininet.dll", SetLastError:=True)>
    Private Shared Function FindCloseUrlCache(ByVal hEnumHandle As IntPtr) As Boolean
    End Function

    Private Enum CacheEnumPattern As Int32
        Visited = 1
        Cookies = 2
        Content = 0
    End Enum

    Private Shared Function GetCacheEntries(pattern As CacheEnumPattern) As List(Of INTERNET_CACHE_ENTRY_INFO)
        Dim UrlSearchPattern As String = Nothing
        Select Case pattern
            Case CacheEnumPattern.Visited
                UrlSearchPattern = "visited:"
            Case CacheEnumPattern.Cookies
                UrlSearchPattern = "cookie:"
            Case CacheEnumPattern.Content
                UrlSearchPattern = Nothing
        End Select
 
        Const ERROR_INSUFFICIENT_BUFFER As Integer = 122
        Const ERROR_NO_MORE_ITEMS As Integer = 259
        Dim ErrorCode As Integer
        Dim FindHandle As IntPtr = IntPtr.Zero
        Dim CacheEntryInfo As INTERNET_CACHE_ENTRY_INFO
        Dim CacheEntryInfoPtr As IntPtr = IntPtr.Zero
        Dim CacheEntryInfoSize As UInt32 = 0
        Dim CacheEntryInfoList As New List(Of INTERNET_CACHE_ENTRY_INFO)

        FindHandle = FindFirstUrlCacheEntry(UrlSearchPattern, CacheEntryInfoPtr, CacheEntryInfoSize)
        ErrorCode = Marshal.GetLastWin32Error
        Select Case ErrorCode
            Case ERROR_INSUFFICIENT_BUFFER
                CacheEntryInfoPtr = Marshal.AllocHGlobal(CInt(CacheEntryInfoSize))
                FindHandle = FindFirstUrlCacheEntry(UrlSearchPattern, CacheEntryInfoPtr, CacheEntryInfoSize)
                CacheEntryInfo = Marshal.PtrToStructure(CacheEntryInfoPtr, GetType(INTERNET_CACHE_ENTRY_INFO))
                CacheEntryInfoList.Add(CacheEntryInfo)
                While True
                    CacheEntryInfoPtr = IntPtr.Zero
                    CacheEntryInfoSize = 0
                    FindNextUrlCacheEntry(FindHandle, CacheEntryInfoPtr, CacheEntryInfoSize)
                    ErrorCode = Marshal.GetLastWin32Error
                    Select Case ErrorCode
                        Case ERROR_INSUFFICIENT_BUFFER
                            CacheEntryInfoPtr = Marshal.AllocHGlobal(CInt(CacheEntryInfoSize))
                            FindNextUrlCacheEntry(FindHandle, CacheEntryInfoPtr, CacheEntryInfoSize)
                            CacheEntryInfo = Marshal.PtrToStructure(CacheEntryInfoPtr, GetType(INTERNET_CACHE_ENTRY_INFO))
                            CacheEntryInfoList.Add(CacheEntryInfo)
                        Case ERROR_NO_MORE_ITEMS
                            Exit While
                        Case Else
                            Dim ex As New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error)
                            ex.Source = "FindNextUrlCacheEntry"
                            Throw ex
                    End Select
                End While
            Case ERROR_NO_MORE_ITEMS

            Case Else
                Dim ex As New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error)
                ex.Source = "FindFirstUrlCacheEntry"
                Throw ex
        End Select
        FindCloseUrlCache(FindHandle)

        Return CacheEntryInfoList
    End Function

    Public Shared Function GetCaches() As Dictionary(Of String, String)
        Dim CacheEntryInfoList As List(Of INTERNET_CACHE_ENTRY_INFO) = GetCacheEntries(CacheEnumPattern.Content)
        Dim CacheDict As New Dictionary(Of String, String)
        For Each CacheEntryInfo In CacheEntryInfoList
            Dim SourceUrl As String = Marshal.PtrToStringAuto(CacheEntryInfo.lpszSourceUrlName)
            Dim LocalFile As String = Marshal.PtrToStringAuto(CacheEntryInfo.lpszLocalFileName)
            CacheDict.Add(SourceUrl, LocalFile)
            Debug.Print("{0}|{1}", SourceUrl, LocalFile)
        Next
        Return CacheDict
    End Function

#End Region

#Region "检索缓存"
    'https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-geturlcacheentryinfow
    'BOOLAPI GetUrlCacheEntryInfoW(
    '  LPCWSTR                      lpszUrlName,
    '  LPINTERNET_CACHE_ENTRY_INFOW lpCacheEntryInfo,
    '  LPDWORD                      lpcbCacheEntryInfo
    ');
    <DllImport("wininet.dll", CharSet:=CharSet.Unicode, EntryPoint:="GetUrlCacheEntryInfoW", SetLastError:=True)>
    Private Shared Function GetUrlCacheEntryInfo(ByVal lpszUrlName As String, ByVal lpCacheEntryInfo As IntPtr, ByRef lpcbCacheEntryInfo As UInt32) As Boolean
    End Function

    Public Shared Function GetCache(ByVal url As String) As String
        Const ERROR_INSUFFICIENT_BUFFER As Integer = 122
        'Const ERROR_FILE_NOT_FOUND As Integer = 2
        Dim ErrorCode As Integer
        Dim CacheEntryInfoPtr As IntPtr = IntPtr.Zero
        Dim CacheEntryInfoSize As Integer = 0
        Dim CacheEntryInfo As INTERNET_CACHE_ENTRY_INFO = Nothing
        Dim LocalFile As String = String.Empty
        GetUrlCacheEntryInfo(url, CacheEntryInfoPtr, CacheEntryInfoSize)
        ErrorCode = Marshal.GetLastWin32Error
        If ErrorCode = ERROR_INSUFFICIENT_BUFFER Then
            CacheEntryInfoPtr = Marshal.AllocHGlobal(CacheEntryInfoSize)
            GetUrlCacheEntryInfo(url, CacheEntryInfoPtr, CacheEntryInfoSize)
            CacheEntryInfo = Marshal.PtrToStructure(CacheEntryInfoPtr, GetType(INTERNET_CACHE_ENTRY_INFO))
            LocalFile = Marshal.PtrToStringAuto(CacheEntryInfo.lpszLocalFileName)
        End If
        Return LocalFile
    End Function
#End Region

#Region "删除缓存"

    'https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-findfirsturlcachegroup
    'void FindFirstUrlCacheGroup(
    '  DWORD   dwFlags,
    '  DWORD   dwFilter,
    '  LPVOID  lpSearchCondition,
    '  DWORD   dwSearchCondition,
    '  GROUPID *lpGroupId,
    '  LPVOID  lpReserved
    ');
    <DllImport("wininet.dll", SetLastError:=True)>
    Private Shared Function FindFirstUrlCacheGroup(ByVal dwFlags As UInt32, ByVal dwFilter As UInt32, ByVal lpSearchCondition As IntPtr, ByVal dwSearchCondition As UInt32, ByRef lpGroupId As Int64, ByVal lpReserved As UInt32) As IntPtr
    End Function

    'https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-findnexturlcachegroup
    'BOOLAPI FindNextUrlCacheGroup(
    '  HANDLE  hFind,
    '  GROUPID *lpGroupId,
    '  LPVOID  lpReserved
    ');
    <DllImport("wininet.dll", SetLastError:=True)>
    Private Shared Function FindNextUrlCacheGroup(ByVal hFind As IntPtr, ByRef lpGroupId As Int64, ByVal lpReserved As UInt32) As Boolean
    End Function

    'https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-deleteurlcachegroup
    'BOOLAPI DeleteUrlCacheGroup(
    '  GROUPID GroupId,
    '  DWORD   dwFlags,
    '  LPVOID  lpReserved
    ');
    <DllImport("wininet.dll", SetLastError:=True)>
    Private Shared Function DeleteUrlCacheGroup(ByVal GroupId As Int64, ByVal dwFlags As UInt32, ByVal lpReserved As UInt32) As Boolean
    End Function

    'https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-deleteurlcacheentryw
    'BOOLAPI DeleteUrlCacheEntryW(
    '  LPCWSTR lpszUrlName
    ');
    <DllImport("wininet.dll", CharSet:=CharSet.Unicode, EntryPoint:="DeleteUrlCacheEntryW", SetLastError:=True)>
    Private Shared Function DeleteUrlCacheEntry(ByVal lpszUrlName As IntPtr) As Boolean
    End Function

    Private Shared Function GetCacheGroups() As List(Of Int64)
        Const CACHEGROUP_SEARCH_ALL As UInt32 = 0

        Dim GroupId As Int64 = -1
        Dim GroupList As New List(Of Int64)
        Dim FindHandle As IntPtr = IntPtr.Zero
        FindHandle = FindFirstUrlCacheGroup(0, CACHEGROUP_SEARCH_ALL, Nothing, 0, GroupId, Nothing)
        If GroupId >= 0 Then
            GroupList.Add(GroupId)
            While True
                If FindNextUrlCacheGroup(FindHandle, GroupId, Nothing) = True Then
                    GroupList.Add(GroupId)
                Else
                    Exit While
                End If
            End While
        End If

        Return GroupList
    End Function

    Public Shared Sub DelCaches()
        Const CACHEGROUP_FLAG_FLUSHURL_ONDELETE As UInt32 = 2

        Dim GroupList As List(Of Int64) = GetCacheGroups()
        For Each GroupId As Int64 In GroupList
            DeleteUrlCacheGroup(GroupId, CACHEGROUP_FLAG_FLUSHURL_ONDELETE, Nothing)
        Next

        Dim CacheEntryInfoList As List(Of INTERNET_CACHE_ENTRY_INFO) = GetCacheEntries(CacheEnumPattern.Content)
        For Each CacheEntryInfo As INTERNET_CACHE_ENTRY_INFO In CacheEntryInfoList
            DeleteUrlCacheEntry(CacheEntryInfo.lpszSourceUrlName)
        Next
    End Sub

#End Region

#Region "删除Cookie"
    Public Shared Sub DelCookies()
        Dim CacheEntryInfoList As List(Of INTERNET_CACHE_ENTRY_INFO) = GetCacheEntries(CacheEnumPattern.Cookies)
        For Each CacheEntryInfo In CacheEntryInfoList
            DeleteUrlCacheEntry(CacheEntryInfo.lpszSourceUrlName)
        Next
    End Sub

    Public Shared Function DelCookie(url As String) As Boolean
        Dim CacheEntryInfoList As List(Of INTERNET_CACHE_ENTRY_INFO) = GetCacheEntries(CacheEnumPattern.Cookies)
        For Each CacheEntryInfo In CacheEntryInfoList
            Dim SourceUrl As String = Marshal.PtrToStringAuto(CacheEntryInfo.lpszSourceUrlName)
            If SourceUrl.Contains("@" & url) Then
                Return DeleteUrlCacheEntry(CacheEntryInfo.lpszSourceUrlName)
            End If
        Next
        Return False
    End Function

#End Region

#Region "禁用Cookie"

    'https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetsetoptionw
    'BOOLAPI InternetSetOptionW(
    '  HINTERNET hInternet,
    '  DWORD     dwOption,
    '  LPVOID    lpBuffer,
    '  DWORD     dwBufferLength
    ');
    <DllImport("wininet.dll", CharSet:=CharSet.Unicode, EntryPoint:="InternetSetOptionW", SetLastError:=True)>
    Private Shared Function InternetSetOption(ByVal hInternet As IntPtr, ByVal dwOption As UInt32, ByVal lpBuffer As IntPtr, ByVal dwBufferLength As UInt32) As Boolean
    End Function

    Private Shared Sub SetCookieOption(enable As Boolean)
        Const INTERNET_OPTION_SUPPRESS_BEHAVIOR As UInt32 = 81
        '禁用抑制
        Const INTERNET_SUPPRESS_RESET_ALL As Int32 = 0
        ''抑制策略
        'Const INTERNET_SUPPRESS_COOKIE_POLICY As Int32 = 1
        '抑制效期
        Const INTERNET_SUPPRESS_COOKIE_PERSIST As Int32 = 3

        Dim BufferSize As Int32 = Marshal.SizeOf(GetType(UInt32))
        Dim BufferPtr As IntPtr = Marshal.AllocHGlobal(BufferSize)
        If enable = True Then
            Marshal.WriteInt32(BufferPtr, INTERNET_SUPPRESS_RESET_ALL)
        Else
            Marshal.WriteInt32(BufferPtr, INTERNET_SUPPRESS_COOKIE_PERSIST)
        End If

        If InternetSetOption(Nothing, INTERNET_OPTION_SUPPRESS_BEHAVIOR, BufferPtr, BufferSize) = False Then
            Throw New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error)
        End If
    End Sub

    Public Shared Sub EnableCookies()
        SetCookieOption(True)
    End Sub

    Public Shared Sub DisableCookies()
        SetCookieOption(False)
    End Sub
#End Region



    Public Function GetElementById(id As String) As HtmlElement
        If MyBase.Document Is Nothing Then
            Return Nothing
        End If

        Return MyBase.Document.GetElementById(id)
    End Function

    Public Function GetElementByAttribute(tagName As String, attrName As String, attrValue As String) As HtmlElement
        If MyBase.Document Is Nothing Then
            Return Nothing
        End If
        'class -> className
        Dim Elements As HtmlElementCollection = MyBase.Document.GetElementsByTagName(tagName)
        Dim Element As HtmlElement = Nothing
        For Each ElementTemp As HtmlElement In Elements
            Dim Attribute As String = ElementTemp.GetAttribute(attrName)
            If Attribute = attrValue Then
                Element = ElementTemp
                Exit For
            End If
        Next
        Return Element
    End Function

    ''' <summary>
    ''' 使用DoEvents延时。每经过msStep毫秒测试一次break函数，直至break函数返回True或者累积毫秒超过msStop。
    ''' </summary>
    ''' <param name="break"></param>
    ''' <param name="msStep"></param>
    ''' <param name="msStop"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function Wait(break As Func(Of Boolean), msStep As Integer, msStop As Integer) As Boolean
        Dim WaitResult As Boolean = False
        Dim TimeBegin As Date = Date.Now
        Dim StepBegin As Date = Date.Now
        While True
            Dim TimeNow As Date = Date.Now
            Dim TimeUse As Integer = TimeNow.Subtract(TimeBegin).TotalMilliseconds
            If TimeUse >= msStop Then
                Exit While
            End If
            Dim StepUse As Integer = TimeNow.Subtract(StepBegin).TotalMilliseconds
            If StepUse < msStep Then
                Application.DoEvents()
                Continue While
            Else
                If break.Invoke = True Then
                    WaitResult = True
                    Exit While
                Else
                    StepBegin = Date.Now
                End If
            End If
        End While
        Return WaitResult
    End Function

    Public Function WaitForKeyword(keyword As String, seconds As Integer) As Boolean
        Dim Break = Function()
                        'If MyBase.IsBusy Then
                        '    Return False
                        'End If
                        'If MyBase.ReadyState < WebBrowserReadyState.Complete Then
                        '    Return False
                        'End If
                        Dim Html As String = Me.HtmlText
                        If String.IsNullOrEmpty(Html) Then
                            Return False
                        End If
                        If Html.Contains(keyword) Then
                            Return True
                        Else
                            Return False
                        End If
                    End Function
        Return Wait(Break, 500, seconds * 1000)
    End Function

    Public Function WaitForKeywordAny(keywords As String(), seconds As Integer) As Boolean
        Dim Break = Function()
                        'If MyBase.IsBusy Then
                        '    Return False
                        'End If
                        'If MyBase.ReadyState < WebBrowserReadyState.Complete Then
                        '    Return False
                        'End If
                        Dim Html As String = Me.HtmlText
                        If String.IsNullOrEmpty(Html) Then
                            Return False
                        End If
                        For Each keyword As String In keywords
                            If Html.Contains(keyword) Then
                                Return True
                            End If
                        Next
                        Return False
                    End Function
        Return Wait(Break, 500, seconds * 1000)
    End Function

    Public Function WaitForNavigate(url As String, seconds As Integer) As Boolean
        Me.UrlCompleted = String.Empty
        MyBase.Navigate(url)

        Dim Break = Function()
                        If String.IsNullOrEmpty(Me.UrlCompleted) Then
                            Return False
                        End If
                        If Me.UrlCompleted = url Then
                            Return True
                        Else
                            Return False
                        End If
                    End Function

        Return Wait(Break, 500, seconds * 1000)
    End Function


End Class
