Imports System.Runtime.InteropServices

Public Class ShellLink
    'https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-shllink/16cb4ca1-9339-4d0c-a68d-bf1d6cc0f943

    Private Shared ReadOnly NullTerminate As String = Char.MinValue.ToString

    Private Shared Function GetObjectString(Obj As Object) As String
        Dim Builder As New System.Text.StringBuilder
        Dim ObjType As System.Type = Obj.GetType
        Builder.AppendLine()
        Builder.AppendLine(ObjType.Name)
        For Each Field As System.Reflection.FieldInfo In ObjType.GetFields
            Dim FieldName As String = Field.Name
            Dim FieldValue As Object = Field.GetValue(Obj)
            Dim FieldValueString As String
            If FieldValue Is Nothing Then
                FieldValueString = "Nothing"
            Else
                FieldValueString = FieldValue.ToString
            End If
            Builder.AppendFormat("{0}:{1}", FieldName, FieldValueString)
            Builder.AppendLine()
        Next
        Return Builder.ToString
    End Function
 
    Private Enum SW_SHOW As UInt32
        SW_SHOWNORMAL = &H1
        SW_SHOWMAXIMIZED = &H3
        SW_SHOWMINNOACTIVE = &H7
    End Enum

    <Flags>
    Private Enum LinkFlag As UInt32
        HasLinkTargetIDList = 1UI
        HasLinkInfo = 1UI << 1
        HasName = 1UI << 2
        HasRelativePath = 1UI << 3
        HasWorkingDir = 1UI << 4
        HasArguments = 1UI << 5
        HasIconLocation = 1UI << 6
        IsUnicode = 1UI << 7
        ForceNoLinkInfo = 1UI << 8
        HasExpString = 1UI << 9
        RunInSeparateProcess = 1UI << 10
        Unused1 = 1UI << 11
        HasDarwinID = 1UI << 12
        RunAsUser = 1UI << 13
        HasExpIcon = 1UI << 14
        NoPidlAlias = 1UI << 15
        Unused2 = 1UI << 16
        RunWithShimLayer = 1UI << 17
        ForceNoLinkTrack = 1UI << 18
        EnableTargetMetadata = 1UI << 19
        DisableLinkPathTracking = 1UI << 20
        DisableKnownFolderTracking = 1UI << 21
        DisableKnownFolderAlias = 1UI << 22
        AllowLinkToLink = 1UI << 23
        UnaliasOnSave = 1UI << 24
        PreferEnvironmentPath = 1UI << 25
        KeepLocalIDListForUNCTarget = 1UI << 26
    End Enum

    <Flags>
    Private Enum FileAttribute As UInt32
        FILE_ATTRIBUTE_READONLY = 1UI
        FILE_ATTRIBUTE_HIDDEN = 1UI << 1
        FILE_ATTRIBUTE_SYSTEM = 1UI << 2
        Reserved1 = 1UI << 3
        FILE_ATTRIBUTE_DIRECTORY = 1UI << 4
        FILE_ATTRIBUTE_ARCHIVE = 1UI << 5
        Reserved2 = 1UI << 6
        FILE_ATTRIBUTE_NORMAL = 1UI << 7
        FILE_ATTRIBUTE_TEMPORARY = 1UI << 8
        FILE_ATTRIBUTE_SPARSE_FILE = 1UI << 9
        FILE_ATTRIBUTE_REPARSE_POINT = 1UI << 10
        FILE_ATTRIBUTE_COMPRESSED = 1UI << 11
        FILE_ATTRIBUTE_OFFLINE = 1UI << 12
        FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 1UI << 13
        FILE_ATTRIBUTE_ENCRYPTED = 1UI << 14
    End Enum

    <Flags>
    Private Enum HotKeyFlag As UInt16
        VK_0 = &H30US
        VK_1 = &H31US
        VK_2 = &H32US
        VK_3 = &H33US
        VK_4 = &H34US
        VK_5 = &H35US
        VK_6 = &H36US
        VK_7 = &H37US
        VK_8 = &H38US
        VK_9 = &H39US
        VK_A = &H41US
        VK_B = &H42US
        VK_C = &H43US
        VK_D = &H44US
        VK_E = &H45US
        VK_F = &H46US
        VK_G = &H47US
        VK_H = &H48US
        VK_I = &H49US
        VK_J = &H4AUS
        VK_K = &H4BUS
        VK_L = &H4CUS
        VK_M = &H4DUS
        VK_N = &H4EUS
        VK_O = &H4FUS
        VK_P = &H50US
        VK_Q = &H51US
        VK_R = &H52US
        VK_S = &H53US
        VK_T = &H54US
        VK_U = &H55US
        VK_V = &H56US
        VK_W = &H57US
        VK_X = &H58US
        VK_Y = &H59US
        VK_Z = &H5AUS

        VK_F1 = &H70US
        VK_F2 = &H71US
        VK_F3 = &H72US
        VK_F4 = &H73US
        VK_F5 = &H74US
        VK_F6 = &H75US
        VK_F7 = &H76US
        VK_F8 = &H77US
        VK_F9 = &H78US
        VK_F10 = &H79US
        VK_F11 = &H7AUS
        VK_F12 = &H7BUS
        VK_F13 = &H7CUS
        VK_F14 = &H7DUS
        VK_F15 = &H7EUS
        VK_F16 = &H7FUS
        VK_F17 = &H80US
        VK_F18 = &H81US
        VK_F19 = &H82US
        VK_F20 = &H83US
        VK_F21 = &H84US
        VK_F22 = &H85US
        VK_F23 = &H86US
        VK_F24 = &H87US
        VK_NUMLOCK = &H90US
        VK_SCROLL = &H91US

        HOTKEYF_SHIFT = &H100US
        HOTKEYF_CONTROL = &H200US
        HOTKEYF_ALT = &H400US
    End Enum

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure ShellLinkHeader
        Public HeaderSize As UInt32
        Public LinkCLSID As Guid
        Public LinkFlags As LinkFlag
        Public FileAttributes As FileAttribute
        Public CreationTime As Date
        Public AccessTime As Date
        Public WriteTime As Date
        Public FileSize As UInt32
        Public IconIndex As Int32
        Public ShowCommand As SW_SHOW
        Public HotKey As HotKeyFlag
        Public Reserved1 As UInt16
        Public Reserved2 As UInt32
        Public Reserved3 As UInt32

        Public Sub Read(Stream As System.IO.Stream)
            Stream.Read(Me.HeaderSize)
            Stream.Read(Me.LinkCLSID)
            Stream.ReadEnum(Me.LinkFlags)
            Stream.ReadEnum(Me.FileAttributes)
            Stream.Read(Me.CreationTime)
            Stream.Read(Me.AccessTime)
            Stream.Read(Me.WriteTime)
            Stream.Read(Me.FileSize)
            Stream.Read(Me.IconIndex)
            Stream.ReadEnum(Me.ShowCommand)
            Stream.ReadEnum(Me.HotKey)
            Stream.Read(Me.Reserved1)
            Stream.Read(Me.Reserved2)
            Stream.Read(Me.Reserved3)
        End Sub

        Public Sub Write(Stream As System.IO.Stream)
            Me.HeaderSize = &H4C
            Me.LinkCLSID = New Guid("00021401-0000-0000-C000-000000000046")
            Stream.Write(Me.HeaderSize)
            Stream.Write(Me.LinkCLSID)
            Stream.Write(Me.LinkFlags)
            Stream.Write(Me.FileAttributes)
            Stream.Write(Me.CreationTime)
            Stream.Write(Me.AccessTime)
            Stream.Write(Me.WriteTime)
            Stream.Write(Me.FileSize)
            Stream.Write(Me.IconIndex)
            Stream.Write(Me.ShowCommand)
            Stream.Write(Me.HotKey)
            Stream.Write(Me.Reserved1)
            Stream.Write(Me.Reserved2)
            Stream.Write(Me.Reserved3)
        End Sub
    End Structure

#Region "IdList与Path转换API"
    'typedef struct _SHITEMID {
    '  USHORT cb;
    '  BYTE   abID[1];
    '} SHITEMID;
    'typedef struct _ITEMIDLIST {
    '  SHITEMID mkid;
    '} ITEMIDLIST;
    Private Structure ITEMIDLIST
        Public mkid As IntPtr
    End Structure

    'SHSTDAPI SHParseDisplayName(
    '  PCWSTR           pszName,
    '  IBindCtx         *pbc,
    '  PIDLIST_ABSOLUTE *ppidl,
    '  SFGAOF           sfgaoIn,
    '  SFGAOF           *psfgaoOut
    ');
    <DllImport("Shell32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function SHParseDisplayName(ByVal pszName As String, ByVal pbc As IntPtr, ByRef ppidl As ITEMIDLIST, ByVal sfgaoIn As UInt32, ByRef psfgaoOut As UInt32) As Int32
    End Function

    'BOOL SHGetPathFromIDListW(
    '  PCIDLIST_ABSOLUTE pidl,
    '  LPWSTR            pszPath
    ');
    <DllImport("Shell32.dll", CharSet:=CharSet.Unicode, EntryPoint:="SHGetPathFromIDListW", SetLastError:=True)> _
    Private Shared Function SHGetPathFromIDList(ByVal pidl As ITEMIDLIST, ByVal pszPath As System.Text.StringBuilder) As Boolean
    End Function
#End Region

    Private Structure LinkTargetIDList
        Public IDListSize As UInt16
        Public ItemIDList As Byte()

        Public Sub Read(Stream As System.IO.Stream)
            Stream.Read(Me.IDListSize)
            ReDim Me.ItemIDList(Me.IDListSize - 1)
            Stream.Read(Me.ItemIDList)
        End Sub

        Public Sub Prepare()
            If Me.ItemIDList IsNot Nothing Then
                Me.IDListSize = Me.ItemIDList.Length
            Else
                Me.IDListSize = 0
            End If
        End Sub

        Public Sub Write(Stream As System.IO.Stream)
            Stream.Write(Me.IDListSize)
            Stream.Write(Me.ItemIDList)
        End Sub

        Public Property Path As String
            Get
                If Me.ItemIDList Is Nothing Then
                    Return Nothing
                Else
                    Me.IDListSize = Me.ItemIDList.Length
                End If

                Dim mkid As IntPtr = Marshal.AllocHGlobal(Me.IDListSize)
                Marshal.Copy(Me.ItemIDList, 0, mkid, Me.ItemIDList.Length)
                Dim pidl As ITEMIDLIST
                pidl.mkid = mkid

                Dim PathBuilder As New System.Text.StringBuilder(Me.IDListSize)
                If SHGetPathFromIDList(pidl, PathBuilder) = True Then
                    Return PathBuilder.ToString
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                If System.IO.File.Exists(value) = False AndAlso System.IO.Directory.Exists(value) = False Then
                    Throw New System.IO.FileNotFoundException()
                End If
                Dim pidl As ITEMIDLIST
                If SHParseDisplayName(value, Nothing, pidl, 0, 0) = 0 Then
                    Dim ItemOffset As Integer = 0
                    Dim ItemLength As Integer
                    Dim Size As Integer = 0
                    While True
                        ItemLength = Marshal.ReadInt16(IntPtr.Add(pidl.mkid, ItemOffset))
                        If ItemLength = 0 Then
                            Size = Size + 2
                            Exit While
                        Else
                            Size = Size + ItemLength
                            ItemOffset = ItemOffset + ItemLength
                        End If
                    End While

                    Me.IDListSize = Size
                    ReDim Me.ItemIDList(Size - 1)
                    Marshal.Copy(pidl.mkid, Me.ItemIDList, 0, Size)
                End If
            End Set
        End Property
    End Structure

    Private Enum LinkInfoFlag As UInt32
        VolumeIDAndLocalBasePath = 1
        CommonNetworkRelativeLinkAndPathSuffix = 2
    End Enum

    Private Enum DRIVE_TYPE As UInt32
        DRIVE_UNKNOWN = &H0
        DRIVE_NO_ROOT_DIR = &H1
        DRIVE_REMOVABLE = &H2
        DRIVE_FIXED = &H3
        DRIVE_REMOTE = &H4
        DRIVE_CDROM = &H5
        DRIVE_RAMDISK = &H6
    End Enum

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure VolumeID
        Public VolumeIDSize As UInt32
        Public DriveType As DRIVE_TYPE
        Public DriveSerialNumber As UInt32
        Public VolumeLabelOffset As UInt32
        Public VolumeLabelOffsetUnicode As UInt32 'Optional
        Public VolumeLabel As String
        Public Data As Byte()

        Public Sub Read(Stream As System.IO.Stream)
            Dim BaseOffset As Int64 = Stream.Position
            Stream.Read(Me.VolumeIDSize)
            Stream.ReadEnum(Me.DriveType)
            Stream.Read(Me.DriveSerialNumber)
            Stream.Read(Me.VolumeLabelOffset)

            If Me.VolumeLabelOffset = &H14 Then
                Stream.Read(Me.VolumeLabelOffsetUnicode)
                ReDim Me.Data(Me.VolumeIDSize - &H14)
                Stream.Read(Me.Data)
                Stream.Position = BaseOffset + Me.VolumeLabelOffsetUnicode
                Stream.Read(Me.VolumeLabel, StreamExtension.ReadStringOption.BreakType_NullTerminate Or StreamExtension.ReadStringOption.Encoding_Unicode)
            Else
                ReDim Me.Data(Me.VolumeIDSize - &H10)
                Stream.Read(Me.Data)
                Stream.Position = BaseOffset + Me.VolumeLabelOffset
                Stream.Read(Me.VolumeLabel, StreamExtension.ReadStringOption.BreakType_NullTerminate Or StreamExtension.ReadStringOption.Encoding_Default)
            End If
            Stream.Position = BaseOffset + VolumeIDSize
        End Sub

        Public Sub Prepare()
            Me.VolumeLabelOffset = &H10
            'Me.VolumeLabelOffsetUnicode = &H14
            Me.Data = System.Text.Encoding.Default.GetBytes(Me.VolumeLabel & NullTerminate)
            Me.VolumeIDSize = Me.VolumeLabelOffset + Me.Data.Length
        End Sub

        Public Sub Write(Stream As System.IO.Stream)
            Stream.Write(Me.VolumeIDSize)
            Stream.Write(Me.DriveType)
            Stream.Write(Me.DriveSerialNumber)
            Stream.Write(Me.VolumeLabelOffset)
            'Stream.Write(Me.VolumeLabelOffsetUnicode)
            Stream.Write(Me.Data)
        End Sub







    End Structure

    <Flags>
    Private Enum CommonNetworkRelativeLinkFlag As UInt32
        ValidDevice = 1
        ValidNetType = 2
    End Enum

    Private Enum WNNC_NET As UInt32
        WNNC_NET_AVID = &H1A0000
        WNNC_NET_DOCUSPACE = &H1B0000
        WNNC_NET_MANGOSOFT = &H1C0000
        WNNC_NET_SERNET = &H1D0000
        WNNC_NET_RIVERFRONT1 = &H1E0000
        WNNC_NET_RIVERFRONT2 = &H1F0000
        WNNC_NET_DECORB = &H200000
        WNNC_NET_PROTSTOR = &H210000
        WNNC_NET_FJ_REDIR = &H220000
        WNNC_NET_DISTINCT = &H230000
        WNNC_NET_TWINS = &H240000
        WNNC_NET_RDR2SAMPLE = &H250000
        WNNC_NET_CSC = &H260000
        WNNC_NET_3IN1 = &H270000
        WNNC_NET_EXTENDNET = &H290000
        WNNC_NET_STAC = &H2A0000
        WNNC_NET_FOXBAT = &H2B0000
        WNNC_NET_YAHOO = &H2C0000
        WNNC_NET_EXIFS = &H2D0000
        WNNC_NET_DAV = &H2E0000
        WNNC_NET_KNOWARE = &H2F0000
        WNNC_NET_OBJECT_DIRE = &H300000
        WNNC_NET_MASFAX = &H310000
        WNNC_NET_HOB_NFS = &H320000
        WNNC_NET_SHIVA = &H330000
        WNNC_NET_IBMAL = &H340000
        WNNC_NET_LOCK = &H350000
        WNNC_NET_TERMSRV = &H360000
        WNNC_NET_SRT = &H370000
        WNNC_NET_QUINCY = &H380000
        WNNC_NET_OPENAFS = &H390000
        WNNC_NET_AVID1 = &H3A0000
        WNNC_NET_DFS = &H3B0000
        WNNC_NET_KWNP = &H3C0000
        WNNC_NET_ZENWORKS = &H3D0000
        WNNC_NET_DRIVEONWEB = &H3E0000
        WNNC_NET_VMWARE = &H3F0000
        WNNC_NET_RSFX = &H400000
        WNNC_NET_MFILES = &H410000
        WNNC_NET_MS_NFS = &H420000
        WNNC_NET_GOOGLE = &H430000
    End Enum

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure CommonNetworkRelativeLink
        Public CommonNetworkRelativeLinkSize As UInt32
        Public CommonNetworkRelativeLinkFlags As CommonNetworkRelativeLinkFlag
        Public NetNameOffset As UInt32
        Public DeviceNameOffset As UInt32
        Public NetworkProviderType As WNNC_NET

        Public NetNameOffsetUnicode As UInt32
        Public DeviceNameOffsetUnicode As UInt32
        Public NetName As String
        Public DeviceName As String

        Public Sub Read(Stream As System.IO.Stream)
            Dim BaseOffset As Int64 = Stream.Position
            Stream.Read(Me.CommonNetworkRelativeLinkSize)
            Stream.ReadEnum(Me.CommonNetworkRelativeLinkFlags)
            Stream.Read(Me.NetNameOffset)
            Stream.Read(Me.DeviceNameOffset)
            Stream.ReadEnum(Me.NetworkProviderType)
            If NetNameOffset > &H14 Then
                Stream.Read(Me.NetNameOffsetUnicode)
                Stream.Read(Me.DeviceNameOffsetUnicode)
            End If
            If NetNameOffset = &H14 Then
                Stream.Position = BaseOffset + NetNameOffset
                Stream.Read(Me.NetName, StreamExtension.ReadStringOption.BreakType_NullTerminate Or StreamExtension.ReadStringOption.Encoding_Default)
            End If
            If NetNameOffsetUnicode > 0 Then
                Stream.Position = BaseOffset + NetNameOffsetUnicode
                Stream.Read(Me.NetName, StreamExtension.ReadStringOption.BreakType_NullTerminate Or StreamExtension.ReadStringOption.Encoding_Unicode)
            End If
            If DeviceNameOffset > 0 Then
                Stream.Position = BaseOffset + DeviceNameOffset
                Stream.Read(Me.DeviceName, StreamExtension.ReadStringOption.BreakType_NullTerminate Or StreamExtension.ReadStringOption.Encoding_Default)
            End If
            If DeviceNameOffsetUnicode > 0 Then
                Stream.Position = BaseOffset + DeviceNameOffsetUnicode
                Stream.Read(Me.DeviceName, StreamExtension.ReadStringOption.BreakType_NullTerminate Or StreamExtension.ReadStringOption.Encoding_Unicode)
            End If
            Stream.Position = BaseOffset + CommonNetworkRelativeLinkSize
        End Sub

        Public Sub Prepare()
            Me.NetNameOffset = &H14
            Dim NetNameSize As Integer = If(String.IsNullOrEmpty(Me.NetName), 0, System.Text.Encoding.Default.GetByteCount(Me.NetName & NullTerminate))
            Dim DeviceNameSize As Integer = If(String.IsNullOrEmpty(Me.DeviceName), 0, System.Text.Encoding.Default.GetByteCount(Me.DeviceName & NullTerminate))
            If DeviceNameSize > 0 Then
                Me.DeviceNameOffset = Me.NetNameOffset + NetNameSize
            End If
            Me.CommonNetworkRelativeLinkSize = Me.NetNameOffset + NetNameSize + DeviceNameSize
        End Sub

        Public Sub Write(Stream As System.IO.Stream)
            Stream.Write(Me.CommonNetworkRelativeLinkSize)
            Stream.Write(Me.CommonNetworkRelativeLinkFlags)
            Stream.Write(Me.NetNameOffset)
            Stream.Write(Me.DeviceNameOffset)
            Stream.Write(Me.NetworkProviderType)
            'Stream.Write(Me.NetNameOffsetUnicode)
            'Stream.Write(Me.DeviceNameOffsetUnicode)
            If String.IsNullOrEmpty(Me.NetName) = False Then
                Stream.Write(Me.NetName & NullTerminate, System.Text.Encoding.Default)
            End If
            If String.IsNullOrEmpty(Me.DeviceName) = False Then
                Stream.Write(Me.DeviceName & NullTerminate, System.Text.Encoding.Default)
            End If
        End Sub
    End Structure

    Private Structure LinkInfo
        Public LinkInfoSize As UInt32
        Public LinkInfoHeaderSize As UInt32
        Public LinkInfoFlags As LinkInfoFlag
        Public VolumeIDOffset As UInt32
        Public LocalBasePathOffset As UInt32
        Public CommonNetworkRelativeLinkOffset As UInt32
        Public CommonPathSuffixOffset As UInt32
        Public LocalBasePathOffsetUnicode As UInt32 'Optional
        Public CommonPathSuffixOffsetUnicode As UInt32 'Optional
        Public VolumeID As VolumeID
        Public LocalBasePath As String
        Public CommonNetworkRelativeLink As CommonNetworkRelativeLink
        Public CommonPathSuffix As String

        Public Sub Read(Stream As System.IO.Stream)
            Dim BaseOffset As Int64 = Stream.Position
            Stream.Read(Me.LinkInfoSize)
            Stream.Read(Me.LinkInfoHeaderSize)
            Stream.ReadEnum(Me.LinkInfoFlags)
            Stream.Read(Me.VolumeIDOffset)
            Stream.Read(Me.LocalBasePathOffset)
            Stream.Read(Me.CommonNetworkRelativeLinkOffset)
            Stream.Read(Me.CommonPathSuffixOffset)
            If LinkInfoHeaderSize >= &H24 Then
                Stream.Read(Me.LocalBasePathOffsetUnicode)
                Stream.Read(Me.CommonPathSuffixOffsetUnicode)
            End If
            If LinkInfoFlags.HasFlag(LinkInfoFlag.VolumeIDAndLocalBasePath) Then
                Stream.Position = BaseOffset + VolumeIDOffset
                VolumeID = New VolumeID
                VolumeID.Read(Stream)
                'Debug.Print(GetObjectString(VolumeID))
            End If
            If LinkInfoFlags.HasFlag(LinkInfoFlag.CommonNetworkRelativeLinkAndPathSuffix) Then
                Stream.Position = BaseOffset + CommonNetworkRelativeLinkOffset
                CommonNetworkRelativeLink = New CommonNetworkRelativeLink
                CommonNetworkRelativeLink.Read(Stream)
                'Debug.Print(GetObjectString(CommonNetworkRelativeLink))
            End If
            If LocalBasePathOffset > 0 Then
                Stream.Position = BaseOffset + LocalBasePathOffset
                Stream.Read(Me.LocalBasePath, StreamExtension.ReadStringOption.BreakType_NullTerminate Or StreamExtension.ReadStringOption.Encoding_Default)
            End If
            If LocalBasePathOffsetUnicode > 0 Then
                Stream.Position = BaseOffset + LocalBasePathOffsetUnicode
                Stream.Read(Me.LocalBasePath, StreamExtension.ReadStringOption.BreakType_NullTerminate Or StreamExtension.ReadStringOption.Encoding_Unicode)
            End If
            If CommonPathSuffixOffset > 0 Then
                Stream.Position = BaseOffset + CommonPathSuffixOffset
                Stream.Read(Me.CommonPathSuffix, StreamExtension.ReadStringOption.BreakType_NullTerminate Or StreamExtension.ReadStringOption.Encoding_Default)
            End If
            If CommonPathSuffixOffsetUnicode > 0 Then
                Stream.Position = BaseOffset + CommonPathSuffixOffsetUnicode
                Stream.Read(Me.CommonPathSuffix, StreamExtension.ReadStringOption.BreakType_NullTerminate Or StreamExtension.ReadStringOption.Encoding_Unicode)
            End If
            Stream.Position = BaseOffset + LinkInfoSize
        End Sub

        Public Sub Prepare()
            Me.LinkInfoHeaderSize = &H1C
            If String.IsNullOrEmpty(Me.LocalBasePath) = False AndAlso System.Text.Encoding.Default.GetString(System.Text.Encoding.Default.GetBytes(Me.LocalBasePath)) <> Me.LocalBasePath Then
                Me.LinkInfoHeaderSize = &H24
            End If
            If String.IsNullOrEmpty(Me.CommonPathSuffix) = False AndAlso System.Text.Encoding.Default.GetString(System.Text.Encoding.Default.GetBytes(Me.CommonPathSuffix)) <> Me.CommonPathSuffix Then
                Me.LinkInfoHeaderSize = &H24
            End If

            If String.IsNullOrEmpty(Me.CommonPathSuffix) Then
                Me.CommonPathSuffix = String.Empty
            End If

            If Me.LinkInfoHeaderSize = &H24 Then
                Select Case Me.LinkInfoFlags
                    Case LinkInfoFlag.VolumeIDAndLocalBasePath
                        Me.VolumeID.Prepare()
                        Me.VolumeIDOffset = Me.LinkInfoHeaderSize
                        Me.LocalBasePathOffset = Me.LinkInfoHeaderSize + Me.VolumeID.VolumeIDSize
                        Dim LocalBasePathSize As Integer = System.Text.Encoding.Default.GetByteCount(Me.LocalBasePath & NullTerminate)
                        Me.CommonPathSuffixOffset = Me.LinkInfoHeaderSize + Me.VolumeID.VolumeIDSize + LocalBasePathSize
                        Dim CommonPathSuffixSize As Integer = System.Text.Encoding.Default.GetByteCount(Me.CommonPathSuffix & NullTerminate)
                        Me.LocalBasePathOffsetUnicode = Me.LinkInfoHeaderSize + Me.VolumeID.VolumeIDSize + LocalBasePathSize + CommonPathSuffixSize
                        Dim LocalBasePathSizeUnicode As Integer = System.Text.Encoding.Unicode.GetByteCount(Me.LocalBasePath & NullTerminate)
                        Me.CommonPathSuffixOffsetUnicode = Me.LinkInfoHeaderSize + Me.VolumeID.VolumeIDSize + LocalBasePathSize + CommonPathSuffixSize + LocalBasePathSizeUnicode
                        Dim CommonPathSuffixSizeUnicode As Integer = System.Text.Encoding.Unicode.GetByteCount(Me.CommonPathSuffix & NullTerminate)
                        Me.LinkInfoSize = Me.LinkInfoHeaderSize + Me.VolumeID.VolumeIDSize + LocalBasePathSize + CommonPathSuffixSize + LocalBasePathSizeUnicode + CommonPathSuffixSizeUnicode
                    Case LinkInfoFlag.CommonNetworkRelativeLinkAndPathSuffix
                        Me.CommonNetworkRelativeLink.Prepare()
                        Me.CommonNetworkRelativeLinkOffset = Me.LinkInfoHeaderSize
                        Me.CommonPathSuffixOffset = Me.LinkInfoHeaderSize + Me.CommonNetworkRelativeLink.CommonNetworkRelativeLinkSize
                        Dim CommonPathSuffixSize As Integer = System.Text.Encoding.Default.GetByteCount(Me.CommonPathSuffix & NullTerminate)
                        Me.CommonPathSuffixOffsetUnicode = Me.LinkInfoHeaderSize + Me.CommonNetworkRelativeLink.CommonNetworkRelativeLinkSize + CommonPathSuffixSize
                        Dim CommonPathSuffixSizeUnicode As Integer = System.Text.Encoding.Unicode.GetByteCount(Me.CommonPathSuffix & NullTerminate)
                        Me.LinkInfoSize = Me.LinkInfoHeaderSize + Me.CommonNetworkRelativeLink.CommonNetworkRelativeLinkSize + CommonPathSuffixSize + CommonPathSuffixSizeUnicode
                    Case Else
                        Throw New Exception("LinkInfo.LinkInfoFlags错误。")
                End Select
            Else
                Select Case Me.LinkInfoFlags
                    Case LinkInfoFlag.VolumeIDAndLocalBasePath
                        Me.VolumeID.Prepare()
                        Me.VolumeIDOffset = Me.LinkInfoHeaderSize
                        Me.LocalBasePathOffset = Me.LinkInfoHeaderSize + Me.VolumeID.VolumeIDSize
                        Dim LocalBasePathSize As Integer = System.Text.Encoding.Default.GetByteCount(Me.LocalBasePath & NullTerminate)
                        Me.CommonPathSuffixOffset = Me.LinkInfoHeaderSize + Me.VolumeID.VolumeIDSize + LocalBasePathSize
                        Dim CommonPathSuffixSize As Integer = System.Text.Encoding.Default.GetByteCount(Me.CommonPathSuffix & NullTerminate)
                        Me.LinkInfoSize = Me.LinkInfoHeaderSize + Me.VolumeID.VolumeIDSize + LocalBasePathSize + CommonPathSuffixSize
                    Case LinkInfoFlag.CommonNetworkRelativeLinkAndPathSuffix
                        Me.CommonNetworkRelativeLink.Prepare()
                        Me.CommonNetworkRelativeLinkOffset = Me.LinkInfoHeaderSize
                        Me.CommonPathSuffixOffset = Me.LinkInfoHeaderSize + Me.CommonNetworkRelativeLink.CommonNetworkRelativeLinkSize
                        Dim CommonPathSuffixSize As Integer = System.Text.Encoding.Default.GetByteCount(Me.CommonPathSuffix & NullTerminate)
                        Me.LinkInfoSize = Me.LinkInfoHeaderSize + Me.CommonNetworkRelativeLink.CommonNetworkRelativeLinkSize + CommonPathSuffixSize
                    Case Else
                        Throw New Exception("LinkInfo.LinkInfoFlags错误。")
                End Select
            End If
        End Sub

        Public Sub Write(Stream As System.IO.MemoryStream)
            Stream.Write(Me.LinkInfoSize)
            Stream.Write(Me.LinkInfoHeaderSize)
            Stream.Write(Me.LinkInfoFlags)
            Stream.Write(Me.VolumeIDOffset)
            Stream.Write(Me.LocalBasePathOffset)
            Stream.Write(Me.CommonNetworkRelativeLinkOffset)
            Stream.Write(Me.CommonPathSuffixOffset)
            If Me.LinkInfoHeaderSize = &H1C Then
                'Stream.Write(Me.LocalBasePathOffsetUnicode)
                'Stream.Write(Me.CommonPathSuffixOffsetUnicode)
                Select Case Me.LinkInfoFlags
                    Case LinkInfoFlag.VolumeIDAndLocalBasePath
                        Me.VolumeID.Write(Stream)
                        Stream.Write(Me.LocalBasePath & NullTerminate, System.Text.Encoding.Default)
                        Stream.Write(Me.CommonPathSuffix & NullTerminate, System.Text.Encoding.Default)
                    Case LinkInfoFlag.CommonNetworkRelativeLinkAndPathSuffix
                        Me.CommonNetworkRelativeLink.Write(Stream)
                        Stream.Write(Me.CommonPathSuffix & NullTerminate, System.Text.Encoding.Default)
                End Select
            Else
                Stream.Write(Me.LocalBasePathOffsetUnicode)
                Stream.Write(Me.CommonPathSuffixOffsetUnicode)
                Select Case Me.LinkInfoFlags
                    Case LinkInfoFlag.VolumeIDAndLocalBasePath
                        Me.VolumeID.Write(Stream)
                        Stream.Write(Me.LocalBasePath & NullTerminate, System.Text.Encoding.Default)
                        Stream.Write(Me.CommonPathSuffix & NullTerminate, System.Text.Encoding.Default)
                        Stream.Write(Me.LocalBasePath & NullTerminate, System.Text.Encoding.Unicode)
                        Stream.Write(Me.CommonPathSuffix & NullTerminate, System.Text.Encoding.Unicode)
                    Case LinkInfoFlag.CommonNetworkRelativeLinkAndPathSuffix
                        Me.CommonNetworkRelativeLink.Write(Stream)
                        Stream.Write(Me.CommonPathSuffix & NullTerminate, System.Text.Encoding.Default)
                        Stream.Write(Me.CommonPathSuffix & NullTerminate, System.Text.Encoding.Unicode)
                End Select
            End If
        End Sub
    End Structure

    Private Structure STRING_DATA
        Public NAME_STRING As String
        Public RELATIVE_PATH As String
        Public WORKING_DIR As String
        Public COMMAND_LINE_ARGUMENTS As String
        Public ICON_LOCATION As String

        Private Function ReadString(Stream As System.IO.Stream, Options As StreamExtension.ReadStringOption) As String
            Dim Length As UInt16
            Stream.Read(Length)
            Dim Text As String = Nothing
            Stream.Read(Text, Options + Length)
            Return Text
        End Function

        Public Sub Read(Stream As System.IO.Stream, LinkFlags As LinkFlag)
            Dim Options As StreamExtension.ReadStringOption = StreamExtension.ReadStringOption.BreakType_FixedChars
            If LinkFlags.HasFlag(LinkFlag.IsUnicode) Then
                Options = Options Or StreamExtension.ReadStringOption.Encoding_Unicode
            Else
                Options = Options Or StreamExtension.ReadStringOption.Encoding_Default
            End If
 
            If LinkFlags.HasFlag(LinkFlag.HasName) Then
                NAME_STRING = ReadString(Stream, Options)
            End If
            If LinkFlags.HasFlag(LinkFlag.HasRelativePath) Then
                RELATIVE_PATH = ReadString(Stream, Options)
            End If
            If LinkFlags.HasFlag(LinkFlag.HasWorkingDir) Then
                WORKING_DIR = ReadString(Stream, Options)
            End If
            If LinkFlags.HasFlag(LinkFlag.HasArguments) Then
                COMMAND_LINE_ARGUMENTS = ReadString(Stream, Options)
            End If
            If LinkFlags.HasFlag(LinkFlag.HasIconLocation) Then
                ICON_LOCATION = ReadString(Stream, Options)
            End If
        End Sub

        Public Sub Write(Stream As System.IO.Stream, LinkFlags As LinkFlag)
            Dim Encoding As System.Text.Encoding = If(LinkFlags.HasFlag(LinkFlag.IsUnicode), System.Text.Encoding.Unicode, System.Text.Encoding.Default)
            If Me.NAME_STRING IsNot Nothing Then
                Stream.Write(CUShort(Me.NAME_STRING.Length))
                Stream.Write(Me.NAME_STRING, Encoding)
            End If
            If Me.RELATIVE_PATH IsNot Nothing Then
                Stream.Write(CUShort(Me.RELATIVE_PATH.Length))
                Stream.Write(Me.RELATIVE_PATH, Encoding)
            End If
            If Me.WORKING_DIR IsNot Nothing Then
                Stream.Write(CUShort(Me.WORKING_DIR.Length))
                Stream.Write(Me.WORKING_DIR, Encoding)
            End If
            If Me.COMMAND_LINE_ARGUMENTS IsNot Nothing Then
                Stream.Write(CUShort(Me.COMMAND_LINE_ARGUMENTS.Length))
                Stream.Write(Me.COMMAND_LINE_ARGUMENTS, Encoding)
            End If
            If Me.ICON_LOCATION IsNot Nothing Then
                Stream.Write(CUShort(Me.ICON_LOCATION.Length))
                Stream.Write(Me.ICON_LOCATION, Encoding)
            End If
        End Sub
    End Structure

    Private Enum ExtraDataBlockSignature As UInt32
        ConsoleDataBlock = &HA0000002UI
        ConsoleFEDataBlock = &HA0000004UI
        DarwinDataBlock = &HA0000006UI
        EnvironmentVariableDataBlock = &HA0000001UI
        IconEnvironmentDataBlock = &HA0000007UI
        KnownFolderDataBlock = &HA000000BUI
        PropertyStoreDataBlock = &HA0000009UI
        ShimDataBlock = &HA0000008UI
        SpecialFolderDataBlock = &HA0000005UI
        TrackerDataBlock = &HA0000058UI
        VistaAndAboveIDListDataBlock = &HA000000CUI
    End Enum

    Private Structure ExtraDataBlock
        Public BlockSize As UInt32
        Public BlockSignature As ExtraDataBlockSignature
        Public Data As Byte()

        Public Sub Read(Stream As System.IO.Stream)
            Stream.Read(BlockSize)
            If BlockSize > 4 Then
                Stream.ReadEnum(BlockSignature)
                ReDim Data(BlockSize - 8 - 1)
                Stream.Read(Data)
            End If
        End Sub

        Public Sub Prepare()
            If Data Is Nothing Then
                BlockSize = 0
            Else
                BlockSize = Data.Length + 8
            End If
        End Sub

        Public Sub Write(Stream As System.IO.MemoryStream)
            If BlockSize > 4 Then
                Stream.Write(BlockSize)
                Stream.Write(BlockSignature)
                Stream.Write(Data)
            Else
                Stream.Write(BlockSize)
            End If
        End Sub

        Public Sub CreateEnvironmentVariableDataBlock(Target As String)
            Me.BlockSize = &H314UI
            Me.BlockSignature = ExtraDataBlockSignature.EnvironmentVariableDataBlock
            ReDim Me.Data(260 + 520)
            Dim TargetAnsi As Byte() = System.Text.Encoding.Default.GetBytes(Target)
            Dim TargetUnicode As Byte() = System.Text.Encoding.Unicode.GetBytes(Target)
            Array.Copy(TargetAnsi, 0, Me.Data, 0, TargetAnsi.Length)
            Array.Copy(TargetUnicode, 0, Me.Data, 260, TargetUnicode.Length)
        End Sub

        Public Sub CreateVistaAndAboveIDListDataBlock(Target As String)
            Me.BlockSignature = ExtraDataBlockSignature.VistaAndAboveIDListDataBlock
            Dim TargetIDList As New LinkTargetIDList
            TargetIDList.Path = Target
            Me.Data = TargetIDList.ItemIDList
            Me.BlockSize = Me.Data.Length + 8
        End Sub

        Public Sub CreateTerminalBlock()
            Me.BlockSize = 0
            Me.Data = Nothing
        End Sub

    End Structure

    Private Structure EXTRA_DATA
        Public EXTRA_DATA_BLOCK_LIST As List(Of ExtraDataBlock)

        Public Sub Read(Stream As System.IO.Stream)
            If Stream.Position < Stream.Length Then
                EXTRA_DATA_BLOCK_LIST = New List(Of ExtraDataBlock)
                While True
                    Dim Block As New ExtraDataBlock
                    Block.Read(Stream)
                    EXTRA_DATA_BLOCK_LIST.Add(Block)
                    If Block.BlockSize <= 4 Then
                        Exit While
                    End If
                End While
            End If
        End Sub

        Public Sub Write(Stream As System.IO.Stream)
            If EXTRA_DATA_BLOCK_LIST IsNot Nothing AndAlso EXTRA_DATA_BLOCK_LIST.Count > 0 Then
                For Each Block As ExtraDataBlock In EXTRA_DATA_BLOCK_LIST
                    Block.Prepare()
                    Block.Write(Stream)
                Next
            End If
        End Sub
    End Structure

#Region "获取Volume信息API"
    'UINT GetDriveTypeW(
    '  LPCWSTR lpRootPathName
    ');
    <DllImport("Kernel32.dll", EntryPoint:="GetDriveTypeW", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function GetDriveType(ByVal lpRootPathName As String) As DRIVE_TYPE
    End Function

    'BOOL GetVolumeInformationW(
    '  LPCWSTR lpRootPathName,
    '  LPWSTR  lpVolumeNameBuffer,
    '  DWORD   nVolumeNameSize,
    '  LPDWORD lpVolumeSerialNumber,
    '  LPDWORD lpMaximumComponentLength,
    '  LPDWORD lpFileSystemFlags,
    '  LPWSTR  lpFileSystemNameBuffer,
    '  DWORD   nFileSystemNameSize
    ');
    <DllImport("Kernel32.dll", CharSet:=CharSet.Unicode, EntryPoint:="GetVolumeInformationW", SetLastError:=True)> _
    Private Shared Function GetVolumeInformation(ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As System.Text.StringBuilder, ByVal nVolumeNameSize As UInt32, ByRef lpVolumeSerialNumber As UInt32, ByRef lpMaximumComponentLength As UInt32, ByRef lpFileSystemFlags As UInt32, ByVal lpFileSystemNameBuffer As System.Text.StringBuilder, ByVal nFileSystemNameSize As UInt32) As Boolean
    End Function

    Private Shared Function GetVolumeInformation(ByVal VolumeLetter As Char, ByRef VolumeSerialNumber As UInt32, ByRef VolumeName As String, ByRef VolumeDriveType As DRIVE_TYPE) As Boolean
        Dim VolumeRoot As String = String.Format("{0}:\", VolumeLetter)
        VolumeDriveType = GetDriveType(VolumeRoot)
        Dim VolumeNameBuilder As New System.Text.StringBuilder(255)
        If GetVolumeInformation(VolumeRoot, VolumeNameBuilder, 255, VolumeSerialNumber, Nothing, Nothing, Nothing, Nothing) = True Then
            VolumeName = VolumeNameBuilder.ToString
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region "获取File信息API"
    'typedef struct _WIN32_FILE_ATTRIBUTE_DATA {
    '  DWORD    dwFileAttributes;
    '  FILETIME ftCreationTime;
    '  FILETIME ftLastAccessTime;
    '  FILETIME ftLastWriteTime;
    '  DWORD    nFileSizeHigh;
    '  DWORD    nFileSizeLow;
    '} WIN32_FILE_ATTRIBUTE_DATA, *LPWIN32_FILE_ATTRIBUTE_DATA;
    <StructLayout(LayoutKind.Sequential)>
    Private Structure WIN32_FILE_ATTRIBUTE_DATA
        Public dwFileAttributes As UInt32
        Public ftCreationTime As ComTypes.FILETIME
        Public ftLastAccessTime As ComTypes.FILETIME
        Public ftLastWriteTime As ComTypes.FILETIME
        Public nFileSizeHigh As UInt32
        Public nFileSizeLow As UInt32
    End Structure

    Private Enum GET_FILEEX_INFO_LEVELS As UInt32
        GetFileExInfoStandard
        GetFileExMaxInfoLevel
    End Enum

    'BOOL GetFileAttributesExW(
    '  LPCWSTR                lpFileName,
    '  GET_FILEEX_INFO_LEVELS fInfoLevelId,
    '  LPVOID                 lpFileInformation
    ');
    <DllImport("Kernel32.dll", EntryPoint:="GetFileAttributesExW", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function GetFileAttributesEx(ByVal lpFileName As String, ByVal fInfoLevelId As GET_FILEEX_INFO_LEVELS, ByRef lpFileInformation As WIN32_FILE_ATTRIBUTE_DATA) As Boolean
    End Function

    ''' <summary>
    ''' 根据Path获取文件信息。
    ''' </summary>
    ''' <param name="Path"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetFileInformation(Path As String) As WIN32_FILE_ATTRIBUTE_DATA
        Dim FileAttributesEx As New WIN32_FILE_ATTRIBUTE_DATA
        If GetFileAttributesEx(Path, GET_FILEEX_INFO_LEVELS.GetFileExInfoStandard, FileAttributesEx) = False Then
            'Message = String.Format("GetFileAttributesEx({0})失败：{1}", Path, New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error).Message)
        End If
        Return FileAttributesEx
    End Function
#End Region

    Public Class ShellLinkInformation
        Public TargetPath As String
        Public TargetArgs As String
        Public WorkingDir As String
        Public IconPath As String
        Public IconIndx As Integer

        Private LinkHeader As New ShellLinkHeader
        Private TargetIDList As New LinkTargetIDList
        Private LinkInfo As New LinkInfo
        Private StringData As New STRING_DATA
        Private ExtraData As New EXTRA_DATA

        Public Sub Load(LnkFile As String)
            Dim LnkStream As New System.IO.MemoryStream(System.IO.File.ReadAllBytes(LnkFile))

            LinkHeader.Read(LnkStream)
            Debug.Print(GetObjectString(LinkHeader))

            If LinkHeader.LinkFlags.HasFlag(LinkFlag.HasLinkTargetIDList) Then
                TargetIDList.Read(LnkStream)
                Debug.Print(GetObjectString(TargetIDList))
                Debug.Print("TargetIDList.Path={0}", TargetIDList.Path)
            End If

            If LinkHeader.LinkFlags.HasFlag(LinkFlag.HasLinkInfo) Then
                LinkInfo.Read(LnkStream)
                Debug.Print(GetObjectString(LinkInfo.VolumeID))
                Debug.Print(GetObjectString(LinkInfo.CommonNetworkRelativeLink))
                Debug.Print(GetObjectString(LinkInfo))
            End If

            StringData.Read(LnkStream, LinkHeader.LinkFlags)
            Debug.Print(GetObjectString(StringData))

            ExtraData.Read(LnkStream)
            Debug.Print(GetObjectString(ExtraData))
            If ExtraData.EXTRA_DATA_BLOCK_LIST IsNot Nothing Then
                For Each Block As ExtraDataBlock In ExtraData.EXTRA_DATA_BLOCK_LIST
                    Debug.Print(GetObjectString(Block))
                Next
            End If


            If LinkHeader.LinkFlags.HasFlag(LinkFlag.HasLinkTargetIDList) Then
                Me.TargetPath = TargetIDList.Path
            Else
                Select Case LinkInfo.LinkInfoFlags
                    Case LinkInfoFlag.VolumeIDAndLocalBasePath
                        Me.TargetPath = LinkInfo.LocalBasePath
                    Case LinkInfoFlag.CommonNetworkRelativeLinkAndPathSuffix
                        Select Case LinkInfo.CommonNetworkRelativeLink.CommonNetworkRelativeLinkFlags
                            Case CommonNetworkRelativeLinkFlag.ValidNetType
                                Me.TargetPath = System.IO.Path.Combine(LinkInfo.CommonNetworkRelativeLink.NetName, LinkInfo.CommonPathSuffix)
                            Case CommonNetworkRelativeLinkFlag.ValidDevice
                                Me.TargetPath = System.IO.Path.Combine(LinkInfo.CommonNetworkRelativeLink.DeviceName, LinkInfo.CommonPathSuffix)
                        End Select
                End Select
            End If
            Me.TargetArgs = StringData.COMMAND_LINE_ARGUMENTS
            Me.WorkingDir = StringData.WORKING_DIR
            Me.IconIndx = LinkHeader.IconIndex
            Me.IconPath = StringData.ICON_LOCATION

            LnkStream.Dispose()
        End Sub

        Public Sub Save(LnkFile As String)
            'TargetPath = "\\XMNBS\Public\Doc\数据共享\数据文件\IvdSystem.02.00.07520.23873.exe"

            TargetPath = System.IO.Path.GetFullPath(TargetPath).TrimEnd("\")
            If System.IO.File.Exists(TargetPath) = False AndAlso System.IO.Directory.Exists(TargetPath) = False Then
                Throw New System.IO.FileNotFoundException
            End If

            Dim TargetInfo As WIN32_FILE_ATTRIBUTE_DATA = GetFileInformation(TargetPath)
            LinkHeader.FileAttributes = TargetInfo.dwFileAttributes
            LinkHeader.CreationTime = TargetInfo.ftCreationTime.ToSysDateTime
            LinkHeader.AccessTime = TargetInfo.ftLastAccessTime.ToSysDateTime
            LinkHeader.WriteTime = TargetInfo.ftLastWriteTime.ToSysDateTime
            LinkHeader.FileSize = TargetInfo.nFileSizeLow

            LinkHeader.IconIndex = Me.IconIndx

            If LinkHeader.FileAttributes.HasFlag(FileAttribute.FILE_ATTRIBUTE_DIRECTORY) Then
                WorkingDir = Nothing
            Else
                If String.IsNullOrEmpty(WorkingDir) Then
                    WorkingDir = System.IO.Path.GetDirectoryName(TargetPath)
                End If
            End If
            StringData.WORKING_DIR = WorkingDir
            StringData.ICON_LOCATION = IconPath
            StringData.COMMAND_LINE_ARGUMENTS = TargetArgs
            If String.IsNullOrEmpty(StringData.WORKING_DIR) = False Then
                LinkHeader.LinkFlags = LinkHeader.LinkFlags Or LinkFlag.HasWorkingDir
            End If
            If String.IsNullOrEmpty(StringData.ICON_LOCATION) = False Then
                LinkHeader.LinkFlags = LinkHeader.LinkFlags Or LinkFlag.HasIconLocation
            End If
            If String.IsNullOrEmpty(StringData.COMMAND_LINE_ARGUMENTS) = False Then
                LinkHeader.LinkFlags = LinkHeader.LinkFlags Or LinkFlag.HasArguments
            End If

            '精简型
            LinkHeader.LinkFlags = LinkHeader.LinkFlags Or LinkFlag.HasLinkTargetIDList Or LinkFlag.IsUnicode
            TargetIDList.Path = TargetPath

            '复杂型
            LinkHeader.LinkFlags = LinkHeader.LinkFlags Or LinkFlag.HasLinkInfo
            If TargetPath.StartsWith("\\") Then
                LinkInfo.LinkInfoFlags = LinkInfoFlag.CommonNetworkRelativeLinkAndPathSuffix
                LinkInfo.CommonNetworkRelativeLink = New CommonNetworkRelativeLink
                LinkInfo.CommonNetworkRelativeLink.CommonNetworkRelativeLinkFlags = CommonNetworkRelativeLinkFlag.ValidNetType
                Dim NetNameSize As Integer = 2
                While True
                    Dim Idx As Integer = TargetPath.IndexOf("\", NetNameSize)
                    If Idx = -1 Then
                        Exit While
                    End If
                    Dim NetNameTemp As String = TargetPath.Substring(0, Idx)
                    If System.IO.Directory.Exists(NetNameTemp) Then
                        LinkInfo.CommonNetworkRelativeLink.NetName = NetNameTemp
                        LinkInfo.CommonPathSuffix = TargetPath.Substring(Idx + 1)
                        Exit While
                    Else
                        NetNameSize = Idx + 1
                    End If
                End While
                Debug.Print("LinkInfo.CommonNetworkRelativeLink.NetName={0}", LinkInfo.CommonNetworkRelativeLink.NetName)
                Debug.Print("LinkInfo.CommonPathSuffix={0}", LinkInfo.CommonPathSuffix)
            Else
                LinkInfo.LinkInfoFlags = LinkInfoFlag.VolumeIDAndLocalBasePath
                Dim VolumeLetter As Char = TargetPath.Substring(0, 1)
                LinkInfo.VolumeID = New VolumeID
                GetVolumeInformation(VolumeLetter, LinkInfo.VolumeID.DriveSerialNumber, LinkInfo.VolumeID.VolumeLabel, LinkInfo.VolumeID.DriveType)
                LinkInfo.LocalBasePath = TargetPath
                Debug.Print("LinkInfo.LocalBasePath={0}", LinkInfo.LocalBasePath)
            End If




            Dim LnkStream As New System.IO.MemoryStream
            LinkHeader.Write(LnkStream)
            'Debug.Print(GetObjectString(LinkHeader))
            If LinkHeader.LinkFlags.HasFlag(LinkFlag.HasLinkTargetIDList) Then
                TargetIDList.Prepare()
                TargetIDList.Write(LnkStream)
                'Debug.Print(GetObjectString(TargetIDList))
                'Debug.Print(TargetIDList.Path)
            End If
            If LinkHeader.LinkFlags.HasFlag(LinkFlag.HasLinkInfo) Then
                LinkInfo.Prepare()
                LinkInfo.Write(LnkStream)
                'Debug.Print(GetObjectString(LinkInfo.VolumeID))
                'Debug.Print(GetObjectString(LinkInfo.CommonNetworkRelativeLink))
                'Debug.Print(GetObjectString(LinkInfo))
            End If
            StringData.Write(LnkStream, LinkHeader.LinkFlags)
            'Debug.Print(GetObjectString(StringData))
            ExtraData.Write(LnkStream)
            'Debug.Print(GetObjectString(ExtraData))

            System.IO.File.WriteAllBytes(LnkFile, LnkStream.ToArray)
        End Sub
    End Class

    Public Shared Sub Test()
        Dim OldFile As String = "C:\Users\liubin\Desktop\Lnk\160911移除.lnkx"
        Dim NewFile As String = "C:\Users\liubin\Desktop\Lnk\Test.lnk"
        Dim LnkInfo As New ShellLinkInformation
        LnkInfo.Load(OldFile)

        'LnkInfo.TargetPath = "C:\Users\liubin\Desktop\Lnk\⁂.txt"
        'LnkInfo.TargetArgs = "-a 123"
        'LnkInfo.IconPath = "C:\Windows\regedit.exe"

        'LnkInfo.Save(NewFile)

        'Debug.Print("比较开始")
        'Dim OldStream As New System.IO.MemoryStream(System.IO.File.ReadAllBytes(OldFile))
        'Dim NewStream As New System.IO.MemoryStream(System.IO.File.ReadAllBytes(NewFile))
        'For i As Integer = 0 To NewStream.Length - 1
        '    If OldStream.ReadByte <> NewStream.ReadByte Then
        '        Debug.Print(i.ToString("X"))
        '        Exit For
        '    End If
        'Next
        'Debug.Print("比较结束")
    End Sub

End Class
