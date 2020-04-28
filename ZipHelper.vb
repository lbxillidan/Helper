Public Class ZipHelper
    <Flags>
    Private Enum GeneralPurposeBitFlagType As UInt16
        Encrypted = CUShort(1) << 15
        HasDataDescriptor = CUShort(1) << 12
        UTF8 = CUShort(1) << 11
    End Enum

    Private Enum CompressionMethodType As UInt16
        Stored = 0
        Deflated = 8
    End Enum

    Private Enum ZipSignature As UInt32
        LocalFileHeader = &H4034B50UI
        CentralFileHeader = &H2014B50UI
        EndOfCentralDirectory = &H6054B50UI

        ArchiveExtraData = &H8064B50UI
        DigitalSignatureHeader = &H5054B50UI
        Zip64EndOfCentralDirectory = &H6064B50UI
        Zip64EndOfCentralDirectoryLocator = &H7064B50UI
    End Enum

    ''' <summary>
    ''' 未压缩文件信息
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ZipFileInfo
        Public FileName As String
        Public FileDate As Date
        Public FileData As Byte()
        Public FileComment As String
    End Class

    ''' <summary>
    ''' Zip压缩包
    ''' </summary>
    ''' <remarks></remarks>
    Private Class ZipArchive
        Public FileBits() As Byte

        Public LocalFileDict As New Dictionary(Of String, ZipLocalFile)
        Public CentralDirectoryFileDict As New Dictionary(Of String, ZipCentralDirectoryFile)
        Public EndOfCentralDirectoryRecord As ZipEndOfCentralDirectoryRecord

        Public Sub Open(data As Byte())
            Me.FileBits = data

            Me.EndOfCentralDirectoryRecord = ZipEndOfCentralDirectoryRecord.Find(Me)
            If Me.EndOfCentralDirectoryRecord Is Nothing Then
                Throw New FormatException("不符合Zip文件结构")
            End If

            Dim CentralDirectoryFileOffset As Integer = Me.EndOfCentralDirectoryRecord.OffsetOfStartOfCentralDirectoryWithRespectToTheStartingDiskNumber
            While ZipCentralDirectoryFile.Check(Me, CentralDirectoryFileOffset)
                Dim CentralDirectoryFileTemp As New ZipCentralDirectoryFile(Me, CentralDirectoryFileOffset)
                Me.CentralDirectoryFileDict.Add(CentralDirectoryFileTemp.FileInfo.FileName, CentralDirectoryFileTemp)
                CentralDirectoryFileOffset = CentralDirectoryFileOffset + CentralDirectoryFileTemp.Length
            End While
        End Sub

        Public Sub Open(file As String)
            Open(System.IO.File.ReadAllBytes(file))
        End Sub

        Public Function Contains(name As String) As Boolean
            Return Me.CentralDirectoryFileDict.ContainsKey(name)
        End Function

        Public Function UnZip(name As String) As ZipFileInfo
            If Me.CentralDirectoryFileDict.ContainsKey(name) = False Then
                Return Nothing
            End If
            Dim LocalFile As New ZipLocalFile(Me, Me.CentralDirectoryFileDict.Item(name).RelativeOffsetOfLocalHeader)
            Return LocalFile.FileInfo
        End Function

        Public Function UnZip() As ZipFileInfo()
            Dim FileInfos(Me.CentralDirectoryFileDict.Count - 1) As ZipFileInfo
            Dim i As Integer = 0
            For Each CentralDirectoryFile As ZipCentralDirectoryFile In CentralDirectoryFileDict.Values
                Dim LocalFileTemp As New ZipLocalFile(Me, CentralDirectoryFile.RelativeOffsetOfLocalHeader)
                FileInfos(i) = LocalFileTemp.FileInfo
                i = i + 1
            Next
            Return FileInfos
        End Function

        Public Function Zip(files As ZipFileInfo()) As Boolean
            Dim FileBitsList As New List(Of Byte)
            Dim offset As Integer = 0
            Dim count As Integer = 0
            Dim size As Integer = 0
            For Each FileInfo As ZipFileInfo In files
                Dim LocalFile As New ZipLocalFile(FileInfo)
                LocalFileDict.Add(FileInfo.FileName, LocalFile)
                FileBitsList.AddRange(LocalFile.Bits)
                Dim CentralDirectoryFile As New ZipCentralDirectoryFile(LocalFile, offset)
                CentralDirectoryFileDict.Add(FileInfo.FileName, CentralDirectoryFile)
                offset = offset + LocalFile.Length
                count = count + 1
                size = size + CentralDirectoryFile.Length
            Next
            For Each CentralDirectoryFile As ZipCentralDirectoryFile In CentralDirectoryFileDict.Values
                FileBitsList.AddRange(CentralDirectoryFile.Bits)
            Next
            Me.EndOfCentralDirectoryRecord = New ZipEndOfCentralDirectoryRecord(count, size, offset)
            FileBitsList.AddRange(Me.EndOfCentralDirectoryRecord.Bits)
            Me.FileBits = FileBitsList.ToArray
            Return True
        End Function

        Public Function Zip(files As String()) As Boolean
            Dim FileBound As Integer = files.Length - 1
            If FileBound = -1 Then
                Return False
            End If
            Dim FileInfos(FileBound) As ZipFileInfo
            Dim FileNames(FileBound) As String
            Dim File As String
            Dim i As Integer

            If FileBound = 0 Then
                FileNames(0) = System.IO.Path.GetFileName(files(0))
            Else
                '去除路径开头的盘符
                For i = 0 To FileBound
                    FileNames(i) = files(i).Substring(3)
                Next
                '去除共同路径
                File = FileNames(0)
                Dim Root As String = String.Empty
                While True
                    Dim RootIndex As Integer = File.IndexOf("\", Root.Length + 1)
                    If RootIndex = -1 Then
                        Exit While
                    End If
                    Dim RootTemp As String = File.Substring(0, RootIndex + 1)
                    For i = 1 To FileBound
                        If FileNames(i).StartsWith(RootTemp) = False Then
                            Exit While
                        End If
                    Next
                    Root = RootTemp
                End While
                'Debug.Print("Root={0}", Root)
                If String.IsNullOrEmpty(Root) = False Then
                    Dim RootLength As Integer = Root.Length
                    For i = 0 To FileBound
                        FileNames(i) = FileNames(i).Substring(RootLength)
                    Next
                End If
            End If

            '构建FileInfos
            For i = 0 To FileBound
                Dim FileInfoTemp As New ZipFileInfo
                File = files(i)
                FileInfoTemp.FileName = FileNames(i)
                'Debug.Print(FileNames(i))
                FileInfoTemp.FileDate = System.IO.File.GetLastWriteTime(File)
                FileInfoTemp.FileData = System.IO.File.ReadAllBytes(File)
                FileInfos(i) = FileInfoTemp
            Next

            Return Zip(FileInfos)
        End Function
    End Class

    ''' <summary> 
    ''' 压缩包内部文件内容
    ''' </summary>
    ''' <remarks></remarks>
    Private Class ZipLocalFile
        Private Enum FieldOffset As Integer
            Signature = 0
            VersionNeededToExtract = 4
            GeneralPurposeBitFlag = 6
            CompressionMethod = 8
            LastModFileTime = 10
            LastModFileDate = 12
            Crc32 = 14
            CompressedSize = 18
            UncompressedSize = 22
            FileNameLength = 26
            ExtraFieldLength = 28
            FileName = 30
            'ExtraField
            'FileData
        End Enum

        Public Signature As UInt32
        Public VersionNeededToExtract As UInt16
        Public GeneralPurposeBitFlag As GeneralPurposeBitFlagType
        Public CompressionMethod As CompressionMethodType
        Public LastModFileTime As UInt16
        Public LastModFileDate As UInt16
        Public Crc32 As UInt32
        Public CompressedSize As UInt32
        Public UncompressedSize As UInt32
        Public FileNameLength As UInt16
        Public ExtraFieldLength As UInt16
        Public FileName As Byte()
        Public ExtraField As Byte()
        Public FileData As Byte()

        Public Length As Integer
        Public Bits As Byte()
        Public FileInfo As ZipFileInfo

        Sub New(file As ZipFileInfo)
            Me.FileInfo = file
            Me.Signature = ZipSignature.LocalFileHeader
            Me.VersionNeededToExtract = 0
            Me.GeneralPurposeBitFlag = GeneralPurposeBitFlagType.UTF8
            Me.CompressionMethod = CompressionMethodType.Deflated
            Me.FileData = DeflateHelper.Encode(file.FileData)
            Me.UncompressedSize = file.FileData.Length
            Me.CompressedSize = Me.FileData.Length
            Me.Crc32 = CRC32Helper.ComputeAsUInt32(file.FileData)
            DosDateTimeHelper.SetSystemDateTime(file.FileDate, Me.LastModFileDate, Me.LastModFileTime)
            Me.FileName = System.Text.Encoding.UTF8.GetBytes(file.FileName)
            Me.FileNameLength = Me.FileName.Length
            Me.ExtraField = New Byte() {}
            Me.ExtraFieldLength = Me.ExtraField.Length

            Me.Length = FieldOffset.FileName + Me.FileNameLength + Me.ExtraFieldLength + Me.CompressedSize
            ReDim Me.Bits(Me.Length - 1)
            bithelper.WriteUInt32(Me.Bits, FieldOffset.Signature, Me.Signature)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.VersionNeededToExtract, Me.VersionNeededToExtract)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.GeneralPurposeBitFlag, Me.GeneralPurposeBitFlag)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.CompressionMethod, Me.CompressionMethod)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.LastModFileTime, Me.LastModFileTime)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.LastModFileDate, Me.LastModFileDate)
            bithelper.WriteUInt32(Me.Bits, FieldOffset.Crc32, Me.Crc32)
            bithelper.WriteUInt32(Me.Bits, FieldOffset.CompressedSize, Me.CompressedSize)
            bithelper.WriteUInt32(Me.Bits, FieldOffset.UncompressedSize, Me.UncompressedSize)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.FileNameLength, Me.FileNameLength)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.ExtraFieldLength, Me.ExtraFieldLength)
            bithelper.WriteBytes(Me.Bits, FieldOffset.FileName, Me.FileName)
            bithelper.WriteBytes(Me.Bits, FieldOffset.FileName + Me.FileNameLength, Me.ExtraField)
            bithelper.WriteBytes(Me.Bits, FieldOffset.FileName + Me.FileNameLength + Me.ExtraFieldLength, Me.FileData)
        End Sub

        Sub New(zip As ZipArchive, offset As Integer)
            Me.Signature = bithelper.ReadUInt32(zip.FileBits, offset + FieldOffset.Signature)
            Me.VersionNeededToExtract = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.VersionNeededToExtract)
            Me.GeneralPurposeBitFlag = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.GeneralPurposeBitFlag)
            Me.CompressionMethod = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.CompressionMethod)
            Me.LastModFileTime = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.LastModFileTime)
            Me.LastModFileDate = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.LastModFileDate)
            Me.Crc32 = bithelper.ReadUInt32(zip.FileBits, offset + FieldOffset.Crc32)
            Me.CompressedSize = bithelper.ReadUInt32(zip.FileBits, offset + FieldOffset.CompressedSize)
            Me.UncompressedSize = bithelper.ReadUInt32(zip.FileBits, offset + FieldOffset.UncompressedSize)
            Me.FileNameLength = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.FileNameLength)
            Me.ExtraFieldLength = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.ExtraFieldLength)
            Me.FileName = bithelper.ReadBytes(zip.FileBits, offset + FieldOffset.FileName, Me.FileNameLength)
            Me.ExtraField = bithelper.ReadBytes(zip.FileBits, offset + FieldOffset.FileName + Me.FileNameLength, Me.ExtraFieldLength)
            Me.FileData = bithelper.ReadBytes(zip.FileBits, offset + FieldOffset.FileName + Me.FileNameLength + Me.ExtraFieldLength, Me.CompressedSize)

            Me.Length = FieldOffset.FileName + Me.FileNameLength + Me.ExtraFieldLength + Me.CompressedSize
            Me.FileInfo = New ZipFileInfo
            If Me.GeneralPurposeBitFlag And GeneralPurposeBitFlagType.UTF8 = GeneralPurposeBitFlagType.UTF8 Then
                Me.FileInfo.FileName = System.Text.Encoding.UTF8.GetString(Me.FileName)
            Else
                If EncodingHelper.IsGBK(Me.FileName) Then
                    Me.FileInfo.FileName = System.Text.Encoding.Default.GetString(Me.FileName)
                Else
                    Me.FileInfo.FileName = System.Text.Encoding.UTF8.GetString(Me.FileName)
                End If
            End If
            Me.FileInfo.FileDate = DosDateTimeHelper.GetSystemDateTime(Me.LastModFileDate, Me.LastModFileTime)
            Select Case Me.CompressionMethod
                Case CompressionMethodType.Stored
                    Me.FileInfo.FileData = Me.FileData
                Case CompressionMethodType.Deflated
                    Me.FileInfo.FileData = DeflateHelper.Decode(Me.FileData)
                Case Else
                    Throw New NotSupportedException("不支持的压缩方法")
            End Select
        End Sub

        Public Shared Function Check(zip As ZipArchive, offset As Integer) As Boolean
            If zip.FileBits.Length < offset + 4 Then
                Return False
            Else
                If bithelper.ReadUInt32(zip.FileBits, offset) = ZipSignature.LocalFileHeader Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Function
    End Class

    ''' <summary>
    ''' 压缩包内部文件信息
    ''' </summary>
    ''' <remarks></remarks>
    Private Class ZipCentralDirectoryFile
        Private Enum FieldOffset As Integer
            Signature = 0
            VersionMadeBy = 4
            VersionNeededToExtract = 6
            GeneralPurposeBitFlag = 8
            CompressionMethod = 10
            LastModFileTime = 12
            LastModFileDate = 14
            Crc32 = 16
            CompressedSize = 20
            UncompressedSize = 24
            FileNameLength = 28
            ExtraFieldLength = 30
            FileCommentLength = 32
            DiskNumberStart = 34
            InternalFileAttributes = 36
            ExternalFileAttributes = 38
            RelativeOffsetOfLocalHeader = 42
            FileName = 46
            'ExtraField
            'FileComment
        End Enum

        Public Signature As UInt32
        Public VersionMadeBy As UInt16
        Public VersionNeededToExtract As UInt16
        Public GeneralPurposeBitFlag As GeneralPurposeBitFlagType
        Public CompressionMethod As UInt16
        Public LastModFileTime As UInt16
        Public LastModFileDate As UInt16
        Public Crc32 As UInt32
        Public CompressedSize As UInt32
        Public UncompressedSize As UInt32
        Public FileNameLength As UInt16
        Public ExtraFieldLength As UInt16
        Public FileCommentLength As UInt16
        Public DiskNumberStart As UInt16
        Public InternalFileAttributes As UInt16
        Public ExternalFileAttributes As UInt32
        Public RelativeOffsetOfLocalHeader As UInt32
        Public FileName As Byte()
        Public ExtraField As Byte()
        Public FileComment As Byte()

        Public Length As Integer
        Public Bits As Byte()
        Public FileInfo As ZipFileInfo

        Sub New(file As ZipLocalFile, offset As Integer)
            Me.FileInfo = file.FileInfo
            Me.Signature = ZipSignature.CentralFileHeader
            Me.VersionMadeBy = 0
            Me.VersionNeededToExtract = file.VersionNeededToExtract
            Me.GeneralPurposeBitFlag = file.GeneralPurposeBitFlag
            Me.CompressionMethod = file.CompressionMethod
            Me.LastModFileTime = file.LastModFileTime
            Me.LastModFileDate = file.LastModFileDate
            Me.Crc32 = file.Crc32
            Me.CompressedSize = file.CompressedSize
            Me.UncompressedSize = file.UncompressedSize
            Me.FileName = file.FileName
            Me.FileNameLength = file.FileNameLength
            Me.ExtraField = file.ExtraField
            Me.ExtraFieldLength = file.ExtraFieldLength
            If String.IsNullOrEmpty(file.FileInfo.FileComment) = False Then
                Me.FileComment = System.Text.Encoding.UTF8.GetBytes(file.FileInfo.FileComment)
            Else
                Me.FileComment = New Byte() {}
            End If
            Me.FileCommentLength = Me.FileComment.Length
            Me.DiskNumberStart = 0
            Me.InternalFileAttributes = 0
            Me.ExternalFileAttributes = 0
            Me.RelativeOffsetOfLocalHeader = offset

            Me.Length = FieldOffset.FileName + Me.FileNameLength + Me.ExtraFieldLength + Me.FileCommentLength

            ReDim Me.Bits(Me.Length - 1)
            bithelper.WriteUInt32(Me.Bits, FieldOffset.Signature, Me.Signature)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.VersionMadeBy, Me.VersionMadeBy)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.VersionNeededToExtract, Me.VersionNeededToExtract)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.GeneralPurposeBitFlag, Me.GeneralPurposeBitFlag)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.CompressionMethod, Me.CompressionMethod)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.LastModFileTime, Me.LastModFileTime)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.LastModFileDate, Me.LastModFileDate)
            bithelper.WriteUInt32(Me.Bits, FieldOffset.Crc32, Me.Crc32)
            bithelper.WriteUInt32(Me.Bits, FieldOffset.CompressedSize, Me.CompressedSize)
            bithelper.WriteUInt32(Me.Bits, FieldOffset.UncompressedSize, Me.UncompressedSize)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.FileNameLength, Me.FileNameLength)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.ExtraFieldLength, Me.ExtraFieldLength)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.FileCommentLength, Me.FileCommentLength)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.DiskNumberStart, Me.DiskNumberStart)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.InternalFileAttributes, Me.InternalFileAttributes)
            bithelper.WriteUInt16(Me.Bits, FieldOffset.ExternalFileAttributes, Me.ExternalFileAttributes)
            bithelper.WriteUInt32(Me.Bits, FieldOffset.RelativeOffsetOfLocalHeader, Me.RelativeOffsetOfLocalHeader)
            bithelper.WriteBytes(Me.Bits, FieldOffset.FileName, Me.FileName)
            bithelper.WriteBytes(Me.Bits, FieldOffset.FileName + Me.FileNameLength, Me.ExtraField)
            bithelper.WriteBytes(Me.Bits, FieldOffset.FileName + Me.FileNameLength + Me.ExtraFieldLength, Me.FileComment)
        End Sub

        Sub New(zip As ZipArchive, offset As Integer)
            Me.Signature = bithelper.ReadUInt32(zip.FileBits, offset + FieldOffset.Signature)
            Me.VersionMadeBy = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.VersionMadeBy)
            Me.VersionNeededToExtract = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.VersionNeededToExtract)
            Me.GeneralPurposeBitFlag = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.GeneralPurposeBitFlag)
            Me.CompressionMethod = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.CompressionMethod)
            Me.LastModFileTime = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.LastModFileTime)
            Me.LastModFileDate = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.LastModFileDate)
            Me.Crc32 = bithelper.ReadUInt32(zip.FileBits, offset + FieldOffset.Crc32)
            Me.CompressedSize = bithelper.ReadUInt32(zip.FileBits, offset + FieldOffset.CompressedSize)
            Me.UncompressedSize = bithelper.ReadUInt32(zip.FileBits, offset + FieldOffset.UncompressedSize)
            Me.FileNameLength = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.FileNameLength)
            Me.ExtraFieldLength = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.ExtraFieldLength)
            Me.FileCommentLength = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.FileCommentLength)
            Me.DiskNumberStart = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.DiskNumberStart)
            Me.InternalFileAttributes = bithelper.ReadUInt16(zip.FileBits, offset + FieldOffset.InternalFileAttributes)
            Me.ExternalFileAttributes = bithelper.ReadUInt32(zip.FileBits, offset + FieldOffset.ExternalFileAttributes)
            Me.RelativeOffsetOfLocalHeader = bithelper.ReadUInt32(zip.FileBits, offset + FieldOffset.RelativeOffsetOfLocalHeader)
            Me.FileName = bithelper.ReadBytes(zip.FileBits, offset + FieldOffset.FileName, Me.FileNameLength)
            Me.ExtraField = bithelper.ReadBytes(zip.FileBits, offset + FieldOffset.FileName + Me.FileNameLength, Me.ExtraFieldLength)
            Me.FileComment = bithelper.ReadBytes(zip.FileBits, offset + FieldOffset.FileName + Me.FileNameLength + Me.ExtraFieldLength, Me.FileCommentLength)

            Me.Length = FieldOffset.FileName + Me.FileNameLength + Me.ExtraFieldLength + Me.FileCommentLength
            Me.FileInfo = New ZipFileInfo
            If Me.GeneralPurposeBitFlag And GeneralPurposeBitFlagType.UTF8 = GeneralPurposeBitFlagType.UTF8 Then
                Me.FileInfo.FileName = System.Text.Encoding.UTF8.GetString(Me.FileName)
                Me.FileInfo.FileComment = System.Text.Encoding.UTF8.GetString(Me.FileComment)
            Else
                If EncodingHelper.IsGBK(Me.FileName) Then
                    Me.FileInfo.FileName = System.Text.Encoding.Default.GetString(Me.FileName)
                    Me.FileInfo.FileComment = System.Text.Encoding.Default.GetString(Me.FileComment)
                Else
                    Me.FileInfo.FileName = System.Text.Encoding.UTF8.GetString(Me.FileName)
                    Me.FileInfo.FileComment = System.Text.Encoding.UTF8.GetString(Me.FileComment)
                End If
            End If
            Me.FileInfo.FileDate = DosDateTimeHelper.GetSystemDateTime(Me.LastModFileDate, Me.LastModFileTime)
        End Sub

        Public Shared Function Check(zip As ZipArchive, offset As Integer) As Boolean
            If zip.FileBits.Length < offset + 4 Then
                Return False
            Else
                If bithelper.ReadUInt32(zip.FileBits, offset) = ZipSignature.CentralFileHeader Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Function
    End Class

    Private Class ZipDigitalSignature
        Public Signature As UInt32
        Public SizeOfData As UInt16
        Public SignatureData As Byte()

        Public Length As Integer

        Sub New()
        End Sub

        Sub New(zip As ZipArchive, offset As Integer)
            Me.Signature = bithelper.ReadUInt32(zip.FileBits, offset)
            Me.SizeOfData = bithelper.ReadUInt16(zip.FileBits, offset + 4)
            Me.SignatureData = bithelper.ReadBytes(zip.FileBits, offset + 6, Me.SizeOfData)

            Me.Length = 6 + Me.SizeOfData
        End Sub

        Public Shared Function Check(zip As ZipArchive, offset As Integer) As Boolean
            If zip.FileBits.Length < offset + 4 Then
                Return False
            Else
                If BitHelper.ReadUInt32(zip.FileBits, offset) = ZipSignature.DigitalSignatureHeader Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Function

    End Class

    ''' <summary>
    ''' Zip压缩包文件尾
    ''' </summary>
    ''' <remarks></remarks>
    Private Class ZipEndOfCentralDirectoryRecord
        Private Enum FieldOffset As Integer
            Signature = 0
            NumberOfThisDisk = 4
            NumberOfTheDiskWithTheStartOfTheCentralDirectory = 6
            TotalNumberOfEntriesInTheCentralDirectoryOnThisDisk = 8
            TotalNumberOfEntriesInTheCentralDirectory = 10
            SizeOfTheCentralDirectory = 12
            OffsetOfStartOfCentralDirectoryWithRespectToTheStartingDiskNumber = 16
            ZipFileCommentLength = 20
            ZipFileComment = 22
        End Enum

        Public Signature As UInt32
        Public NumberOfThisDisk As UInt16
        Public NumberOfTheDiskWithTheStartOfTheCentralDirectory As UInt16
        Public TotalNumberOfEntriesInTheCentralDirectoryOnThisDisk As UInt16
        Public TotalNumberOfEntriesInTheCentralDirectory As UInt16
        Public SizeOfTheCentralDirectory As UInt32
        Public OffsetOfStartOfCentralDirectoryWithRespectToTheStartingDiskNumber As UInt32
        Public ZipFileCommentLength As UInt16
        Public ZipFileComment As Byte()

        Public Length As Integer
        Public Bits As Byte()

        Sub New(count As Integer, size As Integer, offset As Integer)
            Me.Signature = ZipSignature.EndOfCentralDirectory
            Me.NumberOfThisDisk = 0
            Me.NumberOfTheDiskWithTheStartOfTheCentralDirectory = 0
            Me.TotalNumberOfEntriesInTheCentralDirectoryOnThisDisk = count
            Me.TotalNumberOfEntriesInTheCentralDirectory = count
            Me.SizeOfTheCentralDirectory = size
            Me.OffsetOfStartOfCentralDirectoryWithRespectToTheStartingDiskNumber = offset
            Me.ZipFileCommentLength = 0
            Me.ZipFileComment = New Byte() {}

            Me.Length = FieldOffset.ZipFileComment + Me.ZipFileCommentLength
            ReDim Me.Bits(Me.Length - 1)
            BitHelper.WriteUInt32(Me.Bits, FieldOffset.Signature, Me.Signature)
            BitHelper.WriteUInt16(Me.Bits, FieldOffset.NumberOfThisDisk, Me.NumberOfThisDisk)
            BitHelper.WriteUInt16(Me.Bits, FieldOffset.NumberOfTheDiskWithTheStartOfTheCentralDirectory, Me.NumberOfTheDiskWithTheStartOfTheCentralDirectory)
            BitHelper.WriteUInt16(Me.Bits, FieldOffset.TotalNumberOfEntriesInTheCentralDirectoryOnThisDisk, Me.TotalNumberOfEntriesInTheCentralDirectoryOnThisDisk)
            BitHelper.WriteUInt16(Me.Bits, FieldOffset.TotalNumberOfEntriesInTheCentralDirectory, Me.TotalNumberOfEntriesInTheCentralDirectory)
            BitHelper.WriteUInt32(Me.Bits, FieldOffset.SizeOfTheCentralDirectory, Me.SizeOfTheCentralDirectory)
            BitHelper.WriteUInt32(Me.Bits, FieldOffset.OffsetOfStartOfCentralDirectoryWithRespectToTheStartingDiskNumber, Me.OffsetOfStartOfCentralDirectoryWithRespectToTheStartingDiskNumber)
            BitHelper.WriteUInt16(Me.Bits, FieldOffset.ZipFileCommentLength, Me.ZipFileCommentLength)
            BitHelper.WriteBytes(Me.Bits, FieldOffset.ZipFileComment, Me.ZipFileComment)
        End Sub

        Sub New(zip As ZipArchive, offset As Integer)
            Me.Signature = BitHelper.ReadUInt32(zip.FileBits, offset + FieldOffset.Signature)
            Me.NumberOfThisDisk = BitHelper.ReadUInt16(zip.FileBits, offset + FieldOffset.NumberOfThisDisk)
            Me.NumberOfTheDiskWithTheStartOfTheCentralDirectory = BitHelper.ReadUInt16(zip.FileBits, offset + FieldOffset.NumberOfTheDiskWithTheStartOfTheCentralDirectory)
            Me.TotalNumberOfEntriesInTheCentralDirectoryOnThisDisk = BitHelper.ReadUInt16(zip.FileBits, offset + FieldOffset.TotalNumberOfEntriesInTheCentralDirectoryOnThisDisk)
            Me.TotalNumberOfEntriesInTheCentralDirectory = BitHelper.ReadUInt16(zip.FileBits, offset + FieldOffset.TotalNumberOfEntriesInTheCentralDirectory)
            Me.SizeOfTheCentralDirectory = BitHelper.ReadUInt32(zip.FileBits, offset + FieldOffset.SizeOfTheCentralDirectory)
            Me.OffsetOfStartOfCentralDirectoryWithRespectToTheStartingDiskNumber = BitHelper.ReadUInt32(zip.FileBits, offset + FieldOffset.OffsetOfStartOfCentralDirectoryWithRespectToTheStartingDiskNumber)
            Me.ZipFileCommentLength = BitHelper.ReadUInt16(zip.FileBits, offset + FieldOffset.ZipFileCommentLength)
            Me.ZipFileComment = BitHelper.ReadBytes(zip.FileBits, offset + FieldOffset.ZipFileComment, Me.ZipFileCommentLength)

            Me.Length = FieldOffset.ZipFileComment + Me.ZipFileCommentLength
        End Sub

        Public Shared Function Find(zip As ZipArchive) As ZipEndOfCentralDirectoryRecord
            Dim offset As Integer = zip.FileBits.Length - FieldOffset.ZipFileComment
            While offset >= 0
                If BitHelper.ReadUInt32(zip.FileBits, offset) = ZipSignature.EndOfCentralDirectory Then
                    Exit While
                End If
                offset = offset - 1
            End While
            If offset < 0 Then
                Return Nothing
            Else
                Return New ZipEndOfCentralDirectoryRecord(zip, offset)
            End If
        End Function
    End Class

    ''' <summary>
    ''' 从文件创建Zip压缩包
    ''' </summary>
    ''' <param name="files"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Compress(files As String()) As Byte()
        Dim zip As New ZipArchive
        zip.Zip(files)
        Return zip.FileBits
    End Function

    ''' <summary>
    ''' 从对象创建Zip压缩包
    ''' </summary>
    ''' <param name="files"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Compress(files As ZipFileInfo()) As Byte()
        Dim zip As New ZipArchive
        zip.Zip(files)
        Return zip.FileBits
    End Function

    ''' <summary>
    ''' 解压Zip压缩包内的文件名称
    ''' </summary>
    ''' <param name="file"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DecompressNames(file As String) As String()
        Dim zip As New ZipArchive
        zip.Open(file)
        Return zip.CentralDirectoryFileDict.Keys.ToArray
    End Function

    ''' <summary>
    ''' 解压Zip压缩包内的所有文件
    ''' </summary>
    ''' <param name="file"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DecompressAll(file As String) As ZipFileInfo()
        Dim zip As New ZipArchive
        zip.Open(file)
        Return zip.UnZip
    End Function

    ''' <summary>
    ''' 解压Zip压缩包内的单个文件
    ''' </summary>
    ''' <param name="file"></param>
    ''' <param name="name"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DecompressOne(file As String, name As String) As ZipFileInfo
        Dim zip As New ZipArchive
        zip.Open(file)
        Return zip.UnZip(name)
    End Function
End Class
