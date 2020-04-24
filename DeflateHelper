Public Class DeflateHelper
    ''' <summary>
    ''' Deflate解压
    ''' </summary>
    ''' <param name="encodeData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Decode(encodeData As Byte()) As Byte()
        Dim EncodeStream As New System.IO.MemoryStream(encodeData)
        Dim Deflate As New System.IO.Compression.DeflateStream(EncodeStream, System.IO.Compression.CompressionMode.Decompress)
        Deflate.Flush()
        Dim DecodeStream As New System.IO.MemoryStream()
        Deflate.CopyTo(DecodeStream)
        Deflate.Close()
        EncodeStream.Close()
        Dim DecodeData As Byte() = DecodeStream.ToArray
        DecodeStream.Close()
        Return DecodeData
    End Function

    ''' <summary>
    ''' Deflate压缩
    ''' </summary>
    ''' <param name="decodeData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Encode(decodeData As Byte()) As Byte()
        Dim EncodeStream As New System.IO.MemoryStream()
        Dim Deflate As New System.IO.Compression.DeflateStream(EncodeStream, System.IO.Compression.CompressionMode.Compress)
        Deflate.Write(decodeData, 0, decodeData.Length)
        Deflate.Close()
        Dim EncodeData As Byte() = EncodeStream.ToArray
        EncodeStream.Close()
        Return EncodeData
    End Function
End Class
