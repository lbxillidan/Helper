Public Class CRC32Helper
    Public Enum CRC32Seed As UInt32
        '04C11DB7
        IEEE802d3 = &HEDB88320UI
        '82608EDB

        '1EDC6F41
        Castagnoli = &H82F63B78UI
        '8F6E37A0

        '741B8CD7
        Koopman = &HEB31D82EUI
        'BA0DC66B
    End Enum

    Private Shared ReadOnly Table(256) As UInt32
    ''' <summary>
    ''' 设置MagicNumber,默认0xEDB88320
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Shared WriteOnly Property Seed As UInt32
        Set(value As UInt32)
            For i As UInt32 = 0UI To 255UI
                Dim r As UInt32 = i
                For j As UInt32 = 0UI To 7UI
                    If (r And 1UI) <> 0UI Then
                        r = (r >> 1) Xor value
                    Else
                        r >>= 1
                    End If
                Next
                Table(i) = r
            Next
        End Set
    End Property

    Shared Sub New()
        Seed = &HEDB88320UI
    End Sub

    Public Shared Function ComputeAsUInt32(buffer As Byte(), offset As UInt32, count As UInt32) As UInt32
        If count = 0 Then
            Return 0
        End If

        Dim TempUInt As UInt32 = UInt32.MaxValue
        Dim TempByte As Byte
        For i As UInt32 = 0 To count - 1
            TempByte = TempUInt Mod 256UI
            TempUInt = Table(TempByte Xor buffer(offset + i)) Xor (TempUInt >> 8)
        Next
        Return TempUInt Xor UInt32.MaxValue
    End Function

    Public Shared Function ComputeAsUInt32(buffer As Byte()) As UInt32
        Return ComputeAsUInt32(buffer, 0, buffer.Length)
    End Function

    Public Shared Function ComputeAsUInt32(file As String) As UInt32
        Dim buffer As Byte() = System.IO.File.ReadAllBytes(file)
        Return ComputeAsUInt32(buffer, 0, buffer.Length)
    End Function

    Public Shared Function ComputeAsBytes(buffer As Byte(), offset As UInt32, count As UInt32) As Byte()
        Dim TempBytes As Byte() = BitConverter.GetBytes(ComputeAsUInt32(buffer, offset, count))
        Array.Reverse(TempBytes)
        Return TempBytes
    End Function
    Public Shared Function ComputeAsBytes(buffer As Byte()) As Byte()
        Return ComputeAsBytes(buffer, 0, buffer.Length)
    End Function
    Public Shared Function ComputeAsBytes(file As String) As Byte()
        Dim buffer As Byte() = System.IO.File.ReadAllBytes(file)
        Return ComputeAsBytes(buffer, 0, buffer.Length)
    End Function

    Public Shared Function ComputeAsString(buffer As Byte(), offset As UInt32, count As UInt32) As String
        Return BitConverter.ToString(ComputeAsBytes(buffer, offset, count)).Replace("-", "")
    End Function
    Public Shared Function ComputeAsString(buffer As Byte()) As String
        Return ComputeAsString(buffer, 0, buffer.Length)
    End Function
    Public Shared Function ComputeAsString(file As String) As String
        Dim buffer As Byte() = System.IO.File.ReadAllBytes(file)
        Return ComputeAsString(buffer, 0, buffer.Length)
    End Function
End Class
