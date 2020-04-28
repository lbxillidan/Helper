Public Class BitHelper
    Public Enum BitEndian As Integer
        ''' <summary>
        ''' 1234=3400+12
        ''' </summary>
        ''' <remarks></remarks>
        LittleEndian = 1
        ''' <summary>
        ''' 1234=1200+34
        ''' </summary>
        ''' <remarks></remarks>
        BigEndian = 2

        ZipEndian = LittleEndian
    End Enum

    Public Shared Endian As BitEndian = BitEndian.LittleEndian

    Private Shared Bit4H As Byte = Convert.ToByte("11110000", 2)
    Private Shared Bit4L As Byte = Convert.ToByte("00001111", 2)
    Public Shared Function ReadBit4H(buffer() As Byte, startIndex As Integer) As Byte
        Dim Rst As Byte
        Rst = (buffer(startIndex) And Bit4H) >> 4
        Return Rst
    End Function
    Public Shared Function ReadBit4L(buffer() As Byte, startIndex As Integer) As Byte
        Dim Rst As Byte
        Rst = buffer(startIndex) And Bit4L
        Return Rst
    End Function

    Public Shared Function ReadByte(buffer() As Byte, startIndex As Integer) As Byte
        Return buffer(startIndex)
    End Function
    Public Shared Function ReadBytes(buffer As Byte(), offset As Integer, length As Integer) As Byte()
        Dim Rst(length - 1) As Byte
        Array.Copy(buffer, offset, Rst, 0, length)
        Return Rst
    End Function

    Public Shared Function ReadInt32(buffer As Byte(), offset As Integer) As Int32
        Dim Bit0 As Int32, Bit1 As Int32, Bit2 As Int32, Bit3 As Int32
        Bit0 = buffer(offset)
        Bit1 = buffer(offset + 1)
        Bit2 = buffer(offset + 2)
        Bit3 = buffer(offset + 3)

        If Endian = BitEndian.LittleEndian Then
            Bit1 = Bit1 << 8
            Bit2 = Bit2 << 16
            Bit3 = Bit3 << 24
        Else
            Bit0 = Bit0 << 24
            Bit1 = Bit1 << 16
            Bit2 = Bit2 << 8
        End If

        Dim Rst As Int32 = Bit0 + Bit1 + Bit2 + Bit3
        Return Rst
    End Function
    Public Shared Function ReadInt16(buffer As Byte(), offset As Integer) As Int16
        Dim Bit0 As Int16, Bit1 As Int16
        Bit0 = buffer(offset)
        Bit1 = buffer(offset + 1)

        If Endian = BitEndian.LittleEndian Then
            Bit1 = Bit1 << 8
        Else
            Bit0 = Bit0 << 8
        End If

        Dim Rst As Int16 = Bit0 + Bit1
        Return Rst
    End Function

    Public Shared Function ReadUInt32(buffer As Byte(), offset As Integer) As UInt32
        Dim Bit0 As UInt32, Bit1 As UInt32, Bit2 As UInt32, Bit3 As UInt32
        Bit0 = buffer(offset)
        Bit1 = buffer(offset + 1)
        Bit2 = buffer(offset + 2)
        Bit3 = buffer(offset + 3)

        If Endian = BitEndian.LittleEndian Then
            Bit1 = Bit1 << 8
            Bit2 = Bit2 << 16
            Bit3 = Bit3 << 24
        Else
            Bit0 = Bit0 << 24
            Bit1 = Bit1 << 16
            Bit2 = Bit2 << 8
        End If

        Dim Rst As UInt32 = Bit0 + Bit1 + Bit2 + Bit3
        Return Rst
    End Function
    Public Shared Function ReadUInt16(buffer As Byte(), offset As Integer) As UInt16
        Dim Bit0 As UInt16, Bit1 As UInt16
        Bit0 = buffer(offset)
        Bit1 = buffer(offset + 1)

        If Endian = BitEndian.LittleEndian Then
            Bit1 = Bit1 << 8
        Else
            Bit0 = Bit0 << 8
        End If

        Dim Rst As UInt16 = Bit0 + Bit1
        Return Rst
    End Function

    Public Shared Sub WriteByte(buffer() As Byte, startIndex As Integer, value As Byte)
        buffer(startIndex) = value
    End Sub
    Public Shared Sub WriteBytes(buffer As Byte(), offset As Integer, value As Byte())
        Array.Copy(value, 0, buffer, offset, value.Length)
    End Sub

    Public Shared Sub WriteUInt32(buffer As Byte(), offset As Integer, value As UInt32)
        Dim Bit0 As UInt32, Bit1 As UInt32, Bit2 As UInt32, Bit3 As UInt32
        If Endian = BitEndian.LittleEndian Then
            Bit0 = (value << 24) >> 24
            Bit1 = (value << 16) >> 24
            Bit2 = (value << 8) >> 24
            Bit3 = value >> 24
        Else
            Bit3 = (value << 24) >> 24
            Bit2 = (value << 16) >> 24
            Bit1 = (value << 8) >> 24
            Bit0 = value >> 24
        End If
        buffer(offset) = Bit0
        buffer(offset + 1) = Bit1
        buffer(offset + 2) = Bit2
        buffer(offset + 3) = Bit3
    End Sub
    Public Shared Sub WriteUInt16(buffer As Byte(), offset As Integer, value As UInt16)
        Dim Bit0 As UInt16, Bit1 As UInt16
        If Endian = BitEndian.LittleEndian Then
            Bit0 = (value << 8) >> 8
            Bit1 = value >> 8
        Else
            Bit1 = (value << 8) >> 8
            Bit0 = value >> 8
        End If
        buffer(offset) = Bit0
        buffer(offset + 1) = Bit1
    End Sub

    Public Shared Sub InsertBytes(ByRef buffer() As Byte, startIndex As Integer, value() As Byte)
        Dim BufferCopy() As Byte = buffer.Clone
        Dim BufferLength As Integer = buffer.Length
        Dim ValueLength As Integer = value.Length
        ReDim buffer(BufferLength + ValueLength - 1)
        Array.Copy(BufferCopy, 0, buffer, 0, startIndex)
        Array.Copy(value, 0, buffer, startIndex, ValueLength)
        Array.Copy(BufferCopy, startIndex, buffer, startIndex + ValueLength, BufferLength - startIndex)
    End Sub

    Public Shared Sub RemoveBytes(ByRef buffer() As Byte, startIndex As Integer, length As Integer)
        Dim BufferCopy() As Byte = buffer.Clone
        Dim BufferLength As Integer = buffer.Length
        ReDim buffer(BufferLength - length - 1)
        Array.Copy(BufferCopy, 0, buffer, 0, startIndex)
        Array.Copy(BufferCopy, startIndex + length, buffer, startIndex, BufferLength - startIndex - length)
    End Sub

    Public Shared Sub ReplaceBytes(ByRef buffer() As Byte, startIndex As Integer, length As Integer, value() As Byte)
        Dim BufferCopy() As Byte = buffer.Clone
        Dim BufferLength As Integer = buffer.Length
        Dim ValueLength As Integer = value.Length
        ReDim buffer(BufferLength + ValueLength - 1)
        Array.Copy(BufferCopy, 0, buffer, 0, startIndex)
        Array.Copy(value, 0, buffer, startIndex, ValueLength)
        Array.Copy(BufferCopy, startIndex + length, buffer, startIndex + ValueLength, BufferLength - startIndex - length)
    End Sub
End Class
