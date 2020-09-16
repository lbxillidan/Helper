
Public Module StreamExtension
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As UInt64)
        Dim Size As Integer = 8
        Dim Bits(Size - 1) As Byte
        Stream.Read(Bits, 0, Size)
        Value = System.BitConverter.ToUInt64(Bits, 0)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As UInt32)
        Dim Size As Integer = 4
        Dim Bits(Size - 1) As Byte
        Stream.Read(Bits, 0, Size)
        Value = System.BitConverter.ToUInt32(Bits, 0)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As UInt16)
        Dim Size As Integer = 2
        Dim Bits(Size - 1) As Byte
        Stream.Read(Bits, 0, Size)
        Value = System.BitConverter.ToUInt16(Bits, 0)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As Byte)
        Value = Stream.ReadByte
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As Int64)
        Dim Size As Integer = 8
        Dim Bits(Size - 1) As Byte
        Stream.Read(Bits, 0, Size)
        Value = System.BitConverter.ToInt64(Bits, 0)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As Int32)
        Dim Size As Integer = 4
        Dim Bits(Size - 1) As Byte
        Stream.Read(Bits, 0, Size)
        Value = System.BitConverter.ToInt32(Bits, 0)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As Int16)
        Dim Size As Integer = 2
        Dim Bits(Size - 1) As Byte
        Stream.Read(Bits, 0, Size)
        Value = System.BitConverter.ToInt16(Bits, 0)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As SByte)
        Dim UByte As Byte
        Stream.Read(UByte)
        If UByte > 127 Then
            Value = CInt(UByte) - 256
        Else
            Value = UByte
        End If
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As Double)
        Dim Size As Integer = 8
        Dim Bits(Size - 1) As Byte
        Stream.Read(Bits, 0, Size)
        Value = System.BitConverter.ToDouble(Bits, 0)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As Single)
        Dim Size As Integer = 4
        Dim Bits(Size - 1) As Byte
        Stream.Read(Bits, 0, Size)
        Value = System.BitConverter.ToSingle(Bits, 0)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As Boolean)
        Dim Size As Integer = 1
        Dim Bits(Size - 1) As Byte
        Stream.Read(Bits, 0, Size)
        Value = System.BitConverter.ToBoolean(Bits, 0)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As Guid)
        Dim Size As Integer = 16
        Dim Bits(Size - 1) As Byte
        Stream.Read(Bits, 0, Size)
        Value = New Guid(Bits)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As Date)
        Dim Size As Integer = 8
        Dim Bits(Size - 1) As Byte
        Stream.Read(Bits, 0, Size)
        Value = Date.FromFileTime(System.BitConverter.ToInt64(Bits, 0))
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As Byte())
        Stream.Read(Value, 0, Value.Length)
    End Sub

    <Flags>
    Public Enum ReadStringOption As Int32
        BreakType_FixedBytes = &H1000000
        BreakType_FixedChars = &H2000000
        BreakType_NullTerminate = &H4000000

        Encoding_Unicode = &H10000000
        Encoding_Default = &H20000000
    End Enum

    ''' <summary>
    ''' 读取文本。Options需要提供Encoding、BreakType和Length。
    ''' </summary>
    ''' <param name="Stream"></param>
    ''' <param name="Value"></param>
    ''' <param name="Options"></param>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Read(ByRef Stream As System.IO.Stream, ByRef Value As String, Options As ReadStringOption)
        Dim Encoding As System.Text.Encoding
        Select Case Options And &HF0000000
            Case ReadStringOption.Encoding_Unicode
                Encoding = System.Text.Encoding.Unicode
            Case ReadStringOption.Encoding_Default
                Encoding = System.Text.Encoding.Default
            Case Else
                Throw New Exception("Options缺少Encoding。")
        End Select

        Dim BreakType As ReadStringOption = Options And &HF000000
        Dim Length As Integer = Options And &HFFFF

        Select Case BreakType
            Case ReadStringOption.BreakType_FixedBytes
                If Length = 0 Then
                    Value = String.Empty
                Else
                    Dim Buffer(Length - 1) As Byte
                    Stream.Read(Buffer, 0, Length)
                    Value = Encoding.GetString(Buffer)
                End If
            Case ReadStringOption.BreakType_FixedChars
                If Length = 0 Then
                    Value = String.Empty
                Else
                    Dim OriginOffset As Int64 = Stream.Position
                    Dim RemainLength As Int64 = Stream.Length - Stream.Position
                    Dim BufferLength As Int32 = Encoding.GetMaxByteCount(Length)
                    If BufferLength > RemainLength Then
                        BufferLength = RemainLength
                    End If
                    Dim Buffer(BufferLength - 1) As Byte
                    Stream.Read(Buffer, 0, BufferLength)
                    Value = Encoding.GetString(Buffer).Substring(0, Length)
                    Stream.Position = OriginOffset + Encoding.GetByteCount(Value)
                End If
            Case ReadStringOption.BreakType_NullTerminate
                Dim OriginOffset As Int64 = Stream.Position
                Dim RemainLength As Int64 = Stream.Length - Stream.Position
                Dim BufferLength As Int32 = 4
                Dim BufferChars(BufferLength) As Char
                Dim BufferBytes(BufferLength - 1) As Byte
                Dim BytesRead As Integer
                Dim CharsRead As Integer
                Dim IndexNull As Integer

                Dim Builder As New System.Text.StringBuilder
                Dim Decoder As System.Text.Decoder = Encoding.GetDecoder

                While RemainLength > 0L
                    BytesRead = Stream.Read(BufferBytes, 0, BufferLength)
                    CharsRead = Decoder.GetChars(BufferBytes, 0, BytesRead, BufferChars, 0, False)
                    RemainLength = RemainLength - BytesRead
                    IndexNull = Array.IndexOf(BufferChars, Char.MinValue)
                    If IndexNull = -1 OrElse IndexNull >= CharsRead Then
                        Builder.Append(New String(BufferChars, 0, CharsRead))
                    Else
                        Builder.Append(New String(BufferChars, 0, IndexNull))
                        Exit While
                    End If
                End While

                Value = Builder.ToString
                Stream.Position = OriginOffset + Encoding.GetByteCount(Value & Char.MinValue)
            Case Else
                Throw New Exception("Options缺少BreakType。")
        End Select
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub ReadEnum(ByRef Stream As System.IO.Stream, ByRef Value As System.Enum)
        Select Case Value.GetType.GetEnumUnderlyingType
            Case GetType(System.Int64)
                Dim BaseValue As Int64
                Stream.Read(BaseValue)
                Value = [Enum].ToObject(Value.GetType, BaseValue)
            Case GetType(System.Int32)
                Dim BaseValue As Int32
                Stream.Read(BaseValue)
                Value = [Enum].ToObject(Value.GetType, BaseValue)
            Case GetType(System.Int16)
                Dim BaseValue As Int16
                Stream.Read(BaseValue)
                Value = [Enum].ToObject(Value.GetType, BaseValue)
            Case GetType(System.SByte)
                Dim BaseValue As SByte
                Stream.Read(BaseValue)
                Value = [Enum].ToObject(Value.GetType, BaseValue)
            Case GetType(System.UInt64)
                Dim BaseValue As UInt64
                Stream.Read(BaseValue)
                Value = [Enum].ToObject(Value.GetType, BaseValue)
            Case GetType(System.UInt32)
                Dim BaseValue As UInt32
                Stream.Read(BaseValue)
                Value = [Enum].ToObject(Value.GetType, BaseValue)
            Case GetType(System.UInt16)
                Dim BaseValue As UInt16
                Stream.Read(BaseValue)
                Value = [Enum].ToObject(Value.GetType, BaseValue)
            Case GetType(System.Byte)
                Dim BaseValue As Byte
                Stream.Read(BaseValue)
                Value = [Enum].ToObject(Value.GetType, BaseValue)
        End Select
    End Sub



    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As UInt64)
        Stream.Write(System.BitConverter.GetBytes(Value), 0, 8)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As UInt32)
        Stream.Write(System.BitConverter.GetBytes(Value), 0, 4)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As UInt16)
        Stream.Write(System.BitConverter.GetBytes(Value), 0, 2)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As Byte)
        Stream.WriteByte(Value)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As Int64)
        Stream.Write(System.BitConverter.GetBytes(Value), 0, 8)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As Int32)
        Stream.Write(System.BitConverter.GetBytes(Value), 0, 4)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As Int16)
        Stream.Write(System.BitConverter.GetBytes(Value), 0, 2)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As SByte)
        Dim UByte As Byte
        If Value < 0 Then
            UByte = Value + 256
        Else
            UByte = Value
        End If
        Stream.WriteByte(UByte)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As Double)
        Stream.Write(System.BitConverter.GetBytes(Value), 0, 8)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As Single)
        Stream.Write(System.BitConverter.GetBytes(Value), 0, 4)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As Boolean)
        Stream.Write(System.BitConverter.GetBytes(Value), 0, 1)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As Guid)
        Stream.Write(Value.ToByteArray, 0, 16)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As Date)
        Stream.Write(System.BitConverter.GetBytes(Value.ToFileTime), 0, 8)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As Byte())
        Stream.Write(Value, 0, Value.Length)
    End Sub
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Write(ByRef Stream As System.IO.Stream, Value As String, Encoding As System.Text.Encoding)
        Dim Buffer As Byte() = Encoding.GetBytes(Value)
        Stream.Write(Buffer, 0, Buffer.Length)
    End Sub

End Module
