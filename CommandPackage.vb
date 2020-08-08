Public Class CommandPackage
    Public Enum PkgValueType As Integer
        [Auto] = 0
        [Bytes] = 1
        [Boolean] = 2
        [Integer] = 3
        [String] = 4
    End Enum

    Public Enum RstType As Integer
        Success = 1
        Failure = 2
    End Enum

    Public Class PkgValue
        '00-03/04 Length
        '04-07/04 ValueType
        '08-xx/xx ValueData
        Private Const HeadLength As Integer = 8
        Public Bits As Byte()
        Public Length As Integer
        Public ValueType As PkgValueType
        Public ValueData As Object

        Sub New(bytes As Byte(), offset As Integer)
            If bytes Is Nothing Then
                Throw New ArgumentNullException("bytes")
            End If
            If bytes.Length < offset + HeadLength Then
                Throw New FormatException("bytes长度不足。")
            End If
            Me.Length = System.BitConverter.ToInt32(bytes, offset)
            If bytes.Length < offset + Me.Length Then
                Throw New FormatException("bytes长度错误。")
            End If

            ReDim Me.Bits(Me.Length - 1)
            Array.Copy(bytes, offset, Me.Bits, 0, Me.Length)

            Me.ValueType = System.BitConverter.ToInt32(Me.Bits, 4)
            Dim ValueOffset As Integer = HeadLength
            Dim ValueLength As Integer = Me.Length - HeadLength

            Select Case ValueType
                Case PkgValueType.Boolean
                    Me.ValueData = System.BitConverter.ToBoolean(Me.Bits, ValueOffset)
                Case PkgValueType.Integer
                    Me.ValueData = System.BitConverter.ToInt32(Me.Bits, ValueOffset)
                Case PkgValueType.String
                    Me.ValueData = System.Text.Encoding.UTF8.GetString(Me.Bits, ValueOffset, ValueLength)
                Case Else
                    Dim ValueBits(ValueLength - 1) As Byte
                    Array.Copy(Me.Bits, ValueOffset, ValueBits, 0, ValueLength)
                    Me.ValueData = ValueBits
            End Select
        End Sub

        Sub New(valueType As PkgValueType, valueData As Object)
            If valueType = PkgValueType.Auto Then
                Select Case valueData.GetType
                    Case GetType(Boolean)
                        valueType = PkgValueType.Boolean
                    Case GetType(Integer)
                        valueType = PkgValueType.Integer
                    Case GetType(String)
                        valueType = PkgValueType.String
                    Case GetType(Byte())
                        valueType = PkgValueType.Bytes
                    Case Else
                        Throw New NotSupportedException("valueData格式不受支持：" & valueData.GetType.ToString)
                End Select
            End If

            Dim ValueBits() As Byte = New Byte() {}
            Select Case valueType
                Case PkgValueType.Boolean
                    ValueBits = System.BitConverter.GetBytes(valueData)
                Case PkgValueType.Integer
                    ValueBits = System.BitConverter.GetBytes(valueData)
                Case PkgValueType.String
                    ValueBits = System.Text.Encoding.UTF8.GetBytes(valueData)
                Case PkgValueType.Bytes
                    ValueBits = valueData
            End Select
            Me.Length = ValueBits.Length + HeadLength
            ReDim Me.Bits(Me.Length - 1)
            Me.ValueType = valueType
            Me.ValueData = valueData
            Array.Copy(System.BitConverter.GetBytes(Me.Length), 0, Me.Bits, 0, 4)
            Array.Copy(System.BitConverter.GetBytes(Me.ValueType), 0, Me.Bits, 4, 4)
            Array.Copy(ValueBits, 0, Me.Bits, HeadLength, ValueBits.Length)
        End Sub
    End Class

    Private Const HeadLength As Integer = 12
    Public Bits As Byte()
    Public Length As Integer
    Public CmdRst As Integer
    Public ValueCount As Integer
    Public Values As PkgValue()

    '00-03/04 Length
    '04-07/04 CmdRst
    '08-11/04 ValueCount
    Sub New(bytes As Byte())
        If bytes Is Nothing Then
            Throw New ArgumentNullException("bytes")
        End If
        If bytes.Length < HeadLength Then
            Throw New FormatException("bytes长度不足。")
        End If
        Me.Length = System.BitConverter.ToInt32(bytes, 0)
        If bytes.Length <> Me.Length Then
            Throw New FormatException("bytes长度错误。")
        End If
        Me.CmdRst = System.BitConverter.ToInt32(bytes, 4)
        Me.ValueCount = System.BitConverter.ToInt32(bytes, 8)
        Me.Bits = bytes
        ReDim Me.Values(Me.ValueCount - 1)

        If Me.ValueCount > 0 Then
            Dim Offset As Integer = HeadLength
            For Idx As Integer = 0 To Me.ValueCount - 1
                Dim ValueTemp As New PkgValue(bytes, Offset)
                Me.Values(Idx) = ValueTemp
                Offset = Offset + ValueTemp.Length
            Next
        End If
    End Sub

    Sub New(pkgCmdRst As Integer, ParamArray pkgValueDatas As Object())
        Me.Length = HeadLength
        Me.CmdRst = pkgCmdRst
        Me.ValueCount = pkgValueDatas.Length
        ReDim Me.Values(Me.ValueCount - 1)
        For Idx As Integer = 0 To Me.ValueCount - 1
            Dim ValueTemp As New PkgValue(PkgValueType.Auto, pkgValueDatas(Idx))
            Me.Values(Idx) = ValueTemp
            Me.Length = Me.Length + ValueTemp.Length
        Next
        ReDim Me.Bits(Me.Length - 1)
        Array.Copy(System.BitConverter.GetBytes(Me.Length), 0, Me.Bits, 0, 4)
        Array.Copy(System.BitConverter.GetBytes(Me.CmdRst), 0, Me.Bits, 4, 4)
        Array.Copy(System.BitConverter.GetBytes(Me.ValueCount), 0, Me.Bits, 8, 4)
        Dim Offset As Integer = HeadLength
        For Each ValueTemp As PkgValue In Me.Values
            Array.Copy(ValueTemp.Bits, 0, Me.Bits, Offset, ValueTemp.Length)
            Offset = Offset + ValueTemp.Length
        Next
    End Sub

    Public Class CommandResult
        Public Rst As RstType
        Public Values As Object()
        Public ErrMsg As String
    End Class

    Public Shared Function HttpCommand(ByRef web As System.Net.WebClient, url As String, cmd As Integer, ParamArray params As Object()) As CommandResult
        Dim Rst As New CommandResult
        Try
            Dim CmdPkg As New CommandPackage(cmd, params)
            web.Headers.Set("Cmd", cmd.ToString)
            Dim RstBits As Byte() = web.UploadData(url, CmdPkg.Bits)
            Dim RstPkg As New CommandPackage(RstBits)

            If RstPkg.CmdRst = CommandPackage.RstType.Success Then
                Dim RstValues(RstPkg.ValueCount - 1) As Object
                For i As Integer = 0 To RstPkg.ValueCount - 1
                    RstValues(i) = RstPkg.Values(i).ValueData
                Next
                Rst.Rst = RstType.Success
                Rst.Values = RstValues
                Return Rst
            Else
                Rst.Rst = RstType.Failure
                If RstPkg.ValueCount > 0 Then
                    Rst.ErrMsg = RstPkg.Values(0).ValueData.ToString
                Else
                    Rst.ErrMsg = "HttpCommand返回失败结果。"
                End If
                Return Rst
            End If
        Catch ex As Exception
            Rst.Rst = RstType.Failure
            Rst.ErrMsg = ex.ToString
            Return Rst
        End Try
    End Function

    Public Class HttpCommandContext
        Public User As String
        Public EndPoint As String
        Public HttpContext As System.Net.HttpListenerContext
    End Class
End Class
