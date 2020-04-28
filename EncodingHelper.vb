Public Class EncodingHelper
    Private Enum AsciiChar As Byte
        NUL = &H0
        SOH = &H1
        BEL = &H7
        BS = &H8
        CR = &HD
        SO = &HE
        US = &H1F
        Space = &H20
        Tilde = &H7E
        DEL = &H7F
    End Enum

    ''' <summary>
    ''' 检查是否全部属于Ascii字符，文本范围为0x08-0x0D,0x20-0x7E，完整范围为0x00-0x7F
    ''' </summary>
    ''' <param name="data"></param>
    ''' <param name="onlyText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsASCII(data As Byte(), Optional onlyText As Boolean = True) As Boolean
        If data Is Nothing Then
            Return False
        End If
        Dim Bound As Integer = data.Length - 1
        If Bound = -1 Then
            Return False
        End If

        Dim i As Integer

        If onlyText = True Then
            While i <= Bound
                Select Case data(i)
                    Case AsciiChar.BS To AsciiChar.CR
                        '控制字符
                        i = i + 1
                    Case AsciiChar.Space To AsciiChar.Tilde
                        '文本字符
                        i = i + 1
                    Case Else
                        Return False
                End Select
            End While
        Else
            While i <= Bound
                Select Case data(i)
                    Case AsciiChar.NUL To AsciiChar.DEL
                        'Ascii字符
                        i = i + 1
                    Case Else
                        Return False
                End Select
            End While
        End If

        Return True
    End Function

    ''' <summary>
    ''' 检查是否全部属于GBK字符，完整范围8140-FEFE剔除XX7F，文本范围排除自定义区及非文本Ascii
    ''' </summary>
    ''' <param name="data"></param>
    ''' <param name="onlyText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsGBK(data As Byte(), Optional onlyText As Boolean = True) As Boolean
        If data Is Nothing Then
            Return False
        End If
        Dim Bound As Integer = data.Length - 1
        If Bound = -1 Then
            Return False
        End If

        Dim Between = Function(n As Byte, nMin As Byte, nMax As Byte) (n >= nMin AndAlso n <= nMax)

        Dim i As Integer = 0

        If onlyText = True Then
            While i <= Bound
                Select Case data(i)
                    Case AsciiChar.NUL To AsciiChar.BEL, AsciiChar.SO To AsciiChar.US, AsciiChar.DEL
                        Return False
                    Case AsciiChar.BS To AsciiChar.CR, AsciiChar.Space To AsciiChar.Tilde
                        i = i + 1
                    Case &H81 To &HFE
                        'GBK:8140-FEFE，剔除XX7F，共23940个。
                        If i = Bound Then
                            Return False
                        End If
                        If data(i + 1) = &H7F Then
                            Return False
                        End If
                        Dim Bit1 As Byte = data(i)
                        Dim Bit2 As Byte = data(i + 1)
                        'GBK1
                        If Between(Bit1, &HA1, &HA9) AndAlso Between(Bit2, &HA1, &HFE) Then
                            i = i + 2
                            Continue While
                        End If
                        'GBK2
                        If Between(Bit1, &HB0, &HF7) AndAlso Between(Bit2, &HA1, &HFE) Then
                            i = i + 2
                            Continue While
                        End If
                        'GBK3
                        If Between(Bit1, &H81, &HA0) AndAlso Between(Bit2, &H40, &HFE) Then
                            i = i + 2
                            Continue While
                        End If
                        'GBK4
                        If Between(Bit1, &HAA, &H40) AndAlso Between(Bit2, &HFE, &HA0) Then
                            i = i + 2
                            Continue While
                        End If
                        'GBK5
                        If Between(Bit1, &HA8, &H40) AndAlso Between(Bit2, &HA9, &HA0) Then
                            i = i + 2
                            Continue While
                        End If
                        'GBK自定义1
                        If Between(Bit1, &HAA, &HAF) AndAlso Between(Bit2, &HA1, &HFE) Then
                            Return False
                        End If
                        'GBK自定义2
                        If Between(Bit1, &HF8, &HFE) AndAlso Between(Bit2, &HA1, &HFE) Then
                            Return False
                        End If
                        'GBK自定义3
                        If Between(Bit1, &HA1, &HA7) AndAlso Between(Bit2, &H40, &HA0) Then
                            Return False
                        End If
                    Case Else
                        Return False
                End Select
            End While
        Else
            While i <= Bound
                Select Case data(i)
                    Case AsciiChar.NUL To AsciiChar.DEL
                        i = i + 1
                    Case &H81 To &HFE
                        'GBK:8140-FEFE，剔除XX7F，共23940个。
                        If i = Bound Then
                            Return False
                        End If
                        Select Case data(i + 1)
                            Case &H40 To &H7E, &H80 To &HFE
                                i = i + 2
                            Case Else
                                Return False
                        End Select
                    Case Else
                        Return False
                End Select
            End While
        End If

        Return True
    End Function

    Public Shared Function IsUTF8(data As Byte(), Optional onlyText As Boolean = True) As Boolean
        If data Is Nothing Then
            Return False
        End If
        Dim Bound As Integer = data.Length - 1
        If Bound = -1 Then
            Return False
        End If
        If Bound >= 2 AndAlso data(0) = &HEF AndAlso data(1) = &HBB AndAlso data(2) = &HBF Then
            Return True
        End If

        Dim Between = Function(n As Byte, nMin As Byte, nMax As Byte) (n >= nMin AndAlso n <= nMax)

        Dim i As Integer = 0

        If onlyText = True Then
            While i <= Bound
                Select Case data(i)
                    Case AsciiChar.NUL To AsciiChar.BEL, AsciiChar.SO To AsciiChar.US, AsciiChar.DEL
                        Return False
                    Case AsciiChar.BS To AsciiChar.CR, AsciiChar.Space To AsciiChar.Tilde
                        i = i + 1
                    Case &HC0 To &HDF
                        '110xxxxx 10xxxxxx
                        If i + 1 > Bound Then
                            Return False
                        Else
                            If Between(data(i + 1), &H80, &HBF) Then
                                i = i + 2
                            Else
                                Return False
                            End If
                        End If
                    Case &HE0 To &HEF
                        '1110xxxx 10xxxxxx 10xxxxxx 
                        If i + 2 > Bound Then
                            Return False
                        Else
                            If Between(data(i + 1), &H80, &HBF) AndAlso Between(data(i + 2), &H80, &HBF) Then
                                i = i + 3
                            Else
                                Return False
                            End If
                        End If
                    Case &HF0 To &HF7
                        '11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
                        If i + 3 > Bound Then
                            Return False
                        Else
                            If Between(data(i + 1), &H80, &HBF) AndAlso Between(data(i + 2), &H80, &HBF) AndAlso Between(data(i + 3), &H80, &HBF) Then
                                i = i + 4
                            Else
                                Return False
                            End If
                        End If
                    Case Else
                        Return False
                End Select
            End While
        Else
            While i <= Bound
                Select Case data(i)
                    Case AsciiChar.NUL To AsciiChar.DEL
                        i = i + 1
                    Case &HC0 To &HDF
                        '110xxxxx 10xxxxxx
                        If i + 1 > Bound Then
                            Return False
                        Else
                            If Between(data(i + 1), &H80, &HBF) Then
                                i = i + 2
                            Else
                                Return False
                            End If
                        End If
                    Case &HE0 To &HEF
                        '1110xxxx 10xxxxxx 10xxxxxx 
                        If i + 2 > Bound Then
                            Return False
                        Else
                            If Between(data(i + 1), &H80, &HBF) AndAlso Between(data(i + 2), &H80, &HBF) Then
                                i = i + 3
                            Else
                                Return False
                            End If
                        End If
                    Case &HF0 To &HF7
                        '11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
                        If i + 3 > Bound Then
                            Return False
                        Else
                            If Between(data(i + 1), &H80, &HBF) AndAlso Between(data(i + 2), &H80, &HBF) AndAlso Between(data(i + 3), &H80, &HBF) Then
                                i = i + 4
                            Else
                                Return False
                            End If
                        End If
                    Case Else
                        Return False
                End Select
            End While
        End If

        Return True
    End Function

End Class
