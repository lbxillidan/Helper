Public Class PrintCounter
    Private Shared Browser As New WebBrowserEx '= Form1.WebBrowser1

    Public Shared Function M1216(ip As String) As Integer
        Dim Count As Integer = 0
        If Browser.WaitForNavigate(String.Format("http://{0}/SSI/info_configuration.htm", ip), 5) = False Then
            Return 0
        End If

        Dim Html As String = Browser.HtmlText
        If String.IsNullOrEmpty(Html) Then
            Return 0
        End If
        If Html.Contains("打印的总页数") = False Then
            Return 0
        End If

        Html = Html.Replace(vbCrLf, String.Empty).Replace(vbCr, String.Empty).Replace(vbLf, String.Empty)
        'Debug.Print(Html)
        Dim Reg As New System.Text.RegularExpressions.Regex("打印的总页数:</td> <td[^<>]+>(\d+)</td>")
        Dim Mats As System.Text.RegularExpressions.MatchCollection = Reg.Matches(Html)
        If Mats.Count = 1 Then
            Integer.TryParse(Mats(0).Groups(1).Value, Count)
        End If

        'Debug.Print(Count)
        Return Count
    End Function

    Public Shared Function KM350(ip As String) As Integer
        Dim Count As Integer = 0
        If Browser.WaitForNavigate(String.Format("http://{0}/m_s_cnt.htm", ip), 5) = False Then
            Return 0
        End If

        Dim Html As String = Browser.HtmlText
        If String.IsNullOrEmpty(Html) Then
            Return 0
        End If
        If Html.Contains("<TITLE>系统-计数器</TITLE>") = False Then
            Return 0
        End If

        Html = Html.Replace(vbCrLf, String.Empty).Replace(vbCr, String.Empty).Replace(vbLf, String.Empty)
        Dim Reg As New System.Text.RegularExpressions.Regex("<TD[^<>]+>总的</TD><TD[^<>]+>(\d+)</TD>")
        Dim Mats As System.Text.RegularExpressions.MatchCollection = Reg.Matches(Html)
        If Mats.Count = 1 Then
            Integer.TryParse(Mats(0).Groups(1).Value, COunt)
        End If

        'Debug.Print(Count)
        Return Count
    End Function

    Public Shared Function KM3010(ip As String) As Integer
        Dim Count As Integer = 0
        If Browser.WaitForNavigate(String.Format("http://{0}/m_s_cnt.htm", ip), 5) = False Then
            Return 0
        End If

        Dim Html As String = Browser.HtmlText
        If String.IsNullOrEmpty(Html) Then
            Return 0
        End If
        If Html.Contains("<TITLE>System - Counter</TITLE>") = False Then
            Return 0
        End If
        'Debug.Print(Html)
        Html = Html.Replace(vbCrLf, String.Empty).Replace(vbCr, String.Empty).Replace(vbLf, String.Empty)
        Dim Reg As New System.Text.RegularExpressions.Regex("<TD COLSPAN=""3"" HEADERS=""Total"">(\d+)</TD>")
        Dim Mats As System.Text.RegularExpressions.MatchCollection = Reg.Matches(Html)
        If Mats.Count = 1 Then
            Integer.TryParse(Mats(0).Groups(1).Value, Count)
        End If

        'Debug.Print(Count)
        Return Count
    End Function

    Public Shared Function KM363(ip As String) As Integer
        Dim Count As Integer = 0

        Dim CookieCheck = Function()
                              If Browser.Document Is Nothing Then
                                  Return False
                              End If
                              Dim Cookie As String = Browser.Document.Cookie
                              If String.IsNullOrEmpty(Cookie) Then
                                  Return False
                              End If
                              If Cookie.Contains("ID=") Then
                                  Return True
                              Else
                                  Return False
                              End If
                          End Function

        '页面测试
        If Browser.WaitForNavigate(String.Format("http://{0}/wcd/top.xml", ip), 5) = False Then
            Return 0
        End If
        'Dim Html As String = Browser.HtmlText
        'Debug.Print(Html)
        '重置Cookie
        Browser.Document.Cookie = String.Empty
        '刷新页面
        Browser.WaitForNavigate(String.Format("http://{0}/wcd/top.xml", ip), 5)
        '选择语言
        Dim LangSelectElement As HtmlElement
        For Each Element As HtmlElement In Browser.Document.GetElementsByTagName("select")
            If Element.GetAttribute("name") = "Lang" Then
                LangSelectElement = Element
                Element.SetAttribute("selectedIndex", 2)
                Exit For
            End If
        Next
        '点击登录
        Dim SubmitElement As HtmlElement
        For Each Element As HtmlElement In Browser.Document.GetElementsByTagName("input")
            If Element.GetAttribute("type") = "submit" Then
                SubmitElement = Element
                Element.InvokeMember("click")
                WebBrowserEx.Wait(CookieCheck, 500, 5000)
                Exit For
            End If
        Next
        'Debug.Print(Browser.Document.Cookie)
        '计数器页面
        Browser.WaitForNavigate(String.Format("http://{0}/wcd/system_counter.xml", ip), 5)
        Dim Html As String = Browser.HtmlText
        If String.IsNullOrEmpty(Html) Then
            Return 0
        End If
        'Debug.Print(Html)
        '关键词测试
        If Html.Contains("'us_Total_1'") = False Then
            Return 0
        End If
        'Debug.Print(Html)
        '移除换行
        Html = Html.Replace(vbCrLf, String.Empty).Replace(vbCr, String.Empty).Replace(vbLf, String.Empty)
        '查找数量
        Dim Reg As New System.Text.RegularExpressions.Regex("'us_Total_1'[^<>]+>Total</span></th><td[^<>]+>(\d+)<br></td>")
        Dim Mats As System.Text.RegularExpressions.MatchCollection = Reg.Matches(Html)
        If Mats.Count = 1 Then
            Integer.TryParse(Mats(0).Groups(1).Value, Count)
        End If

        'Debug.Print(Count)
        Return Count
    End Function

    Public Shared Function MF4700(ip As String) As Integer
        Dim Count As Integer = 0

        Dim UrlCheck = Function()
                           If Browser.ReadyState < WebBrowserReadyState.Complete Then
                               Return False
                           End If
                           If Browser.Url Is Nothing Then
                               Return False
                           End If
                           Dim Url As String = Browser.Url.ToString
                           If String.IsNullOrEmpty(Url) Then
                               Return False
                           End If
                           If Url.Contains("d_status") Then
                               Return True
                           Else
                               Return False
                           End If
                       End Function

        '页面测试
        If Browser.WaitForNavigate(String.Format("http://{0}/t_welcom.cgi?page=Language_name&lang=1", ip), 5) = False Then
            Return 0
        End If
        'Dim Html As String = Browser.HtmlText
        'Debug.Print(Html)
        '重置Cookie
        Debug.Print(Browser.Document.Cookie)
        Browser.Document.Cookie = String.Empty
        '刷新页面
        Browser.WaitForNavigate(String.Format("http://{0}/t_welcom.cgi?page=Language_name&lang=1", ip), 5)
        '点击登录
        Dim SubmitElement As HtmlElement
        For Each Element As HtmlElement In Browser.Document.GetElementsByTagName("input")
            If Element.GetAttribute("name") = "OK" Then
                SubmitElement = Element
                Element.InvokeMember("click")
                WebBrowserEx.Wait(UrlCheck, 500, 5000)
                Exit For
            End If
        Next
        '计数器页面
        Browser.WaitForNavigate(String.Format("http://{0}/d_count.cgi", ip), 5)
        Dim Html As String = Browser.HtmlText
        If String.IsNullOrEmpty(Html) Then
            Return 0
        End If
        'Debug.Print(Html)
        '关键词测试
        If Html.Contains("<h3>Check Counter</h3>") = False Then
            Return 0
        End If
        'Debug.Print(Html)
        '移除换行
        Html = Html.Replace(vbCrLf, String.Empty).Replace(vbCr, String.Empty).Replace(vbLf, String.Empty)
        '查找数量
        Dim Reg As New System.Text.RegularExpressions.Regex("<td>(\d+)</td>")
        Dim Mats As System.Text.RegularExpressions.MatchCollection = Reg.Matches(Html)
        If Mats.Count = 1 Then
            Integer.TryParse(Mats(0).Groups(1).Value, Count)
        End If

        'Debug.Print(Count)
        Return Count
    End Function

End Class
