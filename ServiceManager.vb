Imports System.ServiceProcess

Public Class ServiceManager
    'Public Shared IvdSystemPath As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "IvdSystem")
    Public Message As String
    Public ServiceName As String
    Public ServiceFile As String
    Public ServiceArgs As String
    Public DependFiles As String()
    Private CommandResultFile As String

    Private Function SetServiceFile() As Boolean
        If String.IsNullOrEmpty(Me.ServiceFile) Then
            '从注册表读取服务信息
            Dim ServiceKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Services\" + ServiceName)
            If ServiceKey IsNot Nothing Then
                Dim ImagePath As String = ServiceKey.GetValue("ImagePath")
                If String.IsNullOrEmpty(ImagePath) = False Then
                    If ImagePath.StartsWith("""") Then
                        Dim Idx As Integer = ImagePath.IndexOf("""", 1)
                        If Idx > 1 Then
                            Me.ServiceFile = ImagePath.Substring(1, Idx - 1)
                        End If
                    Else
                        Dim Idx As Integer = ImagePath.ToLower.IndexOf(".exe")
                        If Idx > 1 Then
                            Me.ServiceFile = ImagePath.Substring(0, Idx + 4)
                        End If
                    End If
                End If
            End If
        End If
        If String.IsNullOrEmpty(Me.ServiceFile) Then
            Message = "获取ServiceFile失败！"
            Return False
        Else
            Return True
        End If
    End Function

    Private Function SetCommandResultFile() As Boolean
        If String.IsNullOrEmpty(Me.CommandResultFile) Then
            If SetServiceFile() = False Then
                Return False
            Else
                Me.CommandResultFile = Me.ServiceFile & ".CommandResult.txt"
                Return True
            End If
        Else
            Return True
        End If
    End Function

    Public Shared Function GetServiceController(svcName As String) As System.ServiceProcess.ServiceController
        Dim ServiceController As System.ServiceProcess.ServiceController = Nothing
        If String.IsNullOrEmpty(svcName) = False Then
            For Each scTemp As System.ServiceProcess.ServiceController In System.ServiceProcess.ServiceController.GetServices()
                If scTemp.ServiceName.ToLower = svcName.ToLower Then
                    ServiceController = scTemp
                    Exit For
                End If
            Next
        End If
        Return ServiceController
    End Function

    Public Function GetServiceState() As System.ServiceProcess.ServiceControllerStatus
        Dim ServiceController As System.ServiceProcess.ServiceController = GetServiceController(ServiceName)
        If ServiceController Is Nothing Then
            Me.Message = "服务没有安装。"
            Return 0
        Else
            Me.Message = String.Format("服务状态：{0}", ServiceController.Status.ToString)
            Return ServiceController.Status
        End If
    End Function

    Private Function CreateServiceFiles(Optional srcFile As String = Nothing) As Boolean
        If SetServiceFile() = False Then
            Return False
        End If

        '默认使用当前文件
        If String.IsNullOrEmpty(srcFile) Then
            srcFile = System.Diagnostics.Process.GetCurrentProcess.MainModule.FileName
        End If
        '原位安装
        If ServiceFile.ToLower = srcFile.ToLower Then
            Return True
        End If
        '复制文件
        Try
            '复制核心文件
            Dim ServicePath As String = System.IO.Path.GetDirectoryName(Me.ServiceFile)
            If System.IO.Directory.Exists(ServicePath) = False Then
                System.IO.Directory.CreateDirectory(ServicePath)
            End If
            System.IO.File.Copy(srcFile, ServiceFile, True)
            '复制辅助文件
            If DependFiles IsNot Nothing Then
                Dim SrcRootPath As String = System.IO.Path.GetDirectoryName(srcFile)
                For Each DependFileName As String In DependFiles
                    Dim DependFile As String = System.IO.Path.Combine(ServicePath, DependFileName)
                    Dim SourceFile As String = System.IO.Path.Combine(SrcRootPath, DependFileName)
                    Dim DependPath As String = System.IO.Path.GetDirectoryName(DependFile)
                    If System.IO.Directory.Exists(DependPath) = False Then
                        System.IO.Directory.CreateDirectory(DependPath)
                    End If
                    System.IO.File.Copy(SourceFile, DependFile, True)
                Next
            End If
            Return True
        Catch ex As Exception
            Me.Message = "复制服务文件失败：" & ex.Message
            Return False
        End Try
    End Function

    Private Function DeleteServiceFiles() As Boolean
        If SetServiceFile() = False Then
            Return False
        End If
        Try
            If System.IO.File.Exists(ServiceFile) Then
                System.IO.File.Delete(ServiceFile)
            End If
            If DependFiles IsNot Nothing Then
                Dim ServicePath As String = System.IO.Path.GetDirectoryName(Me.ServiceFile)
                For Each DependFileName As String In DependFiles
                    Dim DependFile As String = System.IO.Path.Combine(ServicePath, DependFileName)
                    If System.IO.File.Exists(DependFile) Then
                        System.IO.File.Delete(DependFile)
                    End If
                Next
            End If
            Return True
        Catch ex As Exception
            Me.Message = "删除服务文件失败：" & ex.Message
            Return False
        End Try
    End Function

    Public Function InstallService(Optional srcFile As String = Nothing) As Boolean
        '检查权限
        Dim Principal As New System.Security.Principal.WindowsPrincipal(System.Security.Principal.WindowsIdentity.GetCurrent)
        If Principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator) = False Then
            Me.Message = "安装服务需要管理权限！"
            Return False
        End If
        '卸载服务
        If UninstallService() = False Then
            '保留Message
            Return False
        End If
        '复制文件
        If CreateServiceFiles(srcFile) = False Then
            '保留Message
            Return False
        End If
        '安装服务
        Process.Start("sc.exe", String.Format("create {0} binPath= ""\""{1}\"" {2}"" start= auto", ServiceName, ServiceFile, ServiceArgs)).WaitForExit()
        '启动服务
        Process.Start("sc.exe", String.Format("start {0}", ServiceName)).WaitForExit()
        '检查结果
        Dim ServiceController As System.ServiceProcess.ServiceController = GetServiceController(ServiceName)
        If ServiceController IsNot Nothing Then
            Try
                ServiceController.WaitForStatus(ServiceProcess.ServiceControllerStatus.Running, New TimeSpan(0, 0, 10))
            Catch ex As Exception
                Me.Message = "服务启动错误：" & ex.Message
                Return False
            End Try
            If ServiceController.Status = ServiceProcess.ServiceControllerStatus.Running Then
                Me.Message = "服务安装完成！"
                Return True
            Else
                Me.Message = "服务启动失败！"
                Return False
            End If
        Else
            Me.Message = "服务安装失败！"
            Return False
        End If
    End Function

    Public Function UninstallService() As Boolean
        '检查权限
        Dim Principal As New System.Security.Principal.WindowsPrincipal(System.Security.Principal.WindowsIdentity.GetCurrent)
        If Principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator) = False Then
            Me.Message = "卸载服务需要管理权限！"
            Return False
        End If
        '检查服务
        Dim ServiceController As System.ServiceProcess.ServiceController = GetServiceController(ServiceName)
        If ServiceController Is Nothing Then
            Me.Message = "服务没有安装！"
            Return True
        End If
        '停止服务
        Process.Start("sc.exe", String.Format("stop {0}", ServiceName)).WaitForExit()
        '卸载服务
        Process.Start("sc.exe", String.Format("delete {0}", ServiceName)).WaitForExit()
        '删除文件
        Dim DeleteFilesResult As String
        If DeleteServiceFiles() = True Then
            DeleteFilesResult = "文件删除完成！"
        Else
            DeleteFilesResult = "文件删除失败！"
        End If
        '检查结果
        ServiceController = GetServiceController(ServiceName)
        If ServiceController Is Nothing Then
            Me.Message = "服务卸载完成！" & DeleteFilesResult
            Return True
        Else
            Me.Message = "服务卸载失败！" & DeleteFilesResult
            Return False
        End If
    End Function

    Public Function StartService() As Boolean
        '检查权限
        Dim Principal As New System.Security.Principal.WindowsPrincipal(System.Security.Principal.WindowsIdentity.GetCurrent)
        If Principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator) = False Then
            Me.Message = "启动服务需要管理权限！"
            Return False
        End If
        '检查服务
        Dim ServiceController As System.ServiceProcess.ServiceController = GetServiceController(ServiceName)
        If ServiceController Is Nothing Then
            Me.Message = "服务没有安装！"
            Return False
        End If
        '启动服务
        Try
            ServiceController.Start()
            ServiceController.WaitForStatus(ServiceProcess.ServiceControllerStatus.Running, New TimeSpan(0, 0, 10))
        Catch ex As Exception
            Me.Message = "服务启动错误：" & ex.Message
            Return False
        End Try
        '检查结果
        If ServiceController.Status = ServiceProcess.ServiceControllerStatus.Running Then
            Me.Message = "服务启动完成！"
            Return True
        Else
            Me.Message = "服务启动失败！"
            Return False
        End If
    End Function

    Public Function StopService() As Boolean
        '检查权限
        Dim Principal As New System.Security.Principal.WindowsPrincipal(System.Security.Principal.WindowsIdentity.GetCurrent)
        If Principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator) = False Then
            Me.Message = "停止服务需要管理权限！"
            Return False
        End If
        '检查服务
        Dim ServiceController As System.ServiceProcess.ServiceController = GetServiceController(ServiceName)
        If ServiceController Is Nothing Then
            Me.Message = "服务没有安装！"
            Return False
        End If
        '停止服务
        Try
            ServiceController.Stop()
            ServiceController.WaitForStatus(ServiceProcess.ServiceControllerStatus.Stopped, New TimeSpan(0, 0, 10))
        Catch ex As Exception
            Me.Message = "服务停止错误：" & ex.Message
            Return False
        End Try
        '检查结果
        If ServiceController.Status = ServiceProcess.ServiceControllerStatus.Stopped Then
            Me.Message = "服务停止完成！"
            Return True
        Else
            Me.Message = "服务停止失败！"
            Return False
        End If
    End Function

    Private Function WaitForDelCommandResultFile() As Boolean
        Dim TimeStart As Date = Date.Now
        While True
            If System.IO.File.Exists(CommandResultFile) = False Then
                Return True
            End If

            If Date.Now.Subtract(TimeStart).TotalMilliseconds > 1000 Then
                Exit While
            Else
                System.Threading.Thread.Sleep(100)
            End If
        End While
        Return False
    End Function
    Private Function WaitForAddCommandResultFile() As Boolean
        Dim TimeStart As Date = Date.Now
        While True
            If System.IO.File.Exists(CommandResultFile) = True Then
                Return True
            End If

            If Date.Now.Subtract(TimeStart).TotalMilliseconds > 60 * 1000 Then
                Exit While
            Else
                System.Threading.Thread.Sleep(1000)
            End If
        End While
        Return False
    End Function

    Public Function ExecuteCommand(command As Integer, waitResult As Boolean) As Boolean
        '检查服务
        Dim ServiceController As System.ServiceProcess.ServiceController = GetServiceController(ServiceName)
        If ServiceController Is Nothing OrElse ServiceController.Status <> ServiceProcess.ServiceControllerStatus.Running Then
            Me.Message = "服务未安装或未运行！"
            Return False
        End If
        '执行命令，无需返回结果
        If waitResult = False Then
            ServiceController.ExecuteCommand(command)
            Return True
        End If
        '执行命令，等待返回结果
        '检查结果文件路径
        If SetCommandResultFile() = False Then
            Me.Message = "获取CommandResultFile失败！"
            Return False
        End If
        '删除结果文件
        ServiceController.ExecuteCommand(CustomCommand.DelResult)
        If WaitForDelCommandResultFile() = False Then
            Me.Message = "清空命令结果失败！"
            Return False
        End If
        '执行命令
        ServiceController.ExecuteCommand(command)
        If WaitForAddCommandResultFile() = False Then
            Me.Message = "读取命令结果失败！"
            Return False
        End If
        '读取结果文件
        Me.Message = System.IO.File.ReadAllText(CommandResultFile)
        Return True
    End Function

#Region "仅供服务使用"
    Public Enum CustomCommand As Integer
        DelResult = 255
    End Enum

    Public Function DelCommandResult() As Boolean
        If SetCommandResultFile() = False Then
            Message = "获取CommandResultFile失败！"
            Return False
        End If

        Try
            If System.IO.File.Exists(CommandResultFile) Then
                System.IO.File.Delete(CommandResultFile)
            End If
            Return True
        Catch ex As Exception
            Me.Message = "删除结果文件失败：" & ex.Message
            Return False
        End Try
    End Function

    Public Function AddCommandResult(rst As String) As Boolean
        If String.IsNullOrEmpty(rst) Then
            Return True
        End If
        If SetCommandResultFile() = False Then
            Message = "获取CommandResultFile失败！"
            Return False
        End If
        Try
            System.IO.File.WriteAllText(CommandResultFile, Date.Now.ToString("yyyy-MM-dd HH:mm:ss.ff") & vbCrLf & rst)
            Return True
        Catch ex As Exception
            Message = "写入CommandResultFile失败！"
            Return False
        End Try
    End Function
#End Region
End Class

