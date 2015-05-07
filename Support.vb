Imports System
Imports System.Threading
Imports System.Configuration
Imports Fleck

Public Class Websocket
    Dim allSockets = New List(Of IWebSocketConnection)()
    Dim SocketNumber As Integer = 0

    Private Sub OnWebSocketOpen(socket)
        Console.WriteLine("Open!")
        allSockets.Add(socket)
        SocketNumber += 1
    End Sub

    Private Sub OnWebsocketClose(socket)
        Console.WriteLine("Close!")
        allSockets.Remove(socket)
        SocketNumber -= 1
    End Sub

    Private Sub OnMessage(socket As IWebSocketConnection, message As String)
        socket.Send("Echo: " + message)
        Broadcasting("已收到" + Now.ToString)
    End Sub

    Public Sub Broadcasting(message As String)
        If SocketNumber = 0 Then
            Exit Sub
        End If
        For Each s As IWebSocketConnection In allSockets
            s.Send(message)
        Next
    End Sub

    Public Sub InitWebsocket(url As String)
        FleckLog.Level = Fleck.LogLevel.Error
        Dim server = New WebSocketServer(url)
        Try
            server.Start(Sub(socket)
                             socket.OnOpen = Sub() OnWebSocketOpen(socket)
                             socket.OnClose = Sub() OnWebsocketClose(socket)
                             socket.OnMessage = Sub(message) OnMessage(socket, message)
                         End Sub)
        Catch ex As Exception
            ConsoleControl.Log("Websocket监听失败", ConsoleControl.ConsoleLogLevel.Debug)
        End Try

    End Sub

End Class 'Websocket控制类

Class ConsoleControl
    Enum ConsoleLogLevel As Integer
        Errors = 0
        Warning = 1
        Info = 2
        Debug = 3
    End Enum

    Private Shared ConsoleLogLevelString() As String = {"Error", "Warning", "Info", "Debug"}
    Public Shared Property Level As ConsoleLogLevel = ConsoleLogLevel.Errors

    Public Shared Sub Log(MessageString As String, MessageLevel As ConsoleLogLevel)
        If MessageLevel <= Level Then
            Console.WriteLine(Now() + "  " + "[" + ConsoleLogLevelString(MessageLevel) + "]  " + MessageString)
        End If
    End Sub

    Public Shared Sub StartInfo()
        Dim StartInfo As String = "CTP接收服务器" + vbCrLf + _
        "Powered by Backlighting" + vbCrLf + _
         vbCrLf + _
        "启动时间：" + Now() + vbCrLf + _
        "版本编译时间：" + Support.GetSettings("ComplieTime") + vbCrLf + _
        "配置信息：" + vbCrLf + _
        "    CTP配置信息：" + vbCrLf + _
        "        帐号：" + Support.GetSettings("CTPUsername") + vbCrLf + _
        "        配置文件：" + Support.GetSettings("CTPConfigFile") + vbCrLf + _
        "        当前合约：" + Support.GetSettings("CTPInstrumentID") + vbCrLf + _
        "    Websocket监听地址：" + Support.GetSettings("WebsocketURL") + vbCrLf + _
        "    策略文件位置：" + Support.GetSettings("TemplateFilePath") + vbCrLf + _
        "    止损设置：" + Support.GetSettings("Point_StopLoss") + vbCrLf + _
        "    CTP数据输出位置：" + Support.GetSettings("CTPOutput") + vbCrLf + _
        "    JJC-MMC-KPC输出位置：" + Support.GetSettings("TemplateOutput") + vbCrLf + _
        "    结果输出位置：" + Support.GetSettings("ResultOutput") + vbCrLf + _
        "    回测文件位置：" + Support.GetSettings("RecallFilepath") + vbCrLf + _
        "    开仓限制：" + Support.GetSettings("ID_Start_AM") + " ~ " + Support.GetSettings("ID_End_AM") + "   " _
        + Support.GetSettings("ID_Start_PM") + " ~ " + Support.GetSettings("ID_End_PM") + vbCrLf + _
        "    强制平仓：" + Support.GetSettings("ID_Force_Offset") + vbCrLf + _
        "    滑点设置：(多开 多平 空开 空平)  " + Support.GetSettings("Point_Offset_BullOpen") + "  " + Support.GetSettings("Point_Offset_BullOffset") + "  " _
        + Support.GetSettings("Point_Offset_BearOpen") + "  " + Support.GetSettings("Point_Offset_BearOffset") + vbCrLf + _
        "    交易手数：" + Support.GetSettings("OrderVolume") + vbCrLf + _
        vbCrLf + _
        "启动实时模式请输入real，启动回测模式请输入recall"

        Console.WriteLine(StartInfo)
    End Sub
End Class '控制台控制类

Class SQLServer
    Public Property ConnectionString As String
    Private SQLConnection As SqlClient.SqlConnection
    Private SQLCommand As SqlClient.SqlCommand

    Public Function Connect()
        Try
            SQLConnection = New SqlClient.SqlConnection(ConnectionString)
            SQLConnection.Open()
        Catch ex As Exception
            Return -1
        End Try
        Return 0
    End Function
End Class '数据库控制类

Class Support
    Public Shared Function GetOutputNumber0(Number As Double) As String
        Return Format(Number, "0")
    End Function

    Public Shared Function GetOutputNumber1(Number As Double) As String
        Return Format(Number, "0.0")
    End Function

    Public Shared Function GetSettings(SettingName As String) As String
        Return System.Configuration.ConfigurationManager.AppSettings(SettingName)
    End Function

    Public Shared Sub PlaySound(SettingItem As String)
        Dim SoundPlayer As System.Media.SoundPlayer = New System.Media.SoundPlayer()
        SoundPlayer.SoundLocation = Support.GetSettings(SettingItem)
        SoundPlayer.Play()
    End Sub

    Public Shared Function MyTxtReader(ByVal StrPath As String) As String
        MyTxtReader = ""
        If IO.File.Exists(StrPath) = True Then
            Dim TxtReader As IO.StreamReader = New IO.StreamReader(StrPath, System.Text.Encoding.Default)
            MyTxtReader = TxtReader.ReadToEnd
            TxtReader.Close()
        End If
    End Function

    Public Shared Function WriteTextToFile(ByVal StrPath As String, ByVal Message As String)
        Try
            Dim file As New System.IO.StreamWriter(StrPath, False, Text.Encoding.UTF8)
            file.WriteLine(Message)
            file.Close()
        Catch
            Return -1
        End Try
        Return 0
    End Function

    Public Shared Function ID2Time(ByVal ID As Integer) As String
        Dim Result As DateTime
        If ID < 2700 Then
            Result = New DateTime(Now.Year, Now.Month, Now.Day, 9, 15, 0)
            Result = Result.AddSeconds(3 * ID)
        Else
            Result = New DateTime(Now.Year, Now.Month, Now.Day, 13, 0, 0)
            Result = Result.AddSeconds(3 * (ID - 2700))
        End If
        Return Result.ToLongTimeString
    End Function

    Public Shared Function Time2ID(ByVal Time As String) As Integer
        Dim temp() As String
        temp = Time.Split(":")
        Dim Hour, Minutes, Second As Integer
        Hour = CInt(temp(0))
        Minutes = CInt(temp(1))
        Second = CInt(temp(2))
        If (Hour = 9 And Minutes < 15) Or _
            (Hour < 9) Or _
            (Hour = 11 And Minutes = 30 And Second = 0) Or _
            (Hour = 15 And Minutes = 15 And Second = 0) Then
            Return -1
        End If

        If Hour < 12 Then
            Return Int(((Hour - 9) * 3600 + (Minutes - 15) * 60 + Second) / 3)
        Else
            Return Int(((Hour - 13) * 3600 + Minutes * 60 + Second) / 3 + 2700)
        End If
    End Function

    Public Shared Sub RequestExit()
        Console.WriteLine()
        Console.WriteLine("按任意键退出 ...")
        Console.ReadKey()
    End Sub
End Class '支持函数类

