Imports CTPCOMLib
Imports System.Threading

Module Main
    Dim WithEvents ctp As ICTPClientAPI

    Dim WithEvents Core As SPIF_Core = New SPIF_Core

    Dim ws As Websocket = New Websocket
    Dim sql As SQLServer = New SQLServer
    Dim real As Realtime = New Realtime
    Dim TemplateObject As SPIF_Template = New SPIF_Template

    Dim TempBuffer As Realtime.CTPBufferMetaStructure

    Dim CTPSourceData As String = "合约编号,买一价,买一量,卖一价,卖一量,开盘价,最高价,最低价,最新价,持仓量,成交量,涨停价,跌停价,昨结算价,今日平均价,行情更新时间,更新毫秒数" + vbCrLf
    Public TemplateSourceData As String = "ID,Time,JN,JJC,MMC,KPC" + vbCrLf
    Dim Result As String = String.Empty

    Dim CurrentExpectionCommand As Integer = 1
    Dim CurrentExpectionDirection As Integer = -1
    Dim CurrentOrderPrice As Double = 0
    Dim CurrentOrderBenefit As Double = 0
    Dim CurrentOrderID As Integer = 0
    Dim OrderCount As Integer = 0

    Dim FindedTemplateNumber As Integer = -1
    Dim Benefit As Double = 0

    Dim IsSimulation As Boolean = False

    Sub Main()

        ConsoleControl.StartInfo()

        Dim UserCommand As String

        '=====核心参数设置=====
        ConsoleControl.Level = ConsoleControl.ConsoleLogLevel.Debug

        If Core.ChangeTemplate(Support.GetSettings("TemplateFilePath")) = -1 Then
            ConsoleControl.Log("Policy文件编译失败", ConsoleControl.ConsoleLogLevel.Debug)
            Call Support.RequestExit()
        End If

        Core.ID_Start_AM = Support.GetSettings("ID_Start_AM")
        Core.ID_End_AM = Support.GetSettings("ID_End_AM")
        Core.ID_Start_PM = Support.GetSettings("ID_Start_PM")
        Core.ID_End_PM = Support.GetSettings("ID_End_PM")
        Core.Point_StopLoss = Support.GetSettings("Point_StopLoss")
        Core.ID_ForceOffset = Support.GetSettings("ID_Force_Offset")
        Core.Point_Offset_BullOpen = Support.GetSettings("Point_Offset_BullOpen")
        Core.Point_Offset_BullOffset = Support.GetSettings("Point_Offset_BullOffset")
        Core.Point_Offset_BearOpen = Support.GetSettings("Point_Offset_BearOpen")
        Core.Point_Offset_BearOffset = Support.GetSettings("Point_Offset_BearOffset")
        Core.OrderVolume = Support.GetSettings("OrderVolume")
        Core.ID_OrderGap = Support.GetSettings("ID_OrderGap")
        '=====核心参数设置=====

        Do
            Console.WriteLine()
            Console.Write(">  ")
            UserCommand = Console.ReadLine
            
            Select Case UserCommand
                Case "exit"
                    Console.Write("确定要退出吗？(y/n) ")
                    If Console.ReadLine = "y" Then
                        Exit Do
                    End If

                Case "real"  '启动CTP实时
                    Call CommandInReal()

                Case "recall"  '启动回测
                    Call Command()

                Case "ct"  '中途更换策略
                    Console.Write("请输入策略文件位置：（输入q退出） ")
                    Dim FilePath As String = Console.ReadLine
                    If FilePath <> "q" Then
                        Call Core.ChangeTemplate(FilePath)
                    End If

                Case "cld"  '中途锁定方向
                    Console.Write("请输入要限定的开仓方向： （1=多;2=空;-1=不限）")
                    Dim Direction As String = CInt(Console.ReadLine)
                    Call Core.ChangeLockingDirection(Direction)

                Case "octp"
                    If Core.OutputCTP() = 0 Then
                        Console.Write("要立刻打开这个文件吗？(y/n) ")
                        If Console.ReadLine = "y" Then
                            Console.WriteLine("Excel正在启动…… ")
                            Process.Start(Core.Config_Default_Output_CTP)
                            Console.WriteLine("请转置Excel查看")
                        End If
                    End If

                Case "ocus"
                    If Core.OutputCustom() = 0 Then
                        Console.Write("要立刻打开这个文件吗？(y/n) ")
                        If Console.ReadLine = "y" Then
                            Console.WriteLine("Excel正在启动…… ")
                            Process.Start(Core.Config_Default_Output_Custom)
                            Console.WriteLine("请转置Excel查看")
                        End If
                    End If

                Case "otrd"
                    If Core.OutputTrade() = 0 Then
                        Console.Write("要立刻打开这个文件吗？(y/n) ")
                        If Console.ReadLine = "y" Then
                            Process.Start(Core.Config_Default_Output_Trade)
                        End If
                    End If

                Case Else
                    Console.WriteLine("未能识别的命令")
            End Select
        Loop
        Call Support.RequestExit()
    End Sub

    Function Command()
        Core.PushOrder = New SPIF_Core.PushOrderDelegate(AddressOf OnOrder)

        Dim Filepath = Support.GetSettings("RecallFilepath")
        Dim CTPDataString As String = Support.MyTxtReader(Filepath)
        Dim RowData() As String = CTPDataString.Split(vbCrLf)
        CTPDataString = Nothing
        Dim CTPData(RowData.Length)() As String

        IsSimulation = True

        Dim Updatetime As String
        Dim UpdatemilliSecond, AskVolume1, BidVolume1, OpenInterest As Integer
        Dim LastPrice As Double

        Dim RecallDataType As String = String.Empty

        Console.WriteLine("本次回测的文件位置：" + Support.GetSettings("RecallFilepath"))
        Console.Write("回测CTP原始数据请输入1，回测Dafuweng数据请输入2        >  ")
        Select Case Console.ReadLine()
            Case "1"
                RecallDataType = "CTP"
            Case "2"
                RecallDataType = "Dafuweng"
            Case Else
                End
        End Select

        Dim TempString As String = String.Empty

        For Counter = 2 To RowData.Length - 3
            CTPData(Counter) = RowData(Counter).Split(",")
            If RecallDataType = "CTP" Then
                Updatetime = CTPData(Counter)(15)
                UpdatemilliSecond = CTPData(Counter)(16)
                AskVolume1 = CInt(CTPData(Counter)(4))
                BidVolume1 = CInt(CTPData(Counter)(2))
                OpenInterest = CInt(CTPData(Counter)(9))
                LastPrice = CDbl(CTPData(Counter)(8))
            Else
                Updatetime = CTPData(Counter)(2).Substring(11, 8)
                UpdatemilliSecond = 0
                AskVolume1 = CInt(CTPData(Counter)(15))
                BidVolume1 = CInt(CTPData(Counter)(14))
                OpenInterest = CInt(CTPData(Counter)(4))
                LastPrice = CDbl(CTPData(Counter)(3))
            End If
            Core.DataArrive("", 0, BidVolume1, 0, AskVolume1, 0, 0, 0, LastPrice, OpenInterest, 0, 0, 0, 0, 0, Updatetime, UpdatemilliSecond)
        Next
        ConsoleControl.Log("回测完成 ", ConsoleControl.ConsoleLogLevel.Debug)
        Return 0
    End Function

    Private Sub OnOrder(Command As Integer, Direction As Integer, _
                                 Price As Double, Volume As Integer, Reason As String, OrderFakeID As Integer) 
        ConsoleControl.Log(Core.CurrentID.ToString + " " + " 【" + Format(Price, "0.0") + "】 " + " 【" + Reason + "】", ConsoleControl.ConsoleLogLevel.Debug)

        If Command = 0 Then
            Console.WriteLine(Core.GetCurrentOrderInfo)
            Console.WriteLine()
            Console.WriteLine()
        End If

        If Support.GetSettings("IsMute") = "False" Then
            If Command = 1 AndAlso Direction = 1 Then
                Support.PlaySound("SoundOfBullOpen")
            ElseIf Command = 0 AndAlso Direction = 1 Then
                Support.PlaySound("SoundOfBullOffset")
            ElseIf Command = 1 AndAlso Direction = 0 Then
                Support.PlaySound("SoundOfBearOpen")
            ElseIf Command = 0 AndAlso Direction = 0 Then
                Support.PlaySound("SoundOfBearOffset")
            End If
        End If

    End Sub

    Function CommandInReal()
        IsSimulation = False
        Core.PushOrder = New SPIF_Core.PushOrderDelegate(AddressOf OnOrder)

        ctp = New CTPCOMLib.ICTPClientAPI
        Dim ErrorID As Integer
        ConsoleControl.Log("CTP尝试登录", ConsoleControl.ConsoleLogLevel.Info)
        ctp.Login(Support.GetSettings("CTPConfigFile"), Support.GetSettings("CTPUsername"), Support.GetSettings("CTPPassword"), ErrorID)
        If ErrorID <> 0 Then
            ConsoleControl.Log("登陆失败，错误代号：" + ErrorID.ToString, ConsoleControl.ConsoleLogLevel.Errors)
            Return -1
        End If

        ConsoleControl.Log("CTP服务器已连接，等待初始化", ConsoleControl.ConsoleLogLevel.Info)
        Dim InstrumentID As String = Support.GetSettings("CTPInstrumentID")
        Core.Instrument = InstrumentID
        ctp.SubscribeMD(InstrumentID)
        ConsoleControl.Log("行情数据已订阅", ConsoleControl.ConsoleLogLevel.Info)
        'ws.InitWebsocket(Support.GetSettings("WebsocketURL"))
        'ConsoleControl.Log("Websocket已建立", ConsoleControl.ConsoleLogLevel.Info)

        Return 0
    End Function

    'Private Sub CTP_OnInitFinished()
    '    ConsoleControl.Log("CTP初始化完成", ConsoleControl.ConsoleLogLevel.Info)
    '    Dim InstrumentArray As Array
    '    Dim InstrumentCount As Int16
    '    Dim InstrumentID As String = ""
    '    If Support.GetSettings("CTPInstrumentID") <> "default" Then
    '        InstrumentID = Support.GetSettings("CTPInstrumentID")
    '    Else
    '        ctp.GetInstruments(InstrumentArray, InstrumentCount)
    '        For Each Instrument As CTPCOMLib.InstrumentField In InstrumentArray
    '            If Instrument.ProductID = "IF" Then
    '                InstrumentID = Instrument.InstrumentID
    '                Exit For
    '            End If
    '        Next
    '    End If
    '    ctp.SubscribeMD(InstrumentID)
    '    ConsoleControl.Log("CTP开始接收" + InstrumentID + "的行情数据", ConsoleControl.ConsoleLogLevel.Info)
    'End Sub 


    Private Sub CTP_OnMarketData(ByVal InstrumentID As String, _
                             ByVal BidPrice1 As Double, ByVal BidVolume1 As Integer, _
                             ByVal AskPrice1 As Double, ByVal AskVolume1 As Integer, _
                             ByVal OpenPrice As Double, ByVal HighestPrice As Double, ByVal LowestPrice As Double, _
                             ByVal LastPrice As Double, ByVal OpenInterest As Integer, ByVal Volume As Integer, _
                             ByVal UpperLimitPrice As Double, ByVal LowerLimitPrice As Double, _
                             ByVal PreSettlementPrice As Double, ByVal AveragePrice As Double, _
                             ByVal UpdateTime As String, ByVal UpdateMilliSecond As Integer) Handles ctp.OnMarketData

        Core.DataArrive(InstrumentID, _
                              BidPrice1, BidVolume1, _
                              AskPrice1, AskVolume1, _
                              OpenPrice, HighestPrice, LowestPrice, _
                              LastPrice, OpenInterest, Volume, _
                              UpperLimitPrice, LowerLimitPrice, _
                              PreSettlementPrice, AveragePrice, _
                              UpdateTime, UpdateMilliSecond)
    End Sub

    Private Sub CTP_OnMDConnected() Handles ctp.OnMDConnected
        ConsoleControl.Log("行情数据上线", ConsoleControl.ConsoleLogLevel.Info)
    End Sub

    Private Sub CTP_OnMDDisconnected() Handles ctp.OnMDDisconnected
        ConsoleControl.Log("行情数据断线", ConsoleControl.ConsoleLogLevel.Info)
    End Sub

    Private Sub CTP_OnTradeConnected() Handles ctp.OnTradeConnected
        ConsoleControl.Log("交易系统上线", ConsoleControl.ConsoleLogLevel.Info)
    End Sub

    Private Sub CTP_OnTradeDisconnected() Handles ctp.OnTradeDisconnected
        ConsoleControl.Log("交易系统断线", ConsoleControl.ConsoleLogLevel.Info)
    End Sub

End Module
