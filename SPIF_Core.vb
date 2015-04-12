Public Class SPIF_Core
    Private Realtime As Realtime = New Realtime
    Private Template As SPIF_Template = New SPIF_Template

    Public Property ID_Start_AM As Integer = 0
    Public Property ID_End_AM As Integer = 0
    Public Property ID_Start_PM As Integer = 0
    Public Property ID_End_PM As Integer = 0

    Public Property Point_StopLoss As Double = 0
    Public Property ID_ForceOffset As Integer = 0
    Public Property ID_OrderGap As Integer = 0

    Public Property Point_Offset_BullOpen As Double = 0
    Public Property Point_Offset_BullOffset As Double = 0
    Public Property Point_Offset_BearOpen As Double = 0
    Public Property Point_Offset_BearOffset As Double = 0

    Public Property OrderVolume As Integer = 0

    Public Property ID_CancelOrder_Timeout As Integer = 0

    Public Property Instrument As String = String.Empty

    Public Delegate Sub PushOrderDelegate(Command As Integer, Direction As Integer, _
                                 Price As Double, Volume As Integer, Reason As String, OrderFakeID As Integer)
    Public Delegate Sub CancelOrderDelegate(OrderID As Long)
    Public PushOrder As PushOrderDelegate
    Public CancelOrder As CancelOrderDelegate

    Private LastPrice As Double
    Private LastUpdateTime As String
    Private LastUpdateMilliSecond As Integer


    Private ChangingTemplate As Boolean = False
    Private ChangingLockingDirection As Boolean = False
    Private IsInitTemplate As Boolean = False

    Public CurrentID As Integer = 0
    Private CurrentExpectionCommand As Integer = 1  '期待操作：1=开，0-平
    Private CurrentExpectionDirection As Integer = -1   '期待方向：1=多，0=空，-1=不限
    Private LockingDirection As Integer = -1                 '锁定的方向
    Private CurrentOrderOpenPrice As Double = 0  '当前单开仓均价
    Private CurrentOrderBenefit As Double = 0  '当前单收益点数
    Private CurrentOrderMaxBenefit As Double = -9999 '当前单最大收益点数
    Private TotalOrderBenefit As Double = 0
    Private FallBackThreshold As Double = 0 '最大回撤阈值
    Private FallBackPercent As Double = 0 '最大回撤比例

    Private Locking As Boolean = False  '目前是否正在报单
    Private LockingID As Integer = 0      '报单时刻ID，用于超时撤单
    Private LockingOrderID As Long = 0  '目前报单未决的报单编号
    Private LockingCanceled As Boolean = False  '目前报单未决的报单是否已经尝试撤单
    Private LastOpenOrderID As Integer = 0  '上一个已开单的ID号

    Private Pausing As Boolean = False

    Private OriginalCTPData As String = String.Empty
    Private OriginalTradeResult As String = String.Empty

    Private OrderCounter As Integer = 0

    Public Config_Default_Output_CTP, Config_Default_Output_Custom, Config_Default_Output_Trade As String  '默认输出位置

    Private Structure Order
        Dim OrderID As Long
        Dim InsertTime As String
        Dim LastUpdatetime As DateTime
        Dim Instrument As String
        Dim RestVolume As Integer
        Dim InsertPrice As Double
        Dim TradedPrice As Double
        Dim IsFinished As Boolean
    End Structure


    Public Sub New()
        OriginalCTPData = "合约编号,买一价,买一量,卖一价,卖一量,开盘价,最高价,最低价,最新价,持仓量,成交量,涨停价,跌停价,昨结算价,今日平均价,行情更新时间,更新毫秒数" + vbCrLf
        Config_Default_Output_CTP = Support.GetSettings("CTPOutput")
        Config_Default_Output_Custom = Support.GetSettings("TemplateOutput")
        Config_Default_Output_Trade = Support.GetSettings("ResultOutput")

        FallBackThreshold = Support.GetSettings("FallBackThreshold")
        FallBackPercent = Support.GetSettings("FallBackPercent") / 100
    End Sub

    Public Sub SetWebsocket(ws As Websocket)
        Realtime.ws = ws
    End Sub

    Public Function GetCurrentOrderInfo()
        Return "第" + Format(OrderCounter, "00") + "单  本单盈利" + _
            Format(CurrentOrderBenefit, "0.0") + "点，累计盈利" + _
            Format(TotalOrderBenefit, "0.0") + "点"
    End Function

    Private Sub BullOpen(Reason As String)  '多开
        Me.CurrentOrderOpenPrice = Me.LastPrice
        Me.OriginalTradeResult += Support.ID2Time(CurrentID) + "  【" + CurrentID.ToString + "】    " + Reason + "   " + Me.LastPrice.ToString + vbCrLf
        PushOrder(1, 1, Me.LastPrice + Point_Offset_BullOpen, OrderVolume, Reason, OrderCounter)
    End Sub

    Private Sub BullOffset(Reason As String)  '多平
        OrderCounter += 1
        Me.LastOpenOrderID = Me.CurrentID
        Me.OriginalTradeResult += Support.ID2Time(CurrentID) + "  【" + CurrentID.ToString + "】    " + Reason + "   " + Me.LastPrice.ToString + vbCrLf
        TotalOrderBenefit += CurrentOrderBenefit
        Me.OriginalTradeResult += "第" + OrderCounter.ToString + "单收益" + Format(CurrentOrderBenefit, "0.0") + "点  累计收益" + Format(TotalOrderBenefit, "0.0") + "点" + vbCrLf + vbCrLf
        PushOrder(0, 1, Me.LastPrice + Point_Offset_BullOffset, OrderVolume, Reason, OrderCounter)
        Me.CurrentExpectionCommand = 1
        Me.CurrentOrderMaxBenefit = -9999
        RestoreExpectionDirection()
    End Sub

    Private Sub BearOpen(Reason As String)  '空开
        Me.CurrentOrderOpenPrice = Me.LastPrice
        Me.OriginalTradeResult += Support.ID2Time(CurrentID) + "  【" + CurrentID.ToString + "】    " + Reason + "   " + Me.LastPrice.ToString + vbCrLf
        PushOrder(1, 0, Me.LastPrice + Point_Offset_BearOpen, OrderVolume, Reason, OrderCounter)
    End Sub

    Private Sub BearOffset(Reason As String)  '空平
        OrderCounter += 1
        Me.LastOpenOrderID = Me.CurrentID
        Me.OriginalTradeResult += Support.ID2Time(CurrentID) + "  【" + CurrentID.ToString + "】    " + Reason + "   " + Me.LastPrice.ToString + vbCrLf
        TotalOrderBenefit += CurrentOrderBenefit
        Me.OriginalTradeResult += "第" + OrderCounter.ToString + "单收益" + Format(CurrentOrderBenefit, "0.0") + "点  累计收益" + Format(TotalOrderBenefit, "0.0") + "点" + vbCrLf + vbCrLf
        PushOrder(0, 0, Me.LastPrice + Point_Offset_BearOffset, OrderVolume, Reason, OrderCounter)
        Me.CurrentExpectionCommand = 1
        Me.CurrentOrderMaxBenefit = -9999
        RestoreExpectionDirection()
    End Sub

    Private Sub Cancel(OrderID As Long)
        CancelOrder(OrderID)
    End Sub

    Private Sub RestoreExpectionDirection()
        If Me.ChangingLockingDirection Then   '若方向锁定则
            Me.CurrentExpectionDirection = Me.LockingDirection
        Else
            Me.CurrentExpectionDirection = -1
        End If
    End Sub

    Private Function IsNotOpenable() As Boolean
        If CurrentID - LastOpenOrderID < ID_OrderGap Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub DataArrive(ByVal InstrumentID As String, _
                             ByVal BidPrice1 As Double, ByVal BidVolume1 As Integer, _
                             ByVal AskPrice1 As Double, ByVal AskVolume1 As Integer, _
                             ByVal OpenPrice As Double, ByVal HighestPrice As Double, ByVal LowestPrice As Double, _
                             ByVal LastPrice As Double, ByVal OpenInterest As Integer, ByVal Volume As Integer, _
                             ByVal UpperLimitPrice As Double, ByVal LowerLimitPrice As Double, _
                             ByVal PreSettlementPrice As Double, ByVal AveragePrice As Double, _
                             ByVal UpdateTime As String, ByVal UpdateMilliSecond As Integer)

        OriginalCTPData += InstrumentID.ToString + "," + BidPrice1.ToString + "," + BidVolume1.ToString + "," + _
        AskPrice1.ToString + "," + AskVolume1.ToString + "," + OpenPrice.ToString + "," + HighestPrice.ToString + "," + _
        LowestPrice.ToString + "," + LastPrice.ToString + "," + OpenInterest.ToString + "," + Volume.ToString + "," + _
        UpperLimitPrice.ToString + "," + LowerLimitPrice.ToString + "," + PreSettlementPrice.ToString + "," + _
        AveragePrice.ToString + "," + UpdateTime.ToString + "," + UpdateMilliSecond.ToString + vbCrLf

        '数据到达
        If Support.Time2ID(UpdateTime) = -1 Then
            Exit Sub
        End If

        Dim ArriviedData As Realtime.CTPBufferMetaStructure
        Dim CustomData As SPIF_MetaData  '当前跳变的模板数据

        ArriviedData.AskVolume1 = AskVolume1
        ArriviedData.BidVolume1 = BidVolume1
        ArriviedData.LastPrice = LastPrice
        ArriviedData.UpdatemilliSecond = UpdateMilliSecond
        ArriviedData.Updatetime = UpdateTime
        ArriviedData.OpenInterest = OpenInterest

        CurrentID = Realtime.AddCTPMarketData(ArriviedData)  '数据进树

        Me.LastPrice = LastPrice

        If CurrentID = 0 Then  '发生未发生数据跳变则放弃本条数据
            Exit Sub
        End If

        If ChangingTemplate Or Pausing Or Locking Or IsNotOpenable() Then
            Exit Sub  '若 ( 模板正在被调整, 核心暂停, 报单未决 ) 则放弃本条数据，只进树
        End If

        If CurrentExpectionCommand = 0 Then
            CurrentOrderBenefit = LastPrice - CurrentOrderOpenPrice

            If CurrentExpectionDirection = 0 Then
                CurrentOrderBenefit = -CurrentOrderBenefit
            End If  '计算当前单收益

            If CurrentOrderBenefit < -Point_StopLoss Then
                If CurrentExpectionDirection = 1 Then   '多平止损
                    Call BullOffset("多平止损")
                Else  '空平止损
                    Call BearOffset("空平止损")
                End If
                Exit Sub
            End If  '止损

            CurrentOrderMaxBenefit = IIf(CurrentOrderBenefit > CurrentOrderMaxBenefit, CurrentOrderBenefit, CurrentOrderMaxBenefit)

            If CurrentOrderMaxBenefit > FallBackThreshold AndAlso _
               CurrentOrderBenefit < FallBackPercent * CurrentOrderMaxBenefit Then
                If CurrentExpectionDirection = 1 Then   '多平止损
                    Call BullOffset("多平止盈")
                Else  '空平止损
                    Call BearOffset("空平止盈")
                End If
                Exit Sub
            End If

        End If   '目前已开单

        If Not ((CurrentID > ID_Start_AM And CurrentID < ID_End_AM) Or (CurrentID > ID_Start_PM And CurrentID < ID_End_PM)) Then
            Exit Sub  '若没在开单区间则启用
        End If

        If CurrentID > ID_ForceOffset And CurrentExpectionCommand = 0 Then
            If CurrentExpectionDirection = 1 Then   '强制多平
                Call BullOffset("强制多平")
            Else  '强制空平
                Call BearOffset("强制空平")
            End If
            Exit Sub
        End If

        'If CurrentID - LockingID > ID_CancelOrder_Timeout Then
        '    LockingCanceled = True
        '    Call Cancel(LockingOrderID)  '超时撤单
        'End If

        CustomData = Realtime.GetMetaData(CurrentID)    '取当前

        Dim TemplateName As String = String.Empty
        Dim FindedTemplateNumber As Integer = 0
        FindedTemplateNumber = Template.Testing(CustomData, CurrentExpectionCommand, CurrentExpectionDirection, TemplateName)

        If FindedTemplateNumber = -1 Then  '没找到对应模板则弃用
            Exit Sub
        End If

        If CurrentExpectionCommand = 1 Then  '待开
            'TODO 维护订单信息
            CurrentExpectionDirection = Template.AllTemplate(FindedTemplateNumber).Direction  '将待平方向标记为模板对应方向

            Select Case CurrentExpectionDirection
                Case 0  '待开空
                    BearOpen(TemplateName)
                Case 1  '待开多
                    BullOpen(TemplateName)
            End Select

        Else  '待平
            'TODO 维护订单信息

            Select Case CurrentExpectionDirection
                Case 0  '待开空
                    BearOffset(TemplateName)
                Case 1  '待开多
                    BullOffset(TemplateName)
            End Select
        End If

        CurrentExpectionCommand = 1 - Template.AllTemplate(FindedTemplateNumber).Command
    End Sub

    Public Function Pause() As Integer
        If Me.Pausing Then
            Return -1
        Else
            Me.Pausing = True
            Return 0
        End If
    End Function

    Public Function CancelPause() As Integer
        If Me.Pausing Then
            Me.Pausing = False
            Return 0
        Else
            Return -1
        End If
    End Function

    Public Function ChangeLockingDirection(Direction As Integer) As Integer
        If CurrentExpectionCommand = 1 Then '如果目前处于待开状态
            ChangingLockingDirection = True
            LockingDirection = Direction
            ChangingLockingDirection = False
            Return 0
        Else
            Return -1
        End If
    End Function

    Public Function ChangeTemplate(TemplateFilePath As String) As Integer
        ChangingTemplate = True
        IsInitTemplate = True

        Dim ReturnValue As String = 0
        Dim TemplateContent As String = String.Empty

        Template = New SPIF_Template

        TemplateContent = Support.MyTxtReader(TemplateFilePath)
        If TemplateContent <> "" Then
            Dim errmsg As String = String.Empty
            If Template.InitTemplate(TemplateContent, errmsg) = -1 Then
                ReturnValue = -1
            End If
        Else
            ReturnValue = -1
        End If

        ChangingTemplate = False
        Return ReturnValue
    End Function

    Public Function OutputCTP(Optional FilePath As String = "Default") As Integer
        If FilePath = "Default" Then
            FilePath = Config_Default_Output_CTP
        End If
        ConsoleControl.Log("文件写入路径：" + FilePath, ConsoleControl.ConsoleLogLevel.Debug)
        If Support.WriteTextToFile(FilePath, OriginalCTPData) = -1 Then
            ConsoleControl.Log("CTP数据写入失败，请检查是否文件正在打开", ConsoleControl.ConsoleLogLevel.Debug)
            Return -1
        Else
            ConsoleControl.Log("CTP数据写入成功", ConsoleControl.ConsoleLogLevel.Debug)
            Return 0
        End If
    End Function

    Public Function OutputCustom(Optional FilePath As String = "Default") As Integer
        If FilePath = "Default" Then
            FilePath = Config_Default_Output_Custom
        End If
        Dim OriginalCustomData As String = String.Empty
        ConsoleControl.Log("正在生成数据……", ConsoleControl.ConsoleLogLevel.Debug)
        OriginalCustomData = Realtime.OutputCustomData()
        ConsoleControl.Log("文件写入路径：" + FilePath, ConsoleControl.ConsoleLogLevel.Debug)
        If Support.WriteTextToFile(FilePath, OriginalCustomData) = -1 Then
            ConsoleControl.Log("Custom数据写入失败，请检查是否文件正在打开", ConsoleControl.ConsoleLogLevel.Debug)
            Return -1
        Else
            ConsoleControl.Log("Custom数据写入成功", ConsoleControl.ConsoleLogLevel.Debug)
            Return 0
        End If
    End Function

    Public Function OutputTrade(Optional FilePath As String = "Default") As Integer
        If FilePath = "Default" Then
            FilePath = Config_Default_Output_Trade
        End If
        Dim TempTradeConect As String = String.Empty
        TempTradeConect = "CTP报告" + vbCrLf + "日期：" + vbCrLf + _
            "策略：" + Support.GetSettings("TemplateFilePath") + vbCrLf + _
            "回测时间：" + Now() + vbCrLf
        TempTradeConect += "全天" + OrderCounter.ToString + "单  全天收益" + Format(TotalOrderBenefit, "0.0") + "点" + vbCrLf + vbCrLf + OriginalTradeResult
        ConsoleControl.Log("文件写入路径：" + FilePath, ConsoleControl.ConsoleLogLevel.Debug)
        If Support.WriteTextToFile(FilePath, TempTradeConect) = -1 Then
            ConsoleControl.Log("Trade数据写入失败，请检查是否文件正在打开", ConsoleControl.ConsoleLogLevel.Debug)
            Return -1
        Else
            ConsoleControl.Log("Trade数据写入成功", ConsoleControl.ConsoleLogLevel.Debug)
            Return 0
        End If
    End Function
End Class
