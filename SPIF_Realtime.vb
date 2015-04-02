Class Realtime
    Public Structure CTPBufferMetaStructure
        Dim Updatetime As String
        Dim UpdatemilliSecond As Integer
        Dim LastPrice As Double
        Dim OpenInterest As Integer
        Dim BidVolume1 As Integer
        Dim AskVolume1 As Integer
    End Structure

    Structure CTPBufferStructure
        Dim Data() As CTPBufferMetaStructure
        Dim DataCount As Integer
        Dim AverageVolume As Double
        Dim AverageBidAskVolume As Double
        Dim AverageLastPrice As Double
    End Structure

    Structure OutputDataStructure
        Dim JJC As Double
        Dim MMC As Double
        Dim KPC As Double
    End Structure

    Dim CTPBuffer(5400) As CTPBufferStructure
    Dim OutputData(5400) As OutputDataStructure
    Dim CurrentCursor = 0
    Dim LastOpenInterest As Double = 0

    Sub New()
        For Counter = 0 To 5400
            ReDim CTPBuffer(Counter).Data(6)
            CTPBuffer(Counter).DataCount = 0
        Next
    End Sub

    Function GetMetaData(Counter As Integer) As SPIF_MetaData
        Dim TempMetaData As SPIF_MetaData = New SPIF_MetaData
        TempMetaData.SetData(Me.OutputData(Counter - 2).JJC, Me.OutputData(Counter - 1).JJC, Me.OutputData(Counter).JJC, _
                Me.OutputData(Counter - 2).MMC, Me.OutputData(Counter - 1).MMC, Me.OutputData(Counter).MMC, _
                Me.OutputData(Counter - 2).KPC, Me.OutputData(Counter - 1).KPC, Me.OutputData(Counter).KPC)

        Return TempMetaData
    End Function

    Function AddCTPMarketData(MarketData As CTPBufferMetaStructure) As Integer
        Dim ID = Support.Time2ID(MarketData.Updatetime)
        Dim DataCount As Integer
        Dim TotalVolume, TotalBidAskVolume As Integer
        Dim TotalLastPrice, temp As Double
        Dim ReturnValue As Integer = 0

        If ID <> CurrentCursor Then '发生跳变
            '对上一条数据进行结算
            ReturnValue = CurrentCursor
            DataCount = CTPBuffer(CurrentCursor).DataCount
            TotalVolume = 0
            TotalBidAskVolume = 0
            TotalLastPrice = 0
            For Counter = 0 To DataCount - 1
                TotalVolume += CTPBuffer(CurrentCursor).Data(Counter).OpenInterest
                TotalBidAskVolume += CTPBuffer(CurrentCursor).Data(Counter).BidVolume1 - CTPBuffer(CurrentCursor).Data(Counter).AskVolume1
                TotalLastPrice += CTPBuffer(CurrentCursor).Data(Counter).LastPrice
            Next
            CTPBuffer(CurrentCursor).AverageVolume = TotalVolume / DataCount
            CTPBuffer(CurrentCursor).AverageLastPrice = TotalLastPrice / DataCount
            CTPBuffer(CurrentCursor).AverageBidAskVolume = TotalBidAskVolume / DataCount

            If CurrentCursor > 3 Then
                '结算Output
                If CTPBuffer(CurrentCursor - 1).AverageLastPrice = 0 Then
                    OutputData(CurrentCursor).JJC = 0
                Else
                    OutputData(CurrentCursor).JJC = CTPBuffer(CurrentCursor).AverageLastPrice - CTPBuffer(CurrentCursor - 1).AverageLastPrice
                End If
                OutputData(CurrentCursor).MMC = CTPBuffer(CurrentCursor).AverageBidAskVolume + _
                                                                        CTPBuffer(CurrentCursor - 1).AverageBidAskVolume + _
                                                                        CTPBuffer(CurrentCursor - 2).AverageBidAskVolume
                OutputData(CurrentCursor).KPC = CTPBuffer(CurrentCursor).AverageVolume + _
                                                                     CTPBuffer(CurrentCursor - 1).AverageVolume + _
                                                                     CTPBuffer(CurrentCursor - 2).AverageVolume

                TemplateSourceData += CurrentCursor.ToString + "," + Support.ID2Time(CurrentCursor) + "," + Format(CTPBuffer(CurrentCursor).Data(CTPBuffer(CurrentCursor).DataCount - 1).LastPrice, "0.0") + "," + _
                                  Format(OutputData(CurrentCursor).JJC, "0.00") + "," + Format(OutputData(CurrentCursor).MMC, "0.00") + "," + Format(OutputData(CurrentCursor).KPC, "0.00") + vbCrLf
            End If
            CurrentCursor = ID
        End If

        temp = MarketData.OpenInterest
        If LastOpenInterest <> 0 Then
            MarketData.OpenInterest -= LastOpenInterest
        Else
            MarketData.OpenInterest = 0
        End If
        LastOpenInterest = temp

        DataCount = CTPBuffer(CurrentCursor).DataCount
        CTPBuffer(CurrentCursor).Data(DataCount) = MarketData
        CTPBuffer(CurrentCursor).DataCount += 1
        '进本条Buffer.Data数据

        Return ReturnValue
    End Function

    Function GetLastPriceById(Id As Integer) As Double
        Return CTPBuffer(Id).Data(CTPBuffer(Id).DataCount - 1).LastPrice
    End Function

    Function OutputCustomData() As String
        Dim ReturnValue As String = String.Empty

        ReturnValue = "ID,Time,JN,JJC,MMC,KPC" + vbCrLf
        ReturnValue += "1,09:15:03,0,0,0,0" + vbCrLf
        ReturnValue += "2,09:15:06,0,0,0,0" + vbCrLf
        ReturnValue += "3,09:15:09,0,0,0,0" + vbCrLf

        For Counter = 4 To CurrentCursor
            ReturnValue += Counter.ToString + "," + Support.ID2Time(Counter) + "," + _
                Format(GetLastPriceById(Counter), "0.0") + "," + _
                Format(OutputData(Counter).JJC, "0.00") + "," + _
                Format(OutputData(Counter).MMC, "0.00") + "," + _
                Format(OutputData(Counter).KPC, "0.00") + vbCrLf
        Next

        Return ReturnValue
    End Function
End Class