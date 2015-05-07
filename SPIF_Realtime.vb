Imports Newtonsoft.Json

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

    Structure BroadcastingData
        Dim Id As Integer
        Dim Time As String
        Dim LastPrice As String
        Dim JJC As String
        Dim MMC As String
        Dim KPC As String
        Dim JJC3 As String
        Dim MMC3 As String
        Dim KPC3 As String
        Dim BBD As String
        Dim Bull As String
        Dim Bear As String
        Dim AccBBD As String
    End Structure

    Dim CTPBuffer(5400) As CTPBufferStructure
    Dim OutputData(5400) As OutputDataStructure
    Dim CurrentCursor = 0
    Dim LastOpenInterest As Double = 0
    Dim AccumlationBBD As Double

    Public ws As Websocket

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
                    'OutputData(CurrentCursor).JJC = CTPBuffer(CurrentCursor).AverageLastPrice - CTPBuffer(CurrentCursor - 1).AverageLastPrice
                    OutputData(CurrentCursor).JJC = GetLastPriceById(CurrentCursor) - GetLastPriceById(CurrentCursor - 1)
                End If
                OutputData(CurrentCursor).MMC = CTPBuffer(CurrentCursor).AverageBidAskVolume + _
                                                                        CTPBuffer(CurrentCursor - 1).AverageBidAskVolume + _
                                                                        CTPBuffer(CurrentCursor - 2).AverageBidAskVolume
                OutputData(CurrentCursor).KPC = CTPBuffer(CurrentCursor).AverageVolume + _
                                                                     CTPBuffer(CurrentCursor - 1).AverageVolume + _
                                                                     CTPBuffer(CurrentCursor - 2).AverageVolume

                Dim OutputDataToString As String = String.Empty
                Dim Broadcasting As BroadcastingData
                Broadcasting.Id = CurrentCursor
                Broadcasting.Time = Support.ID2Time(CurrentCursor)
                Broadcasting.LastPrice = Support.GetOutputNumber1((GetLastPriceById(CurrentCursor)))
                Broadcasting.JJC = Support.GetOutputNumber1(OutputData(CurrentCursor).JJC)
                Broadcasting.MMC = Support.GetOutputNumber0(OutputData(CurrentCursor).MMC)
                Broadcasting.KPC = Support.GetOutputNumber0(OutputData(CurrentCursor).KPC)

                Dim TempMetaData As SPIF_MetaData
                TempMetaData = GetMetaData(CurrentCursor)

                Dim JJC3 As Double = 0
                Dim MMC3 As Double = 0
                Dim KPC3 As Double = 0
                Dim Bull As Double = 0
                Dim Bear As Double = 0
                JJC3 = TempMetaData.MetaData(3).JJC
                MMC3 = TempMetaData.MetaData(3).MMC
                KPC3 = TempMetaData.MetaData(3).KPC

                Broadcasting.JJC3 = Support.GetOutputNumber1(JJC3)
                Broadcasting.KPC3 = Support.GetOutputNumber0(KPC3)
                Broadcasting.MMC3 = Support.GetOutputNumber0(MMC3)

                Dim TempBBD As Double
                TempBBD = TempMetaData.MetaData(3).KPC * TempMetaData.MetaData(3).JJC

                If JJC3 > 0 Then
                    If KPC3 > 0 Then
                        Bull = TempBBD
                    Else
                        TempBBD = -TempBBD
                        Bear = TempBBD
                    End If
                Else
                    If KPC3 > 0 Then
                        Bear = TempBBD
                    Else
                        TempBBD = -TempBBD
                        Bull = TempBBD
                    End If
                End If

                Broadcasting.Bull = Support.GetOutputNumber0(Bull)
                Broadcasting.Bear = Support.GetOutputNumber0(Bear)

                AccumlationBBD += TempBBD
                Broadcasting.BBD = Support.GetOutputNumber0(TempBBD)
                Broadcasting.AccBBD = Support.GetOutputNumber0(AccumlationBBD)

                OutputDataToString = Broadcasting.Id.ToString + "," + Broadcasting.Time + "," + Broadcasting.LastPrice + "," + _
                                  Broadcasting.JJC + "," + Broadcasting.MMC + "," + Broadcasting.KPC

                TemplateSourceData += OutputDataToString + vbCrLf

                ws.Broadcasting(JsonConvert.SerializeObject(Broadcasting))
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
        If CTPBuffer(Id).DataCount = 0 Then
            Return 0
        Else
            Return CTPBuffer(Id).Data(CTPBuffer(Id).DataCount - 1).LastPrice
        End If
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