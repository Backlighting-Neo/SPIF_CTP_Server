Public Class SPIF_Template
    Const RulesMax = 9999

    Public Structure LimitRange
        Dim IsConsideration As Boolean
        Dim Max As Double
        Dim Min As Double
    End Structure

    Public Structure Rules
        Dim JJC() As LimitRange    '3为合计
        Dim MMC() As LimitRange
        Dim KPC() As LimitRange
    End Structure

    Public Structure Template
        Dim ID As Integer
        Dim Name As String
        Dim Enable As Boolean
        Dim Command As Int16
        Dim Direction As Int16
        Dim Rules As Rules
    End Structure

    Public AllTemplate() As Template
    Public IsInit As Boolean = False

    Public Function InitTemplate(ByVal InitString As String, ByRef ErrorString As String)
        Dim TemplateCode() As String  '所有模板代码的数组
        Dim TemplateNumber As Integer
        Dim TemplateRules() As String
        Dim TemplateConfig() As String
        Dim RulesConfig() As String

        InitString = InitString.Trim()
        InitString = InitString.Replace(vbCrLf + vbCrLf, ";")
        TemplateCode = InitString.Split(";")
        TemplateNumber = TemplateCode.Length

        If TemplateNumber = 0 Then
            ErrorString = "没有找到模板代码"
            Return -1
        End If

        ReDim AllTemplate(TemplateNumber - 1)

        For Counter = 0 To TemplateNumber - 1
            TemplateRules = TemplateCode(Counter).Split(vbCrLf)
            If TemplateRules.Length < 2 Then
                ErrorString = "第" + CStr(Counter + 1) + "个模板代码有误"
                Return -1
            End If

            TemplateConfig = TemplateRules(0).Split(" ")
            If TemplateConfig.Length <> 4 Then
                ErrorString = "第" + CStr(Counter + 1) + "个模板代码的常规配置部分有误"
                Return -1
            End If

            Try
                AllTemplate(Counter).ID = CInt(TemplateConfig(0))
            Catch
                ErrorString = "第" + CStr(Counter + 1) + "个模板代码的常规配置的ID号为非数字"
                Return -1
            End Try

            AllTemplate(Counter).Name = TemplateConfig(1)

            Select Case TemplateConfig(2)
                Case "开启"
                    AllTemplate(Counter).Enable = True
                Case "关闭"
                    AllTemplate(Counter).Enable = False
                Case Else
                    ErrorString = "第" + CStr(Counter + 1) + "个模板代码的常规配置的使能字段只能使用<开启>或<关闭>"
                    Return -1
            End Select

            Select Case TemplateConfig(3)
                Case "开多"
                    AllTemplate(Counter).Command = 1
                    AllTemplate(Counter).Direction = 1
                Case "开空"
                    AllTemplate(Counter).Command = 1
                    AllTemplate(Counter).Direction = 0
                Case "平多"
                    AllTemplate(Counter).Command = 0
                    AllTemplate(Counter).Direction = 1
                Case "平空"
                    AllTemplate(Counter).Command = 0
                    AllTemplate(Counter).Direction = 0
                Case Else
                    ErrorString = "第" + CStr(Counter + 1) + "个模板代码的常规配置的操作字段只能使用<开多>,<开空>,<平多>,<平空>"
                    Return -1
            End Select

            ReDim AllTemplate(Counter).Rules.JJC(3)
            ReDim AllTemplate(Counter).Rules.KPC(3)
            ReDim AllTemplate(Counter).Rules.MMC(3)

            For RulesCounter = 1 To TemplateRules.Length - 1
                
                RulesConfig = TemplateRules(RulesCounter).Split(" ")
                RulesConfig(0) = RulesConfig(0).Trim()

                If RulesConfig.Length <> 2 And RulesConfig.Length <> 3 Then
                    ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则代码有误"
                    Return -1
                End If

                If RulesConfig(0).Length <> 4 Then
                    ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则代码的校验项目名有误"
                    Return -1
                End If

                Dim ItemID As Integer
                Try
                    ItemID = CInt(RulesConfig(0).Substring(3, 1))
                Catch
                    ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则代码的校验项目名有误"
                    Return -1
                End Try

                Select Case RulesConfig(0).Substring(0, 3)
                    Case "JJC"
                        If AllTemplate(Counter).Rules.JJC(ItemID).IsConsideration Then
                            ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则存在重定义"
                            Return -1
                        End If
                        AllTemplate(Counter).Rules.JJC(ItemID).IsConsideration = True
                        If RulesConfig.Length = 2 Then '只有一种
                            Select Case RulesConfig(1).Substring(0, 1)
                                Case ">"
                                    Try
                                        AllTemplate(Counter).Rules.JJC(ItemID).Min = CDbl(RulesConfig(1).Substring(1, RulesConfig(1).Length - 1))
                                    Catch
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End Try
                                    AllTemplate(Counter).Rules.JJC(ItemID).Max = RulesMax
                                Case "<"
                                    Try
                                        AllTemplate(Counter).Rules.JJC(ItemID).Max = CDbl(RulesConfig(1).Substring(1, RulesConfig(1).Length - 1))
                                    Catch
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End Try
                                    AllTemplate(Counter).Rules.JJC(ItemID).Min = -RulesMax
                                Case Else
                                    ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                    Return -1
                            End Select

                        Else '大于小于都有
                            Select Case RulesConfig(1).Substring(0, 1)
                                Case ">"
                                    If RulesConfig(2).Substring(0, 1) <> "<" Then
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End If

                                    Try
                                        AllTemplate(Counter).Rules.JJC(ItemID).Min = CDbl(RulesConfig(1).Substring(1, RulesConfig(1).Length - 1))
                                        AllTemplate(Counter).Rules.JJC(ItemID).Max = CDbl(RulesConfig(2).Substring(1, RulesConfig(2).Length - 1))
                                    Catch
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End Try

                                Case "<"
                                    If RulesConfig(2).Substring(0, 1) <> ">" Then
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End If

                                    Try
                                        AllTemplate(Counter).Rules.JJC(ItemID).Max = CDbl(RulesConfig(1).Substring(1, RulesConfig(1).Length - 1))
                                        AllTemplate(Counter).Rules.JJC(ItemID).Min = CDbl(RulesConfig(2).Substring(1, RulesConfig(2).Length - 1))
                                    Catch
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End Try

                            End Select
                        End If  '大于小于都有，还是只有一个

                    Case "MMC"
                        If AllTemplate(Counter).Rules.MMC(ItemID).IsConsideration Then
                            ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则存在重定义"
                            Return -1
                        End If
                        AllTemplate(Counter).Rules.MMC(ItemID).IsConsideration = True
                        If RulesConfig.Length = 2 Then '只有一种
                            Select Case RulesConfig(1).Substring(0, 1)
                                Case ">"
                                    Try
                                        AllTemplate(Counter).Rules.MMC(ItemID).Min = CDbl(RulesConfig(1).Substring(1, RulesConfig(1).Length - 1))
                                    Catch
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End Try
                                    AllTemplate(Counter).Rules.MMC(ItemID).Max = RulesMax
                                Case "<"
                                    Try
                                        AllTemplate(Counter).Rules.MMC(ItemID).Max = CDbl(RulesConfig(1).Substring(1, RulesConfig(1).Length - 1))
                                    Catch
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End Try
                                    AllTemplate(Counter).Rules.MMC(ItemID).Min = -RulesMax
                                Case Else
                                    ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                    Return -1
                            End Select

                        Else '大于小于都有
                            Select Case RulesConfig(1).Substring(0, 1)
                                Case ">"
                                    If RulesConfig(2).Substring(0, 1) <> "<" Then
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End If

                                    Try
                                        AllTemplate(Counter).Rules.MMC(ItemID).Min = CDbl(RulesConfig(1).Substring(1, RulesConfig(1).Length - 1))
                                        AllTemplate(Counter).Rules.MMC(ItemID).Max = CDbl(RulesConfig(2).Substring(1, RulesConfig(2).Length - 1))
                                    Catch
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End Try

                                Case "<"
                                    If RulesConfig(2).Substring(0, 1) <> ">" Then
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End If

                                    Try
                                        AllTemplate(Counter).Rules.MMC(ItemID).Max = CDbl(RulesConfig(1).Substring(1, RulesConfig(1).Length - 1))
                                        AllTemplate(Counter).Rules.MMC(ItemID).Min = CDbl(RulesConfig(2).Substring(1, RulesConfig(2).Length - 1))
                                    Catch
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End Try

                            End Select
                        End If  '大于小于都有，还是只有一个

                    Case "KPC"
                        If AllTemplate(Counter).Rules.KPC(ItemID).IsConsideration Then
                            ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则存在重定义"
                            Return -1
                        End If
                        AllTemplate(Counter).Rules.KPC(ItemID).IsConsideration = True
                        If RulesConfig.Length = 2 Then '只有一种
                            Select Case RulesConfig(1).Substring(0, 1)
                                Case ">"
                                    Try
                                        AllTemplate(Counter).Rules.KPC(ItemID).Min = CDbl(RulesConfig(1).Substring(1, RulesConfig(1).Length - 1))
                                    Catch
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End Try
                                    AllTemplate(Counter).Rules.KPC(ItemID).Max = RulesMax
                                Case "<"
                                    Try
                                        AllTemplate(Counter).Rules.KPC(ItemID).Max = CDbl(RulesConfig(1).Substring(1, RulesConfig(1).Length - 1))
                                    Catch
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End Try
                                    AllTemplate(Counter).Rules.KPC(ItemID).Min = -RulesMax
                                Case Else
                                    ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                    Return -1
                            End Select

                        Else '大于小于都有
                            Select Case RulesConfig(1).Substring(0, 1)
                                Case ">"
                                    If RulesConfig(2).Substring(0, 1) <> "<" Then
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End If

                                    Try
                                        AllTemplate(Counter).Rules.KPC(ItemID).Min = CDbl(RulesConfig(1).Substring(1, RulesConfig(1).Length - 1))
                                        AllTemplate(Counter).Rules.KPC(ItemID).Max = CDbl(RulesConfig(2).Substring(1, RulesConfig(2).Length - 1))
                                    Catch
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End Try

                                Case "<"
                                    If RulesConfig(2).Substring(0, 1) <> ">" Then
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End If

                                    Try
                                        AllTemplate(Counter).Rules.KPC(ItemID).Max = CDbl(RulesConfig(1).Substring(1, RulesConfig(1).Length - 1))
                                        AllTemplate(Counter).Rules.KPC(ItemID).Min = CDbl(RulesConfig(2).Substring(1, RulesConfig(2).Length - 1))
                                    Catch
                                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则定义不明"
                                        Return -1
                                    End Try

                            End Select
                        End If  '大于小于都有，还是只有一个
                    Case Else
                        ErrorString = "第" + CStr(Counter + 1) + "个模板代码的第" + CStr(RulesCounter) + "个规则代码的校验项目名有误"
                        Return -1
                End Select
            Next
        Next
        Me.IsInit = True
        ErrorString = "模板创建完成"
        Return 0
    End Function

    Public Function Testing(ByVal Data As SPIF_MetaData, ByVal ExpectionCommand As Integer, ByVal ExpextionDirection As Integer, ByRef TemplateName As String)
        If Not Me.IsInit Then
            TemplateName = "No Init"
            Return -1
        End If

        Dim RulesCheckFlag
        For Counter = 0 To AllTemplate.Length - 1

            If Not AllTemplate(Counter).Enable Then  '本模版未使能
                Continue For
            End If
            RulesCheckFlag = True
            TemplateName = AllTemplate(Counter).Name

            '当前Rules  AllTemplate(Counter).Rules(RulesCounter)
            '循环检查JJC
            For RulesCounter = 0 To 3
                If Not AllTemplate(Counter).Rules.JJC(RulesCounter).IsConsideration Then
                    Continue For
                End If
                If Data.MetaData(RulesCounter).JJC < AllTemplate(Counter).Rules.JJC(RulesCounter).Min Or _
                    Data.MetaData(RulesCounter).JJC > AllTemplate(Counter).Rules.JJC(RulesCounter).Max Then
                    RulesCheckFlag = False
                    Exit For
                Else
                    Continue For
                End If
            Next
            If Not RulesCheckFlag Then
                Continue For
            End If
            '循环检查MMC
            For RulesCounter = 0 To 3
                If Not AllTemplate(Counter).Rules.MMC(RulesCounter).IsConsideration Then
                    Continue For
                End If
                If Data.MetaData(RulesCounter).MMC < AllTemplate(Counter).Rules.MMC(RulesCounter).Min Or _
                    Data.MetaData(RulesCounter).MMC > AllTemplate(Counter).Rules.MMC(RulesCounter).Max Then
                    RulesCheckFlag = False
                    Exit For
                Else
                    Continue For
                End If
            Next
            If Not RulesCheckFlag Then
                Continue For
            End If
            '循环检查KPC
            For RulesCounter = 0 To 3
                If Not AllTemplate(Counter).Rules.KPC(RulesCounter).IsConsideration Then
                    Continue For
                End If
                If Data.MetaData(RulesCounter).KPC < AllTemplate(Counter).Rules.KPC(RulesCounter).Min Or _
                    Data.MetaData(RulesCounter).KPC > AllTemplate(Counter).Rules.KPC(RulesCounter).Max Then
                    RulesCheckFlag = False
                    Exit For
                Else
                    Continue For
                End If
            Next
            If Not RulesCheckFlag Then
                Continue For
            End If
            '检查过程中，一个不符合，则立即exitFor，并置RulesCheckFlag = False，在外层exit模板for

            If (ExpectionCommand = -1 Or AllTemplate(Counter).Command = ExpectionCommand) _
            And (ExpextionDirection = -1 Or AllTemplate(Counter).Direction = ExpextionDirection) Then
                Return Counter
            End If

        Next

        TemplateName = "No Match"
        Return -1
    End Function

End Class
