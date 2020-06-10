Sub 全景表()
    Application.ScreenUpdating = False    '不显示
'    On Error GoTo ErrMsg
    Dim Ver, Errxx, StarTime
    Ver = "全景表自动生成工具 " & VerNo
    StarTime = Timer
    UserForm1.Caption = Ver
    Errxx = "未知错误！"

    Dim Temp, Temp1, A_temp, Gdsj, A_gdsj
    Dim nbH
    Dim Code, Name, Price, Zsz, xPE, xPB, MGsy, MGjzc, GSdy, GShy, GSname, GSzy, Sssj, CBsoure
    '代码.名称.股价.市值.PE.PB.每股收益.每股净资产.地域.行业.全名.主营.上市日期,财报来源
    Dim A_nb, A_qjbsc

    Temp = 0
    For x = 1 To Sheets.Count
        If Sheets(x).Name = "全景表工具" Then Temp = Temp + 1
        If Sheets(x).Name = "数据收集工具" Then Temp = Temp + 1
        If Sheets(x).Name = "本福特测试" Then Temp = Temp + 1
        If Sheets(x).Name = "自选股" Then Temp = Temp + 1
    Next
    If Temp <> 4 Then
        Errxx = "程序错误！"
        GoTo ErrMsg
    End If

    Code = Sheets("全景表工具").Cells(2, 2)
    nbH = Sheets("全景表工具").Cells(5, 2)
    CBsoure = Sheets("全景表工具").Cells(7, 2)

    '**********   检查网络   **********
    Temp = CreateObject("Wscript.shell").Run("ping qt.gtimg.cn -n 1", 0, True)
    If Temp <> 0 Then
        Errxx = "请检查网络是否通畅！"
        GoTo ErrMsg
    End If
    '**********   检查网络END   **********
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", "http://qt.gtimg.cn/q=sh000300", False
        .send
        sp = Split(.responsetext, "~")
        If UBound(sp) > 3 Then
            zxjyr = Format(Left(sp(30), 8), "0000-00-00")
        End If
    End With

    '**********   从腾讯获取名称股价市值   **********
    If Left$(Code, 1) <> "6" Then
        URL = "http://qt.gtimg.cn/q=sz" & Code
    Else
        URL = "http://qt.gtimg.cn/q=sh" & Code
    End If
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", URL, False
        .send
        Temp = Split(.responsetext, "~")
        If UBound(Temp) > 3 Then
            Name = Temp(1)
            Price = Temp(3)
            Zsz = Temp(45)
        Else
            Errxx = "网络不畅或股票代码输入错误！"
            GoTo ErrMsg
        End If
    End With
    '**********   从腾讯获取名称股价市值END   **********

    '**********   从网易获取历史行情（市值复权股价）   **********
    If Left(Code, 1) = 6 Then
        Temp = 0
    Else
        Temp = 1
    End If
    '历史市值
    Workbooks.Open "http://quotes.money.163.com/service/chddata.html?code=" & Temp & Code & "&start=19800101&end=20990101&fields=TCAP"
    h = ActiveSheet.UsedRange.Rows.Count
    A_temp = Range(Cells(2, 1), Cells(h, 4))
    Workbooks("chddata.html").Close savechanges:=False

    xx = UBound(A_temp)
    year0 = Left(A_temp(xx, 1), 4)
    year1 = Left(A_temp(1, 1), 4)
    Yy = year1 - year0 + 1
    Dim A_lshqb    '历史行情表
    ReDim A_lshqb(1 To Yy + 1, 1 To 4)    '年度.市值.复权价.涨幅

    For y = 1 To Yy
        For x = 1 To xx
            A_lshqb(y, 1) = year0 + y - 1
            If A_lshqb(y, 1) - Left(A_temp(xx - x + 1, 1), 4) = 0 Then
                If x = xx Then A_lshqb(y, 1) = A_temp(xx - x + 1, 1)
                A_lshqb(y, 2) = A_temp(xx - x + 1, 4) / 100000000
            End If
        Next
    Next

    '历史股价
    URL = "http://img1.money.126.net/data/hs/klinederc/day/times/" & Temp & Code & ".json"
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", URL, False
        .send
        Gdsj = .responsetext
    End With

    '    GoSub GosubLM

    Won = "["
    Woff = "]"
    GoSub GosubZL

    Gdsj = Split(Gdsj, "[")
    close1 = Split(Gdsj(1), ",")(0)

    Temp = Split(Gdsj(2), ",")
    Temp1 = Split(Gdsj(1), ",")
    xx = UBound(Temp)

    For y = 1 To Yy
        For x = 0 To xx
            If Left(A_lshqb(y, 1), 4) - Mid(Temp(x), 2, 4) = 0 Then
                A_lshqb(y, 3) = Temp1(x)
                If y = 1 Then
                    A_lshqb(y, 4) = A_lshqb(y, 3) / close1 - 1
                Else
                    If A_lshqb(y - 1, 3) = 0 Then A_lshqb(y - 1, 3) = close1
                    A_lshqb(y, 4) = A_lshqb(y, 3) / A_lshqb(y - 1, 3) - 1
                End If
            End If
        Next
        A_lshqb(y + 1, 3) = A_lshqb(y, 3)
        A_lshqb(y + 1, 4) = 0
    Next
    A_lshqb(Yy + 1, 1) = "上市累计"
    A_lshqb(Yy + 1, 3) = close1    'Temp1(xx)
    A_lshqb(Yy + 1, 4) = Temp1(xx) / close1 - 1
    '**********   从网易获取历史行情（市值复权股价）END   **********

    If IsNumeric(nbH) Then
        nbH = Round(nbH, 0) + 1
        If nbH < 3 Then nbH = 3
    Else
        nbH = 99
    End If

    If CBsoure = "网易" Then
        '**********   从网易获取财务报表   **********
        '资产负债表
        For x = 1 To Windows.Count
            If Windows(x).Caption = "zcfzb_" & Code & ".html" Then
                Workbooks("zcfzb_" & Code & ".html").Close savechanges:=False
                Exit For
            End If
        Next

        Workbooks.Open "http://quotes.money.163.com/service/zcfzb_" & Code & ".html"
        Temp = ActiveSheet.UsedRange.Rows.Count
        Temp1 = ActiveSheet.UsedRange.Columns.Count
        A_temp = WorksheetFunction.Transpose(Range(Cells(1, 1), Cells(Temp, Temp1)))
        Workbooks("zcfzb_" & Code & ".html").Close savechanges:=False
        Temp = 1
        For x = 2 To UBound(A_temp)
            If Month(A_temp(x, 1)) = 12 And Year(A_temp(x, 1)) > 1970 Then Temp = Temp + 1
        Next

        If nbH > Temp Then nbH = Temp
        If Month(A_temp(2, 1)) = 12 Then
            ReDim A_nb(1 To nbH, 1 To 7)
            xx = UBound(A_nb)
        Else
            ReDim A_nb(1 To nbH + 1, 1 To 7)
            A_nb(nbH + 1, 1) = A_temp(2, 1)
            xx = UBound(A_nb) - 1
        End If

        For x = 2 To UBound(A_temp)
            If Month(A_temp(x, 1)) = 12 Then
                A_nb(xx, 1) = A_temp(x, 1)
                xx = xx - 1
            End If
            If xx = 1 Then Exit For
        Next

        A_nb(1, 1) = "年度"
        A_nb(1, 2) = "股东权益"
        A_nb(1, 3) = "资产总计"
        A_nb(1, 4) = "营业收入"
        A_nb(1, 5) = "净利润"
        A_nb(1, 6) = "经营现金净流量"
        A_nb(1, 7) = "母公司股东权益"

        xx = UBound(A_nb)
        Yy = UBound(A_temp)
        zz = UBound(A_temp, 2)
        For y = 2 To Yy
            If A_nb(xx, 1) = A_temp(y, 1) Then
                For Z = 2 To zz
                    If A_temp(1, Z) = "所有者权益(或股东权益)合计(万元)" Then
                        A_nb(xx, 2) = A_temp(y, Z) / 10000
                    ElseIf A_temp(1, Z) = "资产总计(万元)" Then
                        A_nb(xx, 3) = A_temp(y, Z) / 10000
                    ElseIf A_temp(1, Z) = "归属于母公司股东权益合计(万元)" Then
                        A_nb(xx, 7) = A_temp(y, Z) / 10000
                    End If
                Next
                xx = xx - 1
            End If
            If xx = 1 Then Exit For
        Next
        '利润表
        For x = 1 To Windows.Count
            If Windows(x).Caption = "lrb_" & Code & ".html" Then
                Workbooks("lrb_" & Code & ".html").Close savechanges:=False
                Exit For
            End If
        Next

        Workbooks.Open "http://quotes.money.163.com/service/lrb_" & Code & ".html"
        Temp = ActiveSheet.UsedRange.Rows.Count
        Temp1 = ActiveSheet.UsedRange.Columns.Count
        A_temp = WorksheetFunction.Transpose(Range(Cells(1, 1), Cells(Temp, Temp1)))
        Workbooks("lrb_" & Code & ".html").Close savechanges:=False

        xx = UBound(A_nb)
        Yy = UBound(A_temp)
        zz = UBound(A_temp, 2)

        For y = 2 To Yy
            If A_nb(xx, 1) = A_temp(y, 1) Then
                For Z = 2 To zz
                    If A_temp(1, Z) = "营业总收入(万元)" Then
                        If IsNumeric(A_temp(y, Z)) = True Then A_nb(xx, 4) = A_temp(y, Z) / 10000
                    ElseIf A_temp(1, Z) = "净利润(万元)" Then
                        If IsNumeric(A_temp(y, Z)) = True Then A_nb(xx, 5) = A_temp(y, Z) / 10000
                    End If
                Next
                xx = xx - 1
            End If
            If xx = 1 Then Exit For
        Next
        '现金流量表
        For x = 1 To Windows.Count
            If Windows(x).Caption = "xjllb_" & Code & ".html" Then
                Workbooks("xjllb_" & Code & ".html").Close savechanges:=False
                Exit For
            End If
        Next

        Workbooks.Open "http://quotes.money.163.com/service/xjllb_" & Code & ".html"
        Temp = ActiveSheet.UsedRange.Rows.Count
        Temp1 = ActiveSheet.UsedRange.Columns.Count
        A_temp = WorksheetFunction.Transpose(Range(Cells(1, 1), Cells(Temp, Temp1)))
        Workbooks("xjllb_" & Code & ".html").Close savechanges:=False

        xx = UBound(A_nb)
        Yy = UBound(A_temp)
        zz = UBound(A_temp, 2)
        For y = 2 To Yy
            If A_nb(xx, 1) = A_temp(y, 1) Then
                For Z = 2 To zz
                    If A_temp(1, Z) = " 经营活动产生的现金流量净额(万元)" Then
                        A_nb(xx, 6) = A_temp(y, Z) / 10000
                    End If
                Next
                xx = xx - 1
            End If
            If xx = 1 Then Exit For
        Next
        '年度
        For x = 2 To UBound(A_nb)
            If Month(A_nb(x, 1)) = 12 Then
                A_nb(x, 1) = Year(A_nb(x, 1))
            Else
                A_nb(x, 1) = Year(A_nb(x, 1)) & 0 & Month(A_nb(x, 1))
            End If
        Next
        '**********   从网易获取财务报表END   **********
    Else
        CBsoure = "新浪"
        '**********   从新浪获取财务报表   **********
        For x = 1 To Windows.Count
            If Windows(x).Caption = "all.phtml" Then
                Workbooks("all.phtml").Close savechanges:=False
                Exit For
            End If
        Next
        '资产负债表
        Workbooks.Open "http://money.finance.sina.com.cn/corp/go.php/vDOWN_BalanceSheet/displaytype/4/stockid/" & Code & "/ctrl/all.phtml"
        Temp = ActiveSheet.UsedRange.Rows.Count
        Temp1 = ActiveSheet.UsedRange.Columns.Count
        A_temp = WorksheetFunction.Transpose(Range(Cells(1, 1), Cells(Temp, Temp1)))
        Workbooks("all.phtml").Close savechanges:=False
        Temp = 1
        For x = 2 To UBound(A_temp)
            If Mid(A_temp(x, 1), 5, 2) - 12 = 0 Then Temp = Temp + 1
        Next

        If nbH > Temp Then nbH = Temp
        If Mid(A_temp(2, 1), 5, 2) - 12 = 0 Then
            ReDim A_nb(1 To nbH, 1 To 7)
            xx = UBound(A_nb)
        Else
            ReDim A_nb(1 To nbH + 1, 1 To 7)
            A_nb(nbH + 1, 1) = Left(A_temp(2, 1), 6)
            xx = UBound(A_nb) - 1
        End If

        For x = 2 To UBound(A_temp)
            If Mid(A_temp(x, 1), 5, 2) - 12 = 0 Then
                A_nb(xx, 1) = Left(A_temp(x, 1), 4)
                xx = xx - 1
            End If
            If xx = 1 Then Exit For
        Next

        A_nb(1, 1) = "年度"
        A_nb(1, 2) = "股东权益"
        A_nb(1, 3) = "资产总计"
        A_nb(1, 4) = "营业收入"
        A_nb(1, 5) = "净利润"
        A_nb(1, 6) = "经营现金净流量"
        A_nb(1, 7) = "母公司股东权益"

        xx = UBound(A_nb)
        Yy = UBound(A_temp)
        zz = UBound(A_temp, 2)
        For y = 2 To Yy
            If Left(A_nb(xx, 1) & 12, 6) - Left(A_temp(y, 1), 6) = 0 Then
                For Z = 2 To zz
                    If InStr("所有者权益(或股东权益)合计/股东权益合计/所有者权益合计", A_temp(1, Z)) Then
                        A_nb(xx, 2) = A_temp(y, Z) / 100000000
                    ElseIf InStr("资产总计", A_temp(1, Z)) Then
                        A_nb(xx, 3) = A_temp(y, Z) / 100000000
                    ElseIf InStr("归属于母公司股东权益合计/归属于母公司股东的权益合计", A_temp(1, Z)) Then
                        A_nb(xx, 7) = A_temp(y, Z) / 100000000
                    End If
                Next
                xx = xx - 1
            End If
            If xx = 1 Then Exit For
        Next
        '利润表
        Workbooks.Open "http://money.finance.sina.com.cn/corp/go.php/vDOWN_ProfitStatement/displaytype/4/stockid/" & Code & "/ctrl/all.phtml"
        Temp = ActiveSheet.UsedRange.Rows.Count
        Temp1 = ActiveSheet.UsedRange.Columns.Count
        A_temp = WorksheetFunction.Transpose(Range(Cells(1, 1), Cells(Temp, Temp1)))
        Workbooks("all.phtml").Close savechanges:=False

        xx = UBound(A_nb)
        Yy = UBound(A_temp)
        zz = UBound(A_temp, 2)

        For y = 2 To Yy
            If Left(A_nb(xx, 1) & 12, 6) - Left(A_temp(y, 1), 6) = 0 Then
                For Z = 2 To zz
                    If A_temp(1, Z) = "一、营业总收入" Or A_temp(1, Z) = "一、营业收入" Then
                        A_nb(xx, 4) = A_temp(y, Z) / 100000000
                    ElseIf A_temp(1, Z) = "五、净利润" Then
                        A_nb(xx, 5) = A_temp(y, Z) / 100000000
                    End If
                Next
                xx = xx - 1
            End If
            If xx = 1 Then Exit For
        Next
        '现金流量表
        Workbooks.Open "http://money.finance.sina.com.cn/corp/go.php/vDOWN_CashFlow/displaytype/4/stockid/" & Code & "/ctrl/all.phtml"
        Temp = ActiveSheet.UsedRange.Rows.Count
        Temp1 = ActiveSheet.UsedRange.Columns.Count
        A_temp = WorksheetFunction.Transpose(Range(Cells(1, 1), Cells(Temp, Temp1)))
        Workbooks("all.phtml").Close savechanges:=False

        xx = UBound(A_nb)
        Yy = UBound(A_temp)
        zz = UBound(A_temp, 2)
        For y = 2 To Yy
            If Left(A_nb(xx, 1) & 12, 6) - Left(A_temp(y, 1), 6) = 0 Then
                For Z = 2 To zz
                    If A_temp(1, Z) = "经营活动产生的现金流量净额" Then
                        A_nb(xx, 6) = A_temp(y, Z) / 100000000
                    End If
                Next
                xx = xx - 1
            End If
            If xx = 1 Then Exit For
        Next
        '**********   从新浪获取财务报表END   **********
    End If

    '**********   全景表   **********
    nbH = UBound(A_nb)
    ReDim A_qjbsc(1 To nbH, 1 To 10)
    A_qjbsc(1, 1) = "年度"
    A_qjbsc(1, 2) = "营业收入"
    A_qjbsc(1, 3) = "同比增长%"
    A_qjbsc(1, 4) = "净利润"
    A_qjbsc(1, 5) = "同比增长"
    A_qjbsc(1, 6) = "经营现金净流量"
    A_qjbsc(1, 7) = "母公司股东权益"
    A_qjbsc(1, 8) = "资产负债率"
    A_qjbsc(1, 9) = "销售净利率"
    A_qjbsc(1, 10) = "摊薄ROE"

    For x = 2 To nbH
        A_qjbsc(x, 1) = A_nb(x, 1)
        If A_nb(x, 4) = Empty Then
        Else
            A_qjbsc(x, 2) = A_nb(x, 4)
        End If
        If A_nb(x, 5) = Empty Then
        Else
            A_qjbsc(x, 4) = A_nb(x, 5)
        End If
        If A_nb(x, 6) = Empty Then
        Else
            A_qjbsc(x, 6) = A_nb(x, 6)
        End If
        If A_nb(x, 7) = Empty Then
            If A_nb(x, 2) = Empty Then
            Else
                A_qjbsc(1, 7) = "股东权益合计"
                A_qjbsc(x, 7) = A_nb(x, 2)
            End If
        Else
            A_qjbsc(x, 7) = A_nb(x, 7)
        End If
        If A_nb(x, 3) = Empty Or A_nb(x, 2) = Empty Then
        Else
            A_qjbsc(x, 8) = 1 - A_nb(x, 2) / A_nb(x, 3)
        End If
        If A_nb(x, 4) = Empty Or A_nb(x, 5) = Empty Then
        Else
            A_qjbsc(x, 9) = A_nb(x, 5) / A_nb(x, 4)
        End If
        If A_nb(x, 2) = Empty Or A_nb(x, 5) = Empty Then
        Else
            Temp = Mid(A_nb(x, 1) & 12, 5, 2) / 3
            A_qjbsc(x, 10) = (A_nb(x, 5) / A_nb(x, 2)) * 4 / Temp
        End If
    Next

    If Mid(A_nb(nbH, 1) & 12, 5, 2) - 12 <> 0 Then nbH = nbH - 1
    For x = 3 To nbH
        If A_qjbsc(x - 1, 2) = Empty Then
        Else
            A_qjbsc(x, 3) = A_qjbsc(x, 2) / A_qjbsc(x - 1, 2) - 1
        End If
        If A_qjbsc(x - 1, 4) = Empty Then
        Else
            If A_qjbsc(x, 4) > 0 And A_qjbsc(x - 1, 4) > 0 Then
                A_qjbsc(x, 5) = A_qjbsc(x, 4) / A_qjbsc(x - 1, 4) - 1
            Else
                A_qjbsc(x, 5) = "-"
            End If
        End If
    Next
    '**********   全景表END   **********
    '**********   从同花顺抓取F10资料   **********
    URL = "http://basic.10jqka.com.cn/16/" & Code & "/"
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", URL, False
        .send
        Gdsj = StrConv(.responseBody, vbUnicode) '.responsetext
    End With
    GoSub GosubLM

    Won = ">"
    Woff = "<"
    GoSub GosubZL
    

    A_gdsj = Split(Gdsj, ">")
    y = UBound(A_gdsj)
    For x = 1 To y
        Select Case A_gdsj(x)
        Case "每股收益："
            MGsy = A_gdsj(x + 1)
        Case "每股净资产："
            MGjzc = A_gdsj(x + 1)
        Case "市净率："
            xPB = A_gdsj(x + 1)
        Case "市盈率(动态)："
            dpe = A_gdsj(x + 1)
        Case "市盈率(静态)："
            jpe = A_gdsj(x + 1)
        End Select
    Next
    xPE = dpe & "/" & jpe

    URL = "http://basic.10jqka.com.cn/32/" & Code & "/company.html"
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", URL, False
        .send
        Gdsj = StrConv(.responseBody, vbUnicode)
    End With

    GoSub GosubLM

    Won = ">"
    Woff = "<"
    GoSub GosubZL

    A_gdsj = Split(Gdsj, ">")
    y = UBound(A_gdsj)
    For x = 1 To y
        Select Case A_gdsj(x)
        Case "所属地域："
            GSdy = A_gdsj(x + 1)
        Case "所属行业："
            GShy = A_gdsj(x + 1)
        Case "公司名称："
            GSname = A_gdsj(x + 1)
        Case "主营业务："
            GSzy = A_gdsj(x + 1)
        Case "上市日期："
            Sssj = A_gdsj(x + 1)
        End Select
    Next
    '**********   从同花顺抓取F10资料END   **********

    '**********   从同花顺抓业绩预测   **********
    Dim A_yjyc(1 To 4, 1 To 8)    '业绩预测：年度,机构数,最小，均值,最大，增长率,PE,PEG
    URL = "http://basic.10jqka.com.cn/16/" & Code & "/worth.html"
    
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", URL, False
        .send
        Gdsj = StrConv(.responseBody, vbUnicode)
    End With

    Won = "<table"
    Woff = "</table>"
    GoSub GosubZL

    A_gdsj = Split(Gdsj, "<table")
    xx = UBound(A_gdsj)
    For x = 1 To xx
        If InStr(A_gdsj(x), "汇总--预测年报净利润") Then
            Gdsj = A_gdsj(x)
            Exit For
        Else
            Gdsj = ""
        End If
    Next
    yearf10 = Gdsj

    Won = "<td"
    Woff = "</td>"
    GoSub GosubZL

    Won = ">"
    Woff = "<"
    GoSub GosubZL

    A_gdsj = Split(Gdsj, ">")

    Gdsj = yearf10
    Won = "<th"
    Woff = "</th>"
    GoSub GosubZL

    Won = ">"
    Woff = "<"
    GoSub GosubZL

    a_yearf10 = Split(Gdsj, ">")
    For x = 0 To UBound(a_yearf10)
        If Left(a_yearf10(x), 1) = 2 Then
            A_yjyc(2, 1) = a_yearf10(x)
            A_yjyc(3, 1) = A_yjyc(2, 1) + 1
            A_yjyc(4, 1) = A_yjyc(2, 1) + 2
            Exit For
        End If
    Next

    nbH = UBound(A_qjbsc)
    '    If Mid(A_qjbsc(nbH, 1) & 12, 5, 2) - 12 <> 0 Then nbH = nbH - 1
    For x = 2 To UBound(A_qjbsc)
        If A_yjyc(2, 1) - A_qjbsc(x, 1) = 1 Then
            nbH = x
            Exit For
        End If
    Next
    A_yjyc(1, 1) = "年度"
    A_yjyc(1, 2) = "机构数"
    A_yjyc(1, 3) = "最小"
    A_yjyc(1, 4) = "平均"
    A_yjyc(1, 5) = "最大"
    A_yjyc(1, 6) = "增长率"
    A_yjyc(1, 7) = "PE"
    A_yjyc(1, 8) = "PEG"
    If UBound(A_gdsj) < 5 Then
    Else
        A_yjyc(2, 2) = A_gdsj(1)
        A_yjyc(2, 3) = A_gdsj(2)
        A_yjyc(2, 4) = A_gdsj(3)
        A_yjyc(2, 5) = A_gdsj(4)
        If IsNumeric(A_gdsj(3)) Then A_yjyc(2, 6) = A_gdsj(3) / A_qjbsc(nbH, 4) - 1
        If IsNumeric(A_gdsj(3)) Then A_yjyc(2, 7) = Zsz / A_gdsj(3)
        If IsNumeric(A_gdsj(3)) Then A_yjyc(2, 8) = A_yjyc(2, 7) / (A_yjyc(2, 6) * 100)
        If UBound(A_gdsj) < 10 Then
        Else
            A_yjyc(3, 2) = A_gdsj(6)
            A_yjyc(3, 3) = A_gdsj(7)
            A_yjyc(3, 4) = A_gdsj(8)
            A_yjyc(3, 5) = A_gdsj(9)
            If IsNumeric(A_gdsj(8)) Then A_yjyc(3, 6) = A_gdsj(8) / A_gdsj(3) - 1
            If IsNumeric(A_gdsj(8)) Then A_yjyc(3, 7) = Zsz / A_gdsj(8)
            If IsNumeric(A_gdsj(8)) Then A_yjyc(3, 8) = A_yjyc(3, 7) / (A_yjyc(3, 6) * 100)
            If UBound(A_gdsj) < 15 Then
            Else
                A_yjyc(4, 2) = A_gdsj(11)
                A_yjyc(4, 3) = A_gdsj(12)
                A_yjyc(4, 4) = A_gdsj(13)
                A_yjyc(4, 5) = A_gdsj(14)
                If IsNumeric(A_gdsj(13)) Then A_yjyc(4, 6) = A_gdsj(13) / A_gdsj(8) - 1
                If IsNumeric(A_gdsj(13)) Then A_yjyc(4, 7) = Zsz / A_gdsj(13)
                If IsNumeric(A_gdsj(13)) Then A_yjyc(4, 8) = A_yjyc(4, 7) / (A_yjyc(4, 6) * 100)
            End If
        End If
    End If
    '**********   从同花顺抓业绩预测END   **********

    '**********   输出   **********
    ThisWorkbook.Activate
    For x = 1 To Sheets.Count
        If Sheets(x).Name = "全景表" Then
            Sheets("全景表").Select
            Application.DisplayAlerts = False
            ActiveWindow.SelectedSheets.Delete
            Application.DisplayAlerts = True
            Worksheets.Add().Name = "全景表"
            Exit For
        End If
        If x = Sheets.Count Then
            If IsEmpty(ActiveSheet.UsedRange) Then
                ActiveSheet.Name = "全景表"
            Else
                Worksheets.Add().Name = "全景表"
            End If
        End If
    Next



    '输出全景表
    xx = UBound(A_qjbsc)
    Yy = UBound(A_qjbsc, 2)
    For x = 1 To xx
        For y = 1 To Yy
            Cells(x + 3, y) = A_qjbsc(x, y)
        Next
    Next
    '输出历史行情
    Cells(4, 11) = "年度"
    Cells(4, 12) = "涨幅"
    Cells(4, 13) = "期末市值"
    xx = UBound(A_qjbsc) + 4
    Yy = 1
    For x = 5 To xx + 3
        For y = 1 To UBound(A_lshqb) - 1
            If Cells(x, 1) - A_lshqb(y, 1) = 0 Then
                xx = x
                Yy = y
                x = UBound(A_qjbsc) + 4
                Exit For
            End If
        Next
    Next
    For x = Yy To UBound(A_lshqb)
        Cells(xx + x - Yy, 11) = A_lshqb(x, 1)
        Cells(xx + x - Yy, 12) = A_lshqb(x, 4)
        Cells(xx + x - Yy, 13) = A_lshqb(x, 2)
    Next
    Cells(4, 14) = "历史PE"
    For x = 2 To UBound(A_qjbsc)
        If Cells(x + 4, 11) = "上市累计" Then Exit For
        If Cells(x + 4, 1) - Cells(x + 4, 11) = 0 And Cells(x + 4, 4) > 0 Then
            Cells(x + 4, 14) = Cells(x + 4, 13) / Cells(x + 4, 4)
        End If
    Next

    '输出全景表表尾
    nbH = UBound(A_qjbsc)
    Cells(nbH + 4, 1) = "总股本"
    If Price > 0 Then Cells(nbH + 4, 2) = Round(Zsz / Price, 2)
    Cells(nbH + 4, 3) = "每股收益"
    Cells(nbH + 4, 4) = MGsy
    Cells(nbH + 4, 5) = "每股净资产"
    Cells(nbH + 4, 6) = MGjzc
    Cells(nbH + 4, 7) = "当前价格"
    Cells(nbH + 4, 8) = Price & "元"
    Cells(nbH + 4, 9) = "研发占营收比"
    Cells(nbH + 4, 10) = ""
    Cells(nbH + 5, 1) = "PE(动态/静态)"
    Cells(nbH + 5, 2) = xPE
    Cells(nbH + 5, 3) = "PB"
    Cells(nbH + 5, 4) = xPB
    Cells(nbH + 5, 5) = "市值"
    Cells(nbH + 5, 6) = Zsz
    Cells(nbH + 5, 7) = "现金利润比"
    For x = 1 To nbH - 1
        If A_qjbsc(nbH + 1 - x, 6) = Empty Then
            Exit For
        Else
            xj = xj + A_qjbsc(nbH + 1 - x, 6)
            lr = lr + A_qjbsc(nbH + 1 - x, 4)
        End If
    Next
    Cells(nbH + 5, 8) = xj / lr
    Cells(nbH + 5, 9) = "分红率平均"
    Cells(nbH + 5, 10) = ""
    Cells(nbH + 6, 1) = "营收复合增长率"
    xx = nbH
    If Mid(A_qjbsc(nbH, 1) & 12, 5, 2) - 12 <> 0 Then xx = nbH - 1
    For x = 1 To xx - 2
        If A_qjbsc(xx - x, 2) > 0 Then
            Cells(nbH + 6, 2) = (A_qjbsc(xx, 2) / A_qjbsc(xx - x, 2)) ^ (1 / x) - 1
        End If
    Next
    Cells(nbH + 6, 3) = "利润复合增长率"
    For x = 1 To xx - 2
        If A_qjbsc(xx - x, 4) > 0 And A_qjbsc(xx, 4) > 0 Then
            Cells(nbH + 6, 4) = (A_qjbsc(xx, 4) / A_qjbsc(xx - x, 4)) ^ (1 / x) - 1
        End If
    Next
    Cells(nbH + 6, 5) = "平均净利润率"
    Cells(nbH + 6, 6) = Application.WorksheetFunction.Average(Range(Cells(5, 9), Cells(nbH + 3, 9)))
    Cells(nbH + 6, 7) = "平均ROE"
    Cells(nbH + 6, 8) = Application.WorksheetFunction.Average(Range(Cells(5, 10), Cells(nbH + 3, 10)))
    Cells(nbH + 6, 9) = "平均负债率"
    Cells(nbH + 6, 10) = Application.WorksheetFunction.Average(Range(Cells(5, 8), Cells(nbH + 3, 8)))
    Cells(nbH + 7, 1) = "日期：" & zxjyr
    '输出业绩预测
    nbH = UBound(A_qjbsc)
    Cells(nbH + 8, 3) = "业绩(净利润)预测"
    Cells(nbH + 8, 7) = "来源:同花顺"
    Range(Cells(nbH + 9, 2), Cells(nbH + 12, 9)) = A_yjyc
    '输出报表链接
    Cells(nbH + 14, 2) = Name & "财报"
    Cells(nbH + 15, 1) = "源自新浪:"
    Cells(nbH + 15, 2) = "资产负债表"
    Cells(nbH + 15, 3) = "利润表"
    Cells(nbH + 15, 4) = "现金流量表"
    ActiveSheet.Hyperlinks.Add Anchor:=Cells(nbH + 15, 2), Address:="http://money.finance.sina.com.cn/corp/go.php/vDOWN_BalanceSheet/displaytype/4/stockid/" & Code & "/ctrl/all.phtml"
    ActiveSheet.Hyperlinks.Add Anchor:=Cells(nbH + 15, 3), Address:="http://money.finance.sina.com.cn/corp/go.php/vDOWN_ProfitStatement/displaytype/4/stockid/" & Code & "/ctrl/all.phtml"
    ActiveSheet.Hyperlinks.Add Anchor:=Cells(nbH + 15, 4), Address:="http://money.finance.sina.com.cn/corp/go.php/vDOWN_CashFlow/displaytype/4/stockid/" & Code & "/ctrl/all.phtml"
    Cells(nbH + 16, 1) = "源自网易:"
    Cells(nbH + 16, 2) = "资产负债表"
    Cells(nbH + 16, 3) = "利润表"
    Cells(nbH + 16, 4) = "现金流量表"
    ActiveSheet.Hyperlinks.Add Anchor:=Cells(nbH + 16, 2), Address:="http://quotes.money.163.com/service/zcfzb_" & Code & ".html"
    ActiveSheet.Hyperlinks.Add Anchor:=Cells(nbH + 16, 3), Address:="http://quotes.money.163.com/service/lrb_" & Code & ".html"
    ActiveSheet.Hyperlinks.Add Anchor:=Cells(nbH + 16, 4), Address:="http://quotes.money.163.com/service/xjllb_" & Code & ".html"
    '输出K线图
    If Application.Version - 12 = 0 Then
    Else
        Range("A1").Select
        If Left(Code, 1) = 6 Then
            Temp = "sh" & Code & ".gif"
        Else
            Temp = "sz" & Code & ".gif"
        End If
        ActiveSheet.Pictures.Insert("http://image.sinajs.cn/newchart/daily/n/" & Temp).Select
        Selection.ShapeRange.IncrementLeft 700
        Selection.ShapeRange.IncrementTop 10
        ActiveSheet.Pictures.Insert("http://image.sinajs.cn/newchart/weekly/n/" & Temp).Select
        Selection.ShapeRange.IncrementLeft 700
        Selection.ShapeRange.IncrementTop 230
        ActiveSheet.Pictures.Insert("http://image.sinajs.cn/newchart/monthly/n/" & Temp).Select
        Selection.ShapeRange.IncrementLeft 700
        Selection.ShapeRange.IncrementTop 470
    End If
    '**********   输出END   **********

    '**********   外观优化   **********
    nbH = UBound(A_qjbsc)
    Range("1:1,4:4").Select
    GoSub GosubJZ
    Range(Cells(4, 1), Cells(nbH + 6, 12)).Select
    GoSub GosubJZ

    Range("B:B,D:D,F:F,G:G,m:n").Select
    Selection.NumberFormatLocal = "0.00_ "

    Range("C:C,E:E,H:J,l:l").Select
    Selection.NumberFormatLocal = "0.00%"

    Cells.Select
    Cells.EntireColumn.AutoFit

    Range("f1").Select
    With Selection.Font
        .Name = "微软雅黑"
        .Size = 24
    End With

    Range(Cells(nbH + 4, 1), Cells(nbH + 12, 10)).Select
    Selection.NumberFormatLocal = "G/通用格式"
    GoSub GosubJZ

    Cells(nbH + 5, 8).Select
    Selection.NumberFormatLocal = "0.00%"
    For x = 1 To 5
        Cells(nbH + 6, 2 * x).Select
        Selection.NumberFormatLocal = "0.00%"
    Next

    '    GoSub GosubJZ

    Range(Cells(4, 1), Cells(nbH + 6, 14)).Select
    GoSub GosubBK
    Range(Cells(nbH + 9, 2), Cells(nbH + 12, 9)).Select
    GoSub GosubBK

    Range(Cells(nbH + 9, 4), Cells(nbH + 12, 9)).Select
    Selection.NumberFormatLocal = "0.00_ "
    Range(Cells(nbH + 9, 7), Cells(nbH + 12, 7)).Select
    Selection.NumberFormatLocal = "0.00%"
    '**********   外观优化END   **********
    '**********   输出全景表表头   **********
    Cells(1, 6) = Name & "(" & Code & ")全景表"
    Cells(1, 11) = "单位：亿元"
    Cells(2, 1) = GSname
    Cells(2, 4) = "地域：" & GSdy
    Cells(2, 6) = "行业：" & GShy
    Cells(2, 10) = "上市日期：" & Sssj
    Cells(3, 1) = "主营：" & GSzy
    Cells(3, 10) = "财报数据来自：" & CBsoure
    '**********   输出全景表表头END   **********
    '    xx = UBound(A_lshqb)
    '    For x = 1 To xx
    '        Cells(x + 1, 15) = A_lshqb(x, 1)
    '        Cells(x + 1, 16) = A_lshqb(x, 2)
    '        Cells(x + 1, 17) = A_lshqb(x, 3)
    '        Cells(x + 1, 18) = A_lshqb(x, 4)
    '    Next
    '    xx = UBound(A_nb)
    '    Yy = UBound(A_nb, 2)
    '    For x = 1 To xx
    '        For y = 1 To Yy
    '            Cells(x, y + 19) = A_nb(x, y)
    '        Next
    '    Next

    Range("A1").Select
    Errxx = "生成完毕，耗时" & Format(Timer - StarTime, "0.00秒")
    GoTo ErrMsg

GosubLM:        '清理乱码
    Temp = ""
    For x = 1 To Len(Gdsj)
        If InStr(Chr(8) & Chr(9) & Chr(10) & Chr(13) & Chr(32), Mid(Gdsj, x, 1)) Then
        Else
            Temp = Temp & Mid(Gdsj, x, 1)
        End If
    Next
    Gdsj = Temp
    Temp = ""
    Return

GosubZL:        '数据整理
    W = 0
    Temp = ""
    For x = 1 To Len(Gdsj)
        If Mid(Gdsj, x, Len(Won)) = Won Then W = 1
        If Mid(Gdsj, x, Len(Woff)) = Woff Then W = 0
        If W = 1 Then
            If Mid(Gdsj, x, Len(Won)) = " " Or (Right(Temp, Len(Won)) = Won And Mid(Gdsj, x, Len(Won)) = Won) Then
            Else
                Temp = Temp + Mid(Gdsj, x, 1)
            End If
        End If
    Next
    Gdsj = Temp
    Temp = ""
    Return

GosubJZ:        '格式居中
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
    End With
    Return

GosubBK:        '格式加边框
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
    End With
    Return

ErrMsg:
    With UserForm1
        .Label1.Caption = Errxx
        .CommandButton1.Enabled = True
    End With
End Sub