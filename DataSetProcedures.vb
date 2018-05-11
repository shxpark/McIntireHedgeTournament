Module DataSetProcedures

    Public Function GetInitialCAccount() As Double
        Dim name, value As String
        For Each myRow As DataRow In myDataSet.Tables("SettingsTbl").Rows
            name = myRow("Name").ToString().Trim()
            If name = "CAccount" Then
                value = myRow("Value").ToString().Trim()
                Return Double.Parse(value)
            End If
        Next
        MessageBox.Show("Holy Cow! Could not find 'CAccount'. Returned 0.")
        Return 0
    End Function

    Public Function GetIRate() As Double
        Dim name, value As String
        For Each myRow As DataRow In myDataSet.Tables("SettingsTbl").Rows
            name = myRow("Name").ToString.Trim()
            If name = "RiskFreeRate" Then
                value = myRow("Value").ToString().Trim()
                Return Double.Parse(value)
            End If
        Next
        MessageBox.Show("Holy Cow! Could not find 'IRate'. Returned 0.")
        Return 0
    End Function

    Public Function GetStartDate() As Date
        Dim name, value As String
        For Each myRow As DataRow In myDataSet.Tables("SettingsTbl").Rows
            name = myRow("Name").ToString().Trim()
            If name = "StartDate" Then
                value = myRow("Value").ToString().Trim()
                Return Date.Parse(value)
            End If
        Next
        MessageBox.Show("Holy Cow! Could not find 'StartDate'. Returned nothing.")
        Return Nothing
    End Function

    Public Function GetEndDate() As Date
        Dim name, value As String
        For Each myRow As DataRow In myDataSet.Tables("SettingsTbl").Rows
            name = myRow("Name").ToString().Trim()
            If name = "EndDate" Then
                value = myRow("Value").ToString().Trim()
                Return Date.Parse(value)
            End If
        Next
        MessageBox.Show("Holy Cow! Could not find 'EndDate'. Returned Nothing.")
        Return Nothing
    End Function

    Public Function GetMaxMargins() As Double
        Dim name, value As String
        For Each myRow As DataRow In myDataSet.Tables("SettingsTbl").Rows
            name = myRow("Name").ToString().Trim()
            If name = "MaxMargins" Then
                value = myRow("Value").ToString().Trim()
                Return Double.Parse(value)
            End If
        Next
        MessageBox.Show("Holy Cow! Could not find 'MaxMargin'. Returned 0")
        Return 0
    End Function

    Public Function GetAsk(symbol As String, targetDate As Date) As Double
        symbol = symbol.Trim()
        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If
        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        If targetDate.Date <> lastPriceDownloadDate.Date Then
            DownloadPricesForOneDay(targetDate)
        End If

        If IsAStock(symbol) Then
            For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTbl").Rows
                If myRow("Ticker").trim() = symbol And
                        myRow("Date") = targetDate.ToShortDateString() Then
                    Return myRow("Ask")
                End If
            Next
        Else
            For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTbl").Rows
                If myRow("Symbol").trim() = symbol And
                        myRow("Date") = targetDate.ToShortDateString() Then
                    Return myRow("Ask")
                End If
            Next
        End If
        MessageBox.Show("Holy Cow! Could not find the ask for " + symbol + ". Returned 0.")
        Return 0
    End Function

    Public Function GetBid(symbol As String, targetDate As Date) As Double
        symbol = symbol.Trim()
        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If
        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        If targetDate.Date <> lastPriceDownloadDate.Date Then
            DownloadPricesForOneDay(targetDate)
        End If

        If IsAStock(symbol) Then
            For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTbl").Rows
                If myRow("Ticker").trim() = symbol And
                        myRow("Date") = targetDate.ToShortDateString() Then
                    Return myRow("Bid")
                End If
            Next
        Else
            For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTbl").Rows
                If myRow("Symbol").trim() = symbol And
                        myRow("Date") = targetDate.ToShortDateString() Then
                    Return myRow("Bid")
                End If
            Next
        End If
        MessageBox.Show("Holy Cow! Could not the bid for" + symbol + ". Returned 0.")
        Return 0
    End Function

    Public Function GetCurrentPositionInAP(symbol) As Double
        For Each myRow As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            If myRow("Symbol").ToString().Trim() = symbol Then
                Return Double.Parse(myRow("Units"))
            End If
        Next
        Return 0
    End Function

    Public Function GetCurrentPositionInIP(symbol) As Double
        For Each myRow As DataRow In myDataSet.Tables("InitialPositionsTbl").Rows
            If myRow("Symbol").ToString().Trim() = symbol Then
                Return Double.Parse(myRow("Units"))
            End If
        Next
        Return 0
    End Function

    Public Function GetDividend(ticker As String, targetDate As Date) As Double
        If IsAStock(ticker) Then
            If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
                targetDate = targetDate.AddDays(-1)
            End If
            If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
                targetDate = targetDate.AddDays(-2)
            End If

            If targetDate.Date <> lastPriceDownloadDate.Date Then
                DownloadPricesForOneDay(targetDate)
            End If

            For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTbl").Rows
                If myRow("Ticker").trim() = ticker And myRow("Date") = targetDate.ToShortDateString() Then
                    Return Double.Parse(myRow("Dividend"))
                End If
            Next
        End If
        MessageBox.Show("Holy Cow. I could not find the dividend for " + ticker + ". Returned 0.")
        Return 0
    End Function

    Public Function GetStrike(symbol As String) As Double
        For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTbl").Rows
            If myRow("Symbol").trim() = symbol Then
                Return Double.Parse(myRow("Strike"))
            End If
        Next
        MessageBox.Show("Holy Cow! Could not find the strike for " + symbol + ". Returned 0.")
        Return 0
    End Function

    Public Function GetTrCostCoefficient(secType As String, trType As String) As Double
        For Each myRow As DataRow In myDataSet.Tables("TransactionCostsTbl").Rows
            If myRow("SecurityType").Trim() = secType And myRow("TransactionType").Trim() = trType Then
                Return Double.Parse(myRow("CostCoeff"))
            End If
        Next
        MessageBox.Show("Holy cow! Could not find the transaction cost. Returned 0.")
        Return 0
    End Function


    Public Function GetUnderlier(symbol As String) As String
        symbol = symbol.Trim()
        If Not myDataSet.Tables.Contains("OptionMarketOneDayTbl") Then
            DownloadPricesForOneDay(currentDate)
        End If
        For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTbl").Rows
            If myRow("Symbol").trim() = symbol Then
                Return myRow("Underlier").Trim()
            End If
        Next
        MessageBox.Show("Holy Cow! Could not find the underlier for " + symbol + ". Returned 0.")
        Return ""
    End Function

    Public Function GetExpiration(symbol As String) As Date
        symbol = symbol.Trim()
        For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTbl").Rows
            If myRow("Symbol").trim() = symbol Then
                Return Date.Parse(myRow("Expiration"))
            End If
        Next
        MessageBox.Show("Holy Cow! Could not find the expiration for " + symbol + ". Returned 0.")
        Return Nothing
    End Function

End Module
