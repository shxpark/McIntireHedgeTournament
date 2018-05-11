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
End Module
