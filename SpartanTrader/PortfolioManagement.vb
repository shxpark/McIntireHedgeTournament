Module PortfolioManagement

    Public Function CalcTPVAtStart() As Double
        Return CalcIPValue(startDate) + initialCAccount
    End Function

    Public Function CalcIPValue(targetDate As Date) As Double
        Dim cumulativeValue As Double = 0
        Dim symbol As String
        Dim units As Double
        Dim posValue As Double

        For Each myRow As DataRow In myDataSet.Tables("InitialPositionsTbl").Rows
            symbol = myRow("Symbol").ToString().Trim
            units = myRow("Units")
            posValue = units * CalcMTM(symbol, targetDate)
            cumulativeValue = cumulativeValue + posValue
            myRow("Value") = posValue
        Next

        Return cumulativeValue
    End Function

    Public Function CalcMTM(symbol As String, targetDate As Date) As Double
        Return (GetAsk(symbol, targetDate) + GetBid(symbol, targetDate)) / 2
    End Function

    Public Function IsAStock(symbol As String) As Boolean
        symbol = symbol.Trim()
        For Each myRow As DataRow In myDataSet.Tables("TickersTbl").Rows
            If myRow("Ticker").trim() = symbol Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function CalcAPValue(targetDate As Date) As Double
        Dim cumulativeValue As Double = 0
        Dim symbol As String
        Dim units As Double
        Dim posValue As Double

        For Each myRow As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            symbol = myRow("Symbol").ToString().Trim
            units = myRow("Units")
            If symbol <> "CAccount" Then
                posValue = units * CalcMTM(symbol, targetDate)
                cumulativeValue = cumulativeValue + posValue
                myRow("Value") = posValue
            End If
        Next
        Return cumulativeValue
    End Function
End Module
