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
        If symbol = "CAccount" Then
            Return 1
        Else
            Return (GetAsk(symbol, targetDate) + GetBid(symbol, targetDate)) / 2
        End If
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

    Public Function IsAnOption(symbol As String) As Boolean
        symbol = symbol.Trim()
        For Each myRow As DataRow In myDataSet.Tables("SymbolsTbl").Rows
            If myRow("Symbol").trim() = symbol Then
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

    Public Function IsInIP(s As String) As Boolean
        s = s.Trim()
        For Each myRow As DataRow In myDataSet.Tables("InitialPositionsTbl").Rows
            If myRow("Symbol").trim() = s Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function GetCurrentPosition(sym As String)
        Return GetCurrentPositionInAP(sym) + GetCurrentPositionInIP(sym)
    End Function

    Public Function CalcInterestSLT(toThisDay As Date) As Double
        Dim interest As Double = 0
        Dim ts As TimeSpan = toThisDay.Date - lastTransactionDate.Date
        Dim t As Double = ts.Days / 365.25
        interest = CAccount * (Math.Exp(iRate * t) - 1)
        Return interest
    End Function

    Public Function CalcMargin(targetDate As Date) As Double
        Return CalcAPMargin(targetDate) + CalcIPMargin(targetDate)
    End Function

    Public Function CalcAPMargin(targetDate As Date) As Double
        Dim cumulativeValue As Double = 0
        Dim symbol As String
        Dim units As Double
        Dim posValue As Double

        For Each myRow As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            symbol = myRow("Symbol").ToString().Trim
            units = myRow("Units")
            If symbol <> "CAccount" And units < 0 Then
                posValue = units * CalcMTM(symbol, targetDate)
                cumulativeValue = cumulativeValue + posValue
            End If
        Next
        Return cumulativeValue
    End Function

    Public Function CalcIPMargin(targetDate As Date) As Double
        Dim cumulativeValue As Double = 0
        Dim symbol As String
        Dim units As Double
        Dim posValue As Double

        For Each myRow As DataRow In myDataSet.Tables("InitialPositionsTbl").Rows
            symbol = myRow("Symbol").ToString().Trim
            units = myRow("Units")
            If symbol <> "CAccount" And units < 0 Then
                posValue = units * CalcMTM(symbol, targetDate)
                cumulativeValue = cumulativeValue + posValue
            End If
        Next
        Return cumulativeValue
    End Function

    Public Function CalcTaTPV(targetDate As Date) As Double
        Dim ts As TimeSpan = targetDate.Date - startDate.Date
        Dim t As Double = ts.Days / 365.25
        Return TPVatStart * Math.Exp(iRate * t)
    End Function

    Public Function CalcTE() As Double
        If TPV >= TaTPV Then
            Return (TPV - TaTPV) * 0.25
        Else
            Return TaTPV - TPV
        End If
    End Function

    Public Function IsACall(symbol As String) As Boolean
        symbol = symbol.Trim()
        For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTbl").Rows
            If myRow("Symbol").trim() = symbol And myRow("Type").Trim() = "Call" Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function IsAPut(symbol As String) As Boolean
        symbol = symbol.Trim()
        For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTbl").Rows
            If myRow("Symbol").trim() = symbol And myRow("Type").Trim() = "Put" Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Sub ResetAP()
        If MessageBox.Show("You sure?", "Reset AP?", MessageBoxButtons.YesNo, MessageBoxIcon.Hand) = DialogResult.Yes Then
            initialCAccount = GetInitialCAccount()
            ClearTeamPortfolioOnDB()
            UploadPosition("CAccount", initialCAccount)
            StartTheTrader()
        End If
    End Sub

    Public Sub UploadScreenPortfolioToDB()
        Dim tempSymbol, tempUnits As String
        If Globals.ThisWorkbook.ActiveSheet.Name <> "Dashboard" Then
            MessageBox.Show("Are you looking at the Portfolio that you want me to upload, Dave?", "Portfolio Not Active", MessageBoxButtons.OK, MessageBoxIcon.Hand)

            Return
        End If
        If Globals.Dashboard.AcquiredPositionsLO.IsSelected Then
            MessageBox.Show("Click outside the ListObject to confirm data entry, Dave.", "Edit In Progress", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return
        End If
        ClearTeamPortfolioOnDB()
        For i As Integer = 1 To Globals.Dashboard.AcquiredPositionsLO.DataBodyRange.Rows.Count()
            tempSymbol = Globals.Dashboard.AcquiredPositionsLO.DataBodyRange.Cells(i, 1).Value
            tempUnits = Globals.Dashboard.AcquiredPositionsLO.DataBodyRange.Cells(i, 2).Value
            If IsAPEntryValid(tempSymbol, tempUnits) Then
                UploadPosition(tempSymbol, tempUnits)
            End If
        Next
        Globals.Ribbons.stRibbon.AcquiredPositionsBtn_Click(Nothing, Nothing)
        CAccount = GetCurrentPositionInAP("CAccount")
        StartTheTrader()
    End Sub


    Public Sub UpdatePosition(transType As String, sym As String, q As Double)

        Dim oldPosition, newPosition, newULPosition As Double
        oldPosition = GetCurrentPositionInAP(sym)
        q = Math.Abs(q)

        Select Case transType
            Case "Buy"
                newPosition = oldPosition + q
                UploadPosition(sym, newPosition)

            Case "Sell"
                newPosition = oldPosition - q
                UploadPosition(sym, newPosition)

            Case "SellShort"
                newPosition = oldPosition - q
                UploadPosition(sym, newPosition)

            Case "CashDiv"
            ' only cash effects

            Case "X-Put"
                Dim ul As String = GetUnderlier(sym)
                Dim oldULPosition As Double = GetCurrentPositionInAP(ul)

                If oldPosition > 0 Then
                    newPosition = oldPosition - q
                    UploadPosition(sym, newPosition)
                    newULPosition = oldULPosition - q
                    UploadPosition(ul, newULPosition)
                Else
                    newPosition = oldPosition + q
                    UploadPosition(sym, newPosition)
                    newULPosition = oldULPosition + q
                    UploadPosition(ul, newULPosition)
                End If

            Case "X-Call"
                Dim ul As String = GetUnderlier(sym)
                Dim oldULPosition As Double = GetCurrentPositionInAP(ul)

                If oldPosition > 0 Then
                    newPosition = oldPosition - q
                    UploadPosition(sym, newPosition)
                    newULPosition = oldULPosition + q
                    UploadPosition(ul, newULPosition)
                Else
                    newPosition = oldPosition + q
                    UploadPosition(sym, newPosition)
                    newULPosition = oldULPosition - q
                    UploadPosition(ul, newULPosition)
                End If
        End Select
        UploadPosition("CAccount", CAccountAT)
        GetDataTableFromDB("Select * from " + teamPortfolioTableName + " order by symbol", "AcquiredPositionsTbl")
    End Sub

    Public Function IsInTheFamily(sym As String, ticker As String) As Boolean
        sym = sym.Trim()
        ticker = ticker.Trim()
        If sym = ticker Then
            Return True
        End If
        If sym = "CAccount" Or IsAStock(sym) Then
            Return False
        End If
        If GetUnderlier(sym) = ticker Then
            Return True
        End If
        Return False
    End Function
End Module
