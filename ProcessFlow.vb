Module ProcessFlow

    Public Sub RunDailyRoutine()
        Globals.Dashboard.CurrentDateCell.Value = currentDate.ToLongDateString()
        DoScheduledTransactions(currentDate)
        Select Case traderMode
            Case "Manual", "Synch"
                ClearTransaction()
                CalcFinancialMetrics(currentDate)
                DisplayFinancialMetrics(currentDate)
                CalcRecommendations(currentDate)
                DisplayRecommendations()
            Case "Sim"
                CalcFinancialMetrics(currentDate)
                DisplayFinancialMetrics(currentDate)
                DisplayRecommendations()
                SmartHedgeAll()
            Case "Auto"
                ClearTransaction()
                CalcFinancialMetrics(currentDate)
                DisplayFinancialMetrics(currentDate)
                CalcRecommendations(currentDate)
                DisplayRecommendations()
                SmartHedgeAll()
        End Select
        Globals.Dashboard.UpdateTEChart(currentDate)
        For i = 1 To 10
            Application.DoEvents()
        Next
    End Sub
    Public Sub ClearAllLOS()
        Globals.Dashboard.AcquiredPositionsLO.DataSource = Nothing
        Globals.Dashboard.InitialPositionsLO.DataSource = Nothing
        Globals.Environment.SettingsLO.DataSource = Nothing
        Globals.Environment.TransactionCostLO.DataSource = Nothing
        Globals.Markets.StockMarketLO.DataSource = Nothing
        Globals.Markets.OptionMarketLO.DataSource = Nothing
        Globals.Markets.SP500LO.DataSource = Nothing
        Globals.Transactions.TransactionQueueLO.DataSource = Nothing
        Globals.Transactions.ConfirmationTicketsLO.DataSource = Nothing
        Globals.Dashboard.TELO.DataSource = Nothing
    End Sub

    Public Sub StartTheTrader()
        StopTimers()
        ClearAllLOS()
        ConnectToActiveDB()
        If ThereIsData() = False Then
            Exit Sub
        End If

        Select Case traderMode
            Case "Manual"
                currentDate = DownloadCurrentDate()
                DownloadStaticData()
                DownloadTeamData()
                Globals.Dashboard.InitializeDisplay()
                RunDailyRoutine()

            Case "Sim"
                DownloadStaticData()
                currentDate = GetStartDate()
                DownloadTeamData()
                Globals.Dashboard.InitializeDisplay()
                CalcFinancialMetrics(currentDate)
                CalcRecommendations(currentDate)
                Do
                    RunDailyRoutine()
                    currentDate = currentDate.AddDays(1)
                Loop While currentDate <= endDate

            Case "Synch", "Auto"
                currentDate = DownloadCurrentDate()
                DownloadStaticData()
                DownloadTeamData()
                Globals.Dashboard.InitializeDisplay()
                StartTimers()
        End Select
    End Sub

    Public Sub DownloadStaticData()
        GetDataTableFromDB("Select * from InitialPosition order by symbol", "InitialPositionsTbl")
        GetDataTableFromDB("Select * from TransactionCost", "TransactionCostsTbl")
        GetDataTableFromDB("Select * from EnvironmentVariable", "SettingsTbl")
        GetDataTableFromDB("Select distinct ticker from StockMarket order by ticker", "TickersTbl")
        GetDataTableFromDB("Select distinct Symbol from OptionMarket order by symbol", "SymbolsTbl")

        If turnOffIP Then
            myDataSet.Tables("InitialPositionsTbl").Rows.Clear()
        End If

        lastPriceDownloadDate = "1/1/1"
        initialCAccount = GetInitialCAccount()
        iRate = GetIRate()
        startDate = GetStartDate()
        endDate = GetEndDate()
        maxMargin = GetMaxMargins()

        TPVatStart = CalcTPVAtStart()
        Globals.Dashboard.TeamCellID.Value = "TeamID:" + teamID
        Globals.Dashboard.FillSymbolCBoxes()
        Globals.Dashboard.FillTickerCBoxes()

    End Sub

    Public Sub DownloadTeamData()
        CAccount = DownloadCapitalAccount()
        GetDataTableFromDB("Select * from " + teamPortfolioTableName + " order by symbol", "AcquiredPositionsTbl")
        Globals.Dashboard.TeamCellID.Value = "TeamID:" + teamID
        'CAccount = GetCurrentPosition("CAccount")
        lastTransactionDate = DownloadLastTransactionDate(currentDate)
        lastTEUpdate = Date.Parse("1/1/1")
        sumTE = 0
    End Sub

    Public Sub CalcFinancialMetrics(targetDate As Date)
        IPValue = CalcIPValue(targetDate)
        APValue = CalcAPValue(targetDate)
        margin = CalcMargin(targetDate)
        TPV = IPValue + APValue + CAccount + CalcInterestSLT(currentDate)
        TaTPV = CalcTaTPV(targetDate)
        TE = CalcTE()
        TEpercent = TE / TaTPV
        sumTE = sumTE + UpdateSumTE(targetDate)
    End Sub

    Public Sub DisplayFinancialMetrics(targetDate As Date)

        Globals.Dashboard.Range("F6").Value = CAccount
        Globals.Dashboard.Range("F7").Value = margin
        Globals.Dashboard.Range("F8").Value = margin * 0.3
        Globals.Dashboard.Range("F9").Value = maxMargin
        Globals.Dashboard.Range("F10").Value = ""
        Globals.Dashboard.Range("F11").Value = IPValue
        Globals.Dashboard.Range("F12").Value = APValue
        Globals.Dashboard.Range("F13").Value = TPVatStart
        Globals.Dashboard.Range("F14").Value = TPV
        Globals.Dashboard.Range("F15").Value = TaTPV
        Globals.Dashboard.Range("F16").Value = ""
        Globals.Dashboard.Range("F17").Value = TE
        Globals.Dashboard.Range("F18").Value = TE / TaTPV
        Globals.Dashboard.Range("F19").Value = sumTE

        Globals.Dashboard.AcquiredPositionsLO.DataSource = myDataSet.Tables("AcquiredPositionsTbl")
        Globals.Dashboard.InitialPositionsLO.DataSource = myDataSet.Tables("InitialPositionsTbl")
    End Sub

    Public Function UpdateSumTE(tDate As Date) As Double
        If tDate.DayOfWeek = DayOfWeek.Sunday And tDate > lastTEUpdate Then
            lastTEUpdate = tDate
            Return TE
        Else
            Return 0
        End If
    End Function


    '--- St Part 5


End Module
