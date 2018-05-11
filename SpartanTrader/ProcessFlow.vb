Module ProcessFlow


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

    End Sub

    Public Sub StartTheTrader()
        ClearAllLOS()
        ConnectToActiveDB()
        DownloadStaticData()
        DownloadTeamData()
        Globals.Dashboard.InitializeDisplay()
        currentDate = DownloadCurrentDate()
        RunDailyRoutine()

    End Sub

    Public Sub DownloadStaticData()
        GetDataTableFromDB("Select * from InitialPosition order by symbol", "InitialPositionsTbl")
        GetDataTableFromDB("Select * from TransactionCost", "TransactionCostsTbl")
        GetDataTableFromDB("Select * from EnvironmentVariable", "SettingsTbl")
        GetDataTableFromDB("Select distinct ticker from StockMarket order by ticker", "TickersTbl")
        GetDataTableFromDB("Select distinct Symbol from OptionMarket order by symbol", "SymbolsTbl")

        lastPriceDownloadDate = "1/1/1"
        initialCAccount = GetInitialCAccount()
        iRate = GetIRate()
        startDate = GetStartDate()
        endDate = GetEndDate()
        maxMargin = GetMaxMargins()

        TPVatStart = CalcTPVAtStart()
        Globals.Dashboard.TeamCellID.Value = "TeamID:" + teamID

    End Sub

    Public Sub DownloadTeamData()
        GetDataTableFromDB("Select * from " + teamPortfolioTableName + " order by symbol", "AcquiredPositionsTbl")
        Globals.Dashboard.TeamCellID.Value = "TeamID:" + teamID
    End Sub

    Public Sub RunDailyRoutine()
        Globals.Dashboard.CurrentDateCell.Value = currentDate.ToLongDateString()
        CalcFinancialMetrics(currentDate)
        DisplayFinancialMetrics(currentDate)
    End Sub

    Public Sub CalcFinancialMetrics(targetDate As Date)
        IPValue = CalcIPValue(targetDate)
        APValue = CalcAPValue(targetDate)
    End Sub

    Public Sub DisplayFinancialMetrics(targetDate As Date)

        Globals.Dashboard.Range("F12").Value = IPValue
        Globals.Dashboard.Range("F13").Value = APValue
        Globals.Dashboard.Range("F16").Value = TPVatStart

        Globals.Dashboard.AcquiredPositionsLO.DataSource = myDataSet.Tables("AcquiredPositionsTbl")
        Globals.Dashboard.InitialPositionsLO.DataSource = myDataSet.Tables("InitialPositionsTbl")
    End Sub
End Module
