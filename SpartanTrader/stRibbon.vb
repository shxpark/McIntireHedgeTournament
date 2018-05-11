Imports Microsoft.Office.Tools.Ribbon

Public Class stRibbon

    Private Sub stRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        'this is where it all starts!
        RibbonUI.ActivateTabMso("TabAddIns")
        Globals.ThisWorkbook.Application.DisplayFormulaBar = False
        DashboardBtn_Click(Nothing, Nothing)
        BetaTBtn_Click(Nothing, Nothing)
    End Sub

    Public Sub DashboardBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles DashboardBtn.Click
        Globals.Dashboard.Activate()
    End Sub

    Public Sub AlphaTBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles AlphaTBtn.Click
        AlphaTBtn.Checked = True
        BetaTBtn.Checked = False
        GammaTBtn.Checked = False
        activeDB = "Alpha"
        ManualTBtn_Click(Nothing, Nothing)
    End Sub

    Public Sub BetaTBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles BetaTBtn.Click
        AlphaTBtn.Checked = False
        BetaTBtn.Checked = True
        GammaTBtn.Checked = False
        activeDB = "Beta"
        ManualTBtn_Click(Nothing, Nothing)
    End Sub

    Public Sub GammaTBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles GammaTBtn.Click
        AlphaTBtn.Checked = False
        BetaTBtn.Checked = False
        GammaTBtn.Checked = True
        activeDB = "Gamma"
        ManualTBtn_Click(Nothing, Nothing)

    End Sub

    Public Sub ManualTBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ManualTBtn.Click
        ManualTBtn.Checked = True
        traderMode = "Manual"
        StartTheTrader()
    End Sub

    Public Sub StockMktBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles StockMktBtn.Click
        GetDataTableFromDB("Select * from StockMarket order by date desc", "StockMarketTbl")
        Globals.Markets.StockMarketLO.DataSource = myDataSet.Tables("StockMarketTbl")
        Globals.Markets.StockMarketLO.Range.Columns.AutoFit()
        Globals.Markets.Activate()
        Globals.Markets.Range("A1").Select()
    End Sub

    Public Sub OptionMktBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles OptionMktBtn.Click
        GetDataTableFromDB("Select * from OptionMarket order by date desc", "OptionMarketTbl")
        Globals.Markets.OptionMarketLO.DataSource = myDataSet.Tables("OptionMarketTbl")
        Globals.Markets.OptionMarketLO.Range.Columns.AutoFit()
        Globals.Markets.Activate()
        Globals.Markets.Range("A1").Select()
    End Sub

    Public Sub SP500Btn_Click(sender As Object, e As RibbonControlEventArgs) Handles SP500Btn.Click
        GetDataTableFromDB("Select * from StockIndex order by date desc", "SP500Tbl")
        Globals.Markets.SP500LO.DataSource = myDataSet.Tables("SP500Tbl")
        Globals.Markets.SP500LO.Range.Columns.AutoFit()
        Globals.Markets.Activate()
        Globals.Markets.Range("A1").Select()
    End Sub

    Public Sub InitialPositionsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles InitialPositionsBtn.Click
        GetDataTableFromDB("Select * from InitialPosition order by symbol", "InitialPositionsTbl")
        Globals.Dashboard.InitialPositionsLO.DataSource = myDataSet.Tables("InitialPositionsTbl")
        'Globals.Dashboard.InitialPositionsLO.Range.Columns.AutoFit()
        'Globals.Dashboard.Range("Q5:S5").Interior.Color = System.Drawing.Color.MidnightBlue
        'Globals.Dashboard.Range("Q5:S5").Font.Color = System.Drawing.Color.White
        'Globals.Dashboard.InitialPositionsLO.Range.NumberFormat = "#,##0;[Red]-#,##0"
        Globals.Dashboard.InitializeDisplay()
        Globals.Dashboard.Activate()
    End Sub

    Public Sub AcquiredPositionsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles AcquiredPositionsBtn.Click
        GetDataTableFromDB("Select * from " + teamPortfolioTableName + " order by symbol", "AcquiredPositionsTbl")
        Globals.Dashboard.AcquiredPositionsLO.DataSource = myDataSet.Tables("AcquiredPositionsTbl")
        'Globals.Dashboard.AcquiredPositionsLO.DataBodyRange.Columns.AutoFit()
        'Globals.Dashboard.Range("U5:W5").Interior.Color = System.Drawing.Color.MidnightBlue
        'Globals.Dashboard.Range("U5:W5").Font.Color = System.Drawing.Color.White
        'Globals.Dashboard.AcquiredPositionsLO.DataBodyRange.NumberFormat = "#,##0;[Red]-#,##0"
        Globals.Dashboard.InitializeDisplay()
        Globals.Dashboard.Activate()
    End Sub

    Public Sub TransactionQBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles TransactionQBtn.Click
        GetDataTableFromDB("Select * from TransactionQueue where teamID = " + teamID + " order by date desc", "TransactionQueueTbl")
        Globals.Transactions.TransactionQueueLO.DataSource = myDataSet.Tables("TransactionQueueTbl")
        Globals.Transactions.TransactionQueueLO.Range.Columns.AutoFit()
        Globals.Transactions.Activate()
        Globals.Transactions.Range("A1").Select()
    End Sub

    Public Sub ConfirmationBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ConfirmationBtn.Click
        GetDataTableFromDB("Select * from " + confirmationTicketTableName + " order by date desc", "ConfirmationTicketTbl")
        Globals.Transactions.ConfirmationTicketsLO.DataSource = myDataSet.Tables("ConfirmationTicketTbl")
        Globals.Transactions.ConfirmationTicketsLO.Range.Columns.AutoFit()
        Globals.Transactions.Activate()
        Globals.Transactions.Range("A1").Select()
    End Sub

    Public Sub QuitBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles QuitBtn.Click
        CloseDBConnection()
        Globals.ThisWorkbook.Application.DisplayAlerts = False
        Globals.ThisWorkbook.Application.DisplayFormulaBar = True
        Globals.ThisWorkbook.Application.Quit()
    End Sub

    Public Sub SettingBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles SettingBtn.Click
        GetDataTableFromDB("Select * from EnvironmentVariable order by Name desc", "SettingTbl")
        Globals.Environment.SettingsLO.DataSource = myDataSet.Tables("SettingTbl")
        Globals.Environment.SettingsLO.Range.Columns.AutoFit()
        Globals.Environment.Activate()
        Globals.Environment.Range("A1").Select()
    End Sub

    Public Sub TCostsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles TCostsBtn.Click
        GetDataTableFromDB("Select * from TransactionCost order by TransactionType desc", "TCostTbl")
        Globals.Environment.TransactionCostLO.DataSource = myDataSet.Tables("TCostTbl")
        Globals.Environment.TransactionCostLO.Range.Columns.AutoFit()
        Globals.Environment.Activate()
        Globals.Environment.Range("A1").Select()
    End Sub

End Class
