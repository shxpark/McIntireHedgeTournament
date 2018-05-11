
Public Class FinCharts

    Private Sub Sheet5_Startup() Handles Me.Startup

    End Sub

    Private Sub Sheet5_Shutdown() Handles Me.Shutdown

    End Sub

    Public Sub FillLBoxes()
        TickerLBox.Items.Clear()
        For Each myRow As DataRow In myDataSet.Tables("TickersTbl").Rows
            TickerLBox.Items.Add(myRow("Ticker").ToString().Trim())
        Next
        TickerLBox.Text = "Select Ticker"
        SymbolLBox.Items.Clear()
        For Each myRow As DataRow In myDataSet.Tables("SymbolsTbl").Rows
            SymbolLBox.Items.Add(myRow("Symbol").ToString().Trim())
        Next
    End Sub

    Public Sub SetupFinCharts()
        StockChart.ChartType = Excel.XlChartType.xlLine
        StockDataToChartLO.AutoSetDataBoundColumnHeaders = True

        Dim y As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlValue)
        y.HasTitle = False
        y.HasMinorGridlines = True
        y.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y.TickLabels.NumberFormat = "$###.00"

        Dim x As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlCategory)
        x.HasTitle = False
        x.CategoryType = Excel.XlCategoryType.xlTimeScale
        x.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x.BaseUnit = Excel.XlTimeUnit.xlDays
        x.TickLabels.NumberFormat = "d-mmm"

        OptionChart.ChartType = Excel.XlChartType.xlLine
        OptionDataToChartLO.AutoSetDataBoundColumnHeaders = True

        Dim y2 As Excel.Axis = OptionChart.Axes(Excel.XlAxisType.xlValue)
        y2.HasTitle = False
        y2.HasMinorGridlines = True
        y2.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y2.TickLabels.NumberFormat = "$###.00"

        Dim x2 As Excel.Axis = OptionChart.Axes(Excel.XlAxisType.xlCategory)
        x2.HasTitle = False
        x2.CategoryType = Excel.XlCategoryType.xlTimeScale
        x2.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x2.BaseUnit = Excel.XlTimeUnit.xlDays
        x2.TickLabels.NumberFormat = "d-mmm"

    End Sub

    Private Sub TickerLBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TickerLBox.SelectedIndexChanged
        Dim t As String = ""
        Dim sql As String = ""
        t = TickerLBox.SelectedItem.Trim()
        sql = "Select date, bid, ask from stockmarket where ticker = '" + t + "'"
        GetDataTableFromDB(sql, "StockDataToChartTbl")
        StockDataToChartLO.DataSource = myDataSet.Tables("StockDataToChartTbl")
        StockChart.SetSourceData(StockDataToChartLO.Range)
        StockChart.ChartTitle.Text = "Daily Closings for " + TickerLBox.SelectedItem.ToString()
        Dim y As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlValue)
        y.MinimumScale = Math.Truncate(FindMinBid("StockDataToChartTbl") / 10) * 10
    End Sub

    Private Sub SymbolLBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SymbolLBox.SelectedIndexChanged
        Dim s As String = ""
        Dim sql As String = ""
        s = SymbolLBox.SelectedItem.Trim()
        sql = "Select date, bid, ask from optionmarket where symbol = '" + s + "'"
        GetDataTableFromDB(sql, "OptionDataToChartTbl")
        OptionDataToChartLO.DataSource = myDataSet.Tables("OptionDataToChartTbl")
        OptionChart.SetSourceData(OptionDataToChartLO.Range)
        OptionChart.ChartTitle.Text = "Daily Closings for " + SymbolLBox.SelectedItem.ToString()
        Dim y As Excel.Axis = OptionChart.Axes(Excel.XlAxisType.xlValue)
        y.MinimumScale = Math.Truncate(FindMinBid("OptionDataToChartTbl") / 10) * 10
    End Sub

    Public Function FindMinBid(tableName As String) As Double
        Dim tempMin As Double = 100000000
        For Each myRow As DataRow In myDataSet.Tables(tableName).Rows
            tempMin = Math.Min(myRow("Bid"), tempMin)
        Next
        Return tempMin
    End Function
End Class
