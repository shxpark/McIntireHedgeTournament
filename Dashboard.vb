
Public Class Dashboard

    Private Sub Sheet4_Startup() Handles Me.Startup

    End Sub

    Private Sub Sheet4_Shutdown() Handles Me.Shutdown

    End Sub

    Public Sub InitializeDisplay()
        InitialPositionsLO.DataBodyRange.Interior.Color = System.Drawing.Color.MidnightBlue
        InitialPositionsLO.ListColumns(1).Range.Font.Color = System.Drawing.Color.White
        InitialPositionsLO.ListColumns(2).Range.NumberFormat = "[White]#,##0;[Red]-#,##0"
        InitialPositionsLO.ListColumns(3).Range.NumberFormat = "[Green]$ #,##0;[Red]$ -#,##0;-"
        InitialPositionsLO.Range.Columns.ColumnWidth = 13

        AcquiredPositionsLO.DataBodyRange.Interior.Color = System.Drawing.Color.MidnightBlue
        AcquiredPositionsLO.ListColumns(1).Range.Font.Color = System.Drawing.Color.White
        AcquiredPositionsLO.ListColumns(2).Range.NumberFormat = "[White]#,##0;[Red]-#,##0"
        AcquiredPositionsLO.ListColumns(3).Range.NumberFormat = "[Green]$ #,##0;[Red]$ -#,##0;-"
        AcquiredPositionsLO.Range.Columns.ColumnWidth = 13

        ResetRecommendations()
        Globals.Dashboard.SetupTETracker()
        TransactionsTB.Text = "------ Ready ------"
    End Sub



    Private Sub BuyStockBtn_Click(sender As Object, e As EventArgs) Handles BuyStockBtn.Click
        ClearTransaction()
        trType = "Buy"
        trSecurityType = "Stock"
        If IsStockInputValid() = True Then
            ComputeTransactionProperties()
            DisplayTransactionProperties()
        End If
    End Sub

    Private Sub SellStockBtn_Click(sender As Object, e As EventArgs) Handles SellStockBtn.Click
        ClearTransaction()
        trType = "Sell"
        trSecurityType = "Stock"
        If IsStockInputValid() = True Then
            ComputeTransactionProperties()
            DisplayTransactionProperties()
        End If
    End Sub

    Private Sub SellShortStockBtn_Click(sender As Object, e As EventArgs) Handles SellShortStockBtn.Click
        ClearTransaction()
        trType = "SellShort"
        trSecurityType = "Stock"
        If IsStockInputValid() = True Then
            ComputeTransactionProperties()
            DisplayTransactionProperties()
        End If
    End Sub

    Private Sub CashDivBtn_Click(sender As Object, e As EventArgs) Handles CashDivBtn.Click
        ClearTransaction()
        trType = "CashDiv"
        trSecurityType = "Stock"
        If IsStockInputValid() = True Then
            ComputeTransactionProperties()
            DisplayTransactionProperties()
        End If
    End Sub

    Private Sub ExecStockTransactionBtn_Click(sender As Object, e As EventArgs) Handles ExecStockTransactionBtn.Click
        If IsStockInputValid() = True Then
            ComputeTransactionProperties()
            If IsTransactionValid(trType, trSymbol, trQty) Then
                ExecuteTransaction()
                HighlightTransaction()
                CalcFinancialMetrics(currentDate)
                DisplayFinancialMetrics(currentDate)
                CalcRecommendations(currentDate)
                DisplayRecommendations()
            End If
        End If
    End Sub

    '---Start part 4---
    Public Sub FillTickerCBoxes()
        TickerCBox.Items.Clear()
        For Each myRow As DataRow In myDataSet.Tables("TickersTbl").Rows
            TickerCBox.Items.Add(myRow("Ticker").ToString().Trim())
        Next
        TickerCBox.Text = "Select Ticker"
    End Sub

    Public Sub FillSymbolCBoxes()
        SymbolCBox.Items.Clear()
        For Each myRow As DataRow In myDataSet.Tables("SymbolsTbl").Rows
            SymbolCBox.Items.Add(myRow("Symbol").ToString().Trim())
        Next
        SymbolCBox.Text = "Select Symbol"
    End Sub


    Private Sub BuyOptionBtn_Click(sender As Object, e As EventArgs) Handles BuyOptionBtn.Click
        ClearTransaction()
        trType = "Buy"
        trSecurityType = "Option"
        If IsOptionInputValid() = True Then
            ComputeTransactionProperties()
            DisplayTransactionProperties()
        End If
    End Sub

    Private Sub SellOptionBtn_Click(sender As Object, e As EventArgs) Handles SellOptionBtn.Click
        ClearTransaction()
        trType = "Sell"
        trSecurityType = "Option"
        If IsOptionInputValid() = True Then
            ComputeTransactionProperties()
            DisplayTransactionProperties()
        End If
    End Sub

    Private Sub SellShortOptionBtn_Click(sender As Object, e As EventArgs) Handles SellShortOptionBtn.Click
        ClearTransaction()
        trType = "SellShort"
        trSecurityType = "Option"
        If IsOptionInputValid() = True Then
            ComputeTransactionProperties()
            DisplayTransactionProperties()
        End If
    End Sub

    Private Sub XOptionBtn_Click(sender As Object, e As EventArgs) Handles XOptionBtn.Click
        ClearTransaction()
        If SymbolCBox.SelectedItem <> Nothing Then
            If IsACall(SymbolCBox.SelectedItem) Then
                trType = "X-Call"
            End If
            If IsAPut(SymbolCBox.SelectedItem) Then
                trType = "X-Put"
            End If
        End If
        trSecurityType = "Option"
        If IsOptionInputValid() = True Then
            ComputeTransactionProperties()
            DisplayTransactionProperties()
        End If
    End Sub

    Private Sub ExecOptionTransactionBtn_Click(sender As Object, e As EventArgs) Handles ExecOptionTransactionBtn.Click
        If IsOptionInputValid() = True Then
            ComputeTransactionProperties()
            If IsTransactionValid(trType, trSymbol, trQty) Then
                ExecuteTransaction()
                HighlightTransaction()
                CalcFinancialMetrics(currentDate)
                DisplayFinancialMetrics(currentDate)
                CalcRecommendations(currentDate)
                DisplayRecommendations()
            End If
        End If
    End Sub

    Public Sub ExecuteRecommendation(ButtonNumber As Integer)
        Globals.Dashboard.RecTypeRange.Value = "Hold"
        trType = myDataSet.Tables("RecommendationsTbl").Rows(ButtonNumber)("RecType").Trim()
        trQty = myDataSet.Tables("RecommendationsTbl").Rows(ButtonNumber)("RecQty")
        trSymbol = myDataSet.Tables("RecommendationsTbl").Rows(ButtonNumber)("RecSymbol").Trim()
        If trType = "Hold" Then
            Exit Sub
        End If
        ComputeTransactionProperties()
        If IsTransactionValid(trType, trSymbol, trQty) Then
            ExecuteTransaction()
            DisplayTransactionProperties()
            HighlightTransaction()
            CalcFinancialMetrics(currentDate)
            DisplayFinancialMetrics(currentDate)
            CalcRecommendations(currentDate)
            DisplayRecommendations()

        Else
            Dim msg As String
            msg = String.Format("Holy Batbeans! I could not execute '{0}' '{1}' '{2}'.", trType, trQty, trSymbol)
            MessageBox.Show(msg)
        End If



    End Sub



    Private Sub Rec1Btn_Click(sender As Object, e As EventArgs) Handles Rec1Btn.Click
        ExecuteRecommendation(0)
    End Sub

    Private Sub Rec2Btn_Click(sender As Object, e As EventArgs) Handles Rec2Btn.Click
        ExecuteRecommendation(1)
    End Sub

    Private Sub Rec3Btn_Click(sender As Object, e As EventArgs) Handles Rec3Btn.Click
        ExecuteRecommendation(2)
    End Sub

    Private Sub Rec4Btn_Click(sender As Object, e As EventArgs) Handles Rec4Btn.Click
        ExecuteRecommendation(3)
    End Sub

    Private Sub Rec5Btn_Click(sender As Object, e As EventArgs) Handles Rec5Btn.Click
        ExecuteRecommendation(4)
    End Sub

    Private Sub Rec6Btn_Click(sender As Object, e As EventArgs) Handles Rec6Btn.Click
        ExecuteRecommendation(5)
    End Sub

    Private Sub Rec7Btn_Click(sender As Object, e As EventArgs) Handles Rec7Btn.Click
        ExecuteRecommendation(6)
    End Sub

    Private Sub Rec8Btn_Click(sender As Object, e As EventArgs) Handles Rec8Btn.Click
        ExecuteRecommendation(7)
    End Sub

    Private Sub Rec9Btn_Click(sender As Object, e As EventArgs) Handles Rec9Btn.Click
        ExecuteRecommendation(8)
    End Sub

    Private Sub Rec10Btn_Click(sender As Object, e As EventArgs) Handles Rec10Btn.Click
        ExecuteRecommendation(9)
    End Sub

    Private Sub Rec11Btn_Click(sender As Object, e As EventArgs) Handles Rec11Btn.Click
        ExecuteRecommendation(10)
    End Sub

    Private Sub Rec12Btn_Click(sender As Object, e As EventArgs) Handles Rec12Btn.Click
        ExecuteRecommendation(11)
    End Sub

    Public Sub SetupTETracker()
        If myDataSet.Tables.Contains("TETbl") Then
            myDataSet.Tables("TETbl").Clear()
        Else
            myDataSet.Tables.Add("TETbl")
            myDataSet.Tables("TETbl").Columns.Add("Date", GetType(Date))
            myDataSet.Tables("TETbl").Columns.Add("TaTPV", GetType(Double))
            myDataSet.Tables("TETbl").Columns.Add("NoHedge", GetType(Double))
            myDataSet.Tables("TETbl").Columns.Add("TPV", GetType(Double))
        End If
        TELO.DataSource = myDataSet.Tables("TETbl")

        TEChart.ChartType = Excel.XlChartType.xlLine
        TEChart.HasTitle = False
        TEChart.HasLegend = True

        Dim y As Excel.Axis = TEChart.Axes(Excel.XlAxisType.xlValue)
        y.HasTitle = False
        y.HasMinorGridlines = True
        y.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y.TickLabels.NumberFormat = "$#,###"
        y.MinimumScaleIsAuto = False
        y.MaximumScaleIsAuto = True

        Dim x As Excel.Axis = TEChart.Axes(Excel.XlAxisType.xlCategory)
        x.CategoryType = Excel.XlCategoryType.xlTimeScale
        x.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x.BaseUnit = Excel.XlTimeUnit.xlDays
        x.TickLabels.NumberFormat = "[$-409]d-mmm;@"

        Dim s As Excel.SeriesCollection = TEChart.SeriesCollection
        s(0).Format.Line.Weight = 2
        s(0).Format.Line.ForeColor.RGB = System.Drawing.Color.SteelBlue
        s(1).Format.Line.Weight = 2
        s(1).Format.Line.ForeColor.RGB = System.Drawing.Color.Gray
        s(2).Format.Line.Weight = 2
        s(2).Format.Line.ForeColor.RGB = System.Drawing.Color.Orange

    End Sub

    Public Sub UpdateTEChart(targetDate As Date)
        Dim interestOnInitialCA As Double = 0
        Dim ts As TimeSpan = targetDate.Date - startDate.Date
        Dim t As Double = ts.Days / 365.25

        interestOnInitialCA = initialCAccount * (Math.Exp(iRate * t) - 1)

        Dim tempRow As DataRow
        tempRow = myDataSet.Tables("TETbl").Rows.Add()
        tempRow("Date") = targetDate.ToShortDateString
        tempRow("TPV") = TPV
        tempRow("TaTPV") = TaTPV
        tempRow("NoHedge") = IPValue + initialCAccount + interestOnInitialCA

        TEChart.SetSourceData(TELO.Range)

        Dim y As Excel.Axis = TEChart.Axes(Excel.XlAxisType.xlValue)
        y.MinimumScale = Math.Truncate((FindMinInTPVTrackingTable() / 10000000)) * 10000000
    End Sub

    Public Function FindMinInTPVTrackingTable() As Double
        Dim tempMin As Double = 100000000
        For Each myRow As DataRow In myDataSet.Tables("TETbl").Rows
            tempMin = Math.Min(myRow("TPV"), tempMin)
            tempMin = Math.Min(myRow("TaTPV"), tempMin)
            tempMin = Math.Min(myRow("NoHedge"), tempMin)
        Next
        Return tempMin
    End Function


End Class
