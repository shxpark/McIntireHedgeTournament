Module ScheduledTransactions
    Public Sub DoScheduledTransactions(targetDate As Date)
        Select Case targetDate.ToShortDateString()
            Case "5/10/2017"
                Dim qty = GetCurrentPosition("XOM")
                ExecuteScheduledTransaction("CashDiv", "XOM", qty, targetDate)
            Case "5/30/2017"
                Dim qty = GetCurrentPosition("TEVA")
                ExecuteScheduledTransaction("CashDiv", "TEVA", qty, targetDate)
            Case "6/1/2017"
                Dim qty = GetCurrentPosition("NKE")
                ExecuteScheduledTransaction("CashDiv", "NKE", qty, targetDate)
            Case "6/14/2017"
                Dim qty = GetCurrentPosition("RGR")
                ExecuteScheduledTransaction("CashDiv", "RGR", qty, targetDate)
        End Select
    End Sub

    Public Sub ExecuteScheduledTransaction(type As String, sym As String, qty As Double, tDate As Date)
        trType = type
        trSymbol = sym
        trQty = qty
        ComputeTransactionProperties()
        If IsTransactionValid(trType, trSymbol, trQty) Then
            ExecuteTransaction()
            CalcFinancialMetrics(currentDate)
            DisplayFinancialMetrics(currentDate)
            CalcRecommendations(tDate)
            DisplayRecommendations()
        End If
    End Sub
End Module
