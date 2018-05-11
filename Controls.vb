Module Controls

    Public Function IsStockInputValid() As Boolean
        If Globals.Dashboard.TickerCBox.SelectedItem = Nothing Then
            MessageBox.Show("Pick a stock.",
                            "No Ticker", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Else
            trSymbol = Globals.Dashboard.TickerCBox.SelectedItem
        End If

        If trType = "" Then
            MessageBox.Show("To buy or not to buy?",
                            "No transaction type", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        Try
            trQty = Integer.Parse(Globals.Dashboard.StockQtyTBox.Text)
        Catch
            MessageBox.Show("Quantity?",
                            "No quantity", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
        Return True
    End Function

    Public Function IsOptionInputValid() As Boolean
        If Globals.Dashboard.SymbolCBox.SelectedItem = Nothing Then
            MessageBox.Show("Pick an Option.",
                            "No Ticker", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Else
            trSymbol = Globals.Dashboard.SymbolCBox.SelectedItem
        End If

        If trType = "" Then
            MessageBox.Show("To buy or not to buy?",
                            "No transaction type", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        Try
            trQty = Integer.Parse(Globals.Dashboard.OptionQtyTBox.Text)
        Catch
            MessageBox.Show("Quantity?",
                            "No quantity", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
        Return True
    End Function


    Public Function IsTransactionValid(transType As String, transSym As String, transQty As Double) As Boolean
        If IsInIP(transSym) And (transType = "Buy" Or transType = "Sell" Or transType = "SellShort") Then
            MessageBox.Show("Holy Cow! You cannot trade securities in IP. Not sent.", "Accounting controls",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        If (currentDate.DayOfWeek = DayOfWeek.Saturday Or currentDate.DayOfWeek = DayOfWeek.Sunday) And
                (transType = "Buy" Or transType = "Sell" Or transType = "SellShort" Or transType = "CashDiv") Then
            MessageBox.Show("Holy Cow! Weekend not sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        If trQty = 0 Then
            MessageBox.Show("Holy Cow! Zero quantity. Not sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        Return True
    End Function

    Public Function IsApEntryValid(sym As String, unit As String) As Boolean
        If sym = "" Or unit = "" Then
            Return False
        End If
        sym = sym.Trim()
        If Not IsNumeric(unit) Then
            Return False
        End If
        If Double.Parse(unit) = 0 And sym <> "CAccount" Then
            Return False
        End If
        If Not (IsAStock(sym) Or IsAnOption(sym) Or sym = "CAccount") Then
            MessageBox.Show("I am afraid I cannot process this. Unknown security (" + sym + ")", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        Return True
    End Function
End Module
