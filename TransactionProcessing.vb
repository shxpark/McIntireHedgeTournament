Module TransactionProcessing

    Public Sub ClearTransaction()
        Globals.Dashboard.Range("C6:C19").ClearContents()
        Globals.Dashboard.Range("C6:C8").Font.Color = System.Drawing.Color.White
        trType = ""
        trSecurityType = ""
    End Sub

    Public Sub ComputeTransactionProperties()
        Select Case trType
            Case "Buy"
                trPrice = GetAsk(trSymbol, currentDate)
            Case "Sell"
                trPrice = GetBid(trSymbol, currentDate)
            Case "SellShort"
                trPrice = GetBid(trSymbol, currentDate)
            Case "CashDiv"
                trPrice = GetDividend(trSymbol, currentDate)
            Case "X-Call"
                trPrice = GetStrike(trSymbol)
            Case "X-Put"
                trPrice = GetStrike(trSymbol)
            Case Else
                MessageBox.Show("Unknown transaction type.",
                                "Unkown trType", MessageBoxButtons.OK, MessageBoxIcon.Error)
                trPrice = 0
        End Select

        If IsAStock(trSymbol) Then
            trSecurityType = "Stock"
        Else
            trSecurityType = "Option"
        End If


        If trSecurityType = "Option" Then
            trStrike = GetStrike(trSymbol)
        Else
            trStrike = 0
        End If
        trCost = CalcTransactionCost()
        trTotValue = CalcTotValue()
        interestSLT = CalcInterestSLT(currentDate)
        CAccountAT = CAccount + trTotValue + interestSLT
        marginAT = margin + EffectOfTransactionOnMargin(trType, trSymbol, trQty, currentDate)
        trDelta = CalcDelta(trSymbol, currentDate)
    End Sub

    Public Function CalcTransactionCost() As Double
        Return GetTrCostCoefficient(trSecurityType, trType) * Math.Abs(trQty) * trPrice
    End Function

    Public Function CalcTotValue() As Double
        Select Case trType
            Case "Buy"
                Return -(trPrice * trQty) - trCost
            Case "Sell"
                Return (trPrice * trQty) - trCost
            Case "SellShort"
                Return (trPrice * trQty) - trCost
            Case "CashDivd"
                Return (trPrice * trQty) - trCost
            Case "X-Put"
                Return (trPrice * trQty) - trCost
            Case "X-Call"
                Return -(trPrice * trQty) - trCost
            Case Else
                Return 0
        End Select
    End Function


    Public Sub ExecuteTransaction()
        Dim mySQL As String
        Dim trans As String
        mySQL = String.Format("INSERT INTO TransactionQueue (Date, TeamID, Symbol, Type, Qty, Price, Cost, TotValue, " +
                              "InterestSinceLastTransaction, CashPositionAfterTransaction, TotMargin) VALUES " +
                              "('{0}', {1}, '{2}', '{3}', {4}, {5}, {6}, {7}, {8}, {9}, {10})",
                              currentDate.ToShortDateString,
                              teamID,
                              trSymbol,
                              trType,
                              trQty,
                              trPrice,
                              trCost,
                              trTotValue,
                              interestSLT,
                              CAccountAT,
                              marginAT)

        RunQuery(mySQL)
        lastTransactionDate = currentDate
        CAccount = CAccountAT
        margin = marginAT
        UpdatePosition(trType, trSymbol, trQty)
        trans = String.Format("{0}> {1} {2} {3} for ${4})",
                                currentDate.DayOfWeek, trType, trQty.ToString("N0"), trSymbol, trTotValue.ToString("N0"))
        Globals.Dashboard.TransactionsTB.Text = trans + vbCrLf + Globals.Dashboard.TransactionsTB.Text

    End Sub

    Public Sub HighlightTransaction()
        Globals.Dashboard.Range("C6:C8").Font.Color = System.Drawing.Color.Lime
        Globals.Dashboard.Range("J6:J17").Font.Color = System.Drawing.Color.Lime
    End Sub

    Public Function EffectOfTransactionOnMargin(transactionType As String, sym As String, q As Double, targetDate As Date) As Double
        ' The code is provided as a link from the homework.
        ' Make sure that you understand the financial logic.
        ' This function finds the net effect of the current transaction on your overall margin
        Dim currPos As Integer = 0
        Dim underlierPos As Integer = 0
        Dim effectOnMargin As Double = 0
        Dim underlier As String = ""

        Select Case trType
            Case "Sell"
                '  Sell has no effect on margin because you can only sell what you have long
                Return 0 ' effect of transaction on margin

            Case "Buy"
                currPos = GetCurrentPositionInAP(sym)
                If currPos >= 0 Then
                    Return 0
                Else
                    If q >= Math.Abs(currPos) Then
                        Return -currPos * CalcMTM(trSymbol, targetDate)
                        ' buying eliminates all margin for this symbol
                    Else
                        Return (q * CalcMTM(sym, targetDate))
                        ' buying reduces the margin
                    End If
                End If

            Case "SellShort"
                Return -q * CalcMTM(sym, targetDate)
                ' Selling short is easy: it always increases the margin

            Case "CashDiv"
                Return 0

            Case "X-Call"
                ' Two effects on margin: the change in options and the change in stocks
                Dim OptionEffect As Double = 0
                currPos = GetCurrentPositionInAP(sym)   '  here currPosition is the option position
                underlier = GetUnderlier(sym)
                underlierPos = GetCurrentPositionInAP(underlier)

                ' first the effect of exercising the call on the call position
                If currPos < 0 Then
                    OptionEffect = q * CalcMTM(sym, targetDate)
                    ' i.e., it reduces the margin
                Else
                    OptionEffect = 0
                End If

                ' next, the effect of the called stock
                ' two cases: long and short calls.
                If currPos >= 0 Then      ' long call is like buying
                    If underlierPos >= 0 Then
                        Return OptionEffect
                    Else
                        If q >= Math.Abs(underlierPos) Then
                            Return OptionEffect - (underlierPos * CalcMTM(underlier, targetDate))
                            ' X-call eliminates all margin for this symbol
                        Else
                            Return OptionEffect + (q * CalcMTM(underlier, targetDate))
                            ' X-call reduces the margin
                        End If
                    End If

                Else      '  short call is like selling

                    If underlierPos <= 0 Then
                        Return OptionEffect - (q * CalcMTM(underlier, targetDate))
                    Else   ' underlier positive
                        If underlierPos >= q Then
                            Return OptionEffect
                        Else
                            Return OptionEffect - ((q - underlierPos) * CalcMTM(underlier, targetDate))
                        End If
                    End If
                End If

            Case "X-Put"
                ' Two effects on margin: the change in options and the change in stocks
                Dim OptionEffect As Double = 0
                currPos = GetCurrentPositionInAP(sym)   '  here currPosition is the option position
                underlier = GetUnderlier(sym)
                underlierPos = GetCurrentPositionInAP(underlier)

                ' first the effect of exercising the option on the put position
                If currPos < 0 Then
                    OptionEffect = q * CalcMTM(sym, targetDate)
                    ' and it reduces the margin
                Else
                    OptionEffect = 0
                End If

                ' next, the effect of the put stock
                ' two cases: long and short puts
                If currPos < 0 Then      ' short put is like buying

                    If underlierPos >= 0 Then
                        Return OptionEffect
                    Else
                        If q >= Math.Abs(underlierPos) Then
                            Return OptionEffect - (underlierPos * CalcMTM(underlier, targetDate))
                            ' X-put eliminates all margin for this symbol
                        Else
                            Return OptionEffect + (q * CalcMTM(underlier, targetDate))
                            ' X-put reduces the margin
                        End If
                    End If

                Else      ' long put is like selling

                    If underlierPos <= 0 Then
                        Return OptionEffect - (q * CalcMTM(underlier, targetDate))
                    Else   ' underlier positive
                        If underlierPos >= q Then
                            Return OptionEffect
                        Else
                            Return OptionEffect - ((q - underlierPos) * CalcMTM(underlier, targetDate))
                        End If
                    End If
                End If
        End Select
        MessageBox.Show("Holy BatBeans! I could not figure out the impact of " + sym + " on margin.  I returned $0.")
        Return 0
    End Function

    Public Sub DisplayTransactionProperties()
        Globals.Dashboard.Range("C6").Value = trType
        Globals.Dashboard.Range("C7").Value = trQty
        Globals.Dashboard.Range("C8").Value = trSymbol

        Globals.Dashboard.Range("C10").Value = trStrike
        Globals.Dashboard.Range("C11").Value = trDelta

        Globals.Dashboard.Range("C13").Value = trPrice
        Globals.Dashboard.Range("C14").Value = trCost
        Globals.Dashboard.Range("C15").Value = trTotValue
        Globals.Dashboard.Range("C17").Value = CAccount
        Globals.Dashboard.Range("C18").Value = interestSLT
        Globals.Dashboard.Range("C19").Value = marginAT

    End Sub



End Module
