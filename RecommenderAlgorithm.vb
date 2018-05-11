Module RecommenderAlgorithm
    Public Sub ResetRecommendations()
        Dim myRow As DataRow
        If myDataSet.Tables.Contains("RecommendationsTbl") Then
            myDataSet.Tables("RecommendationsTbl").Clear()
        Else
            myDataSet.Tables.Add("RecommendationsTbl")
            myDataSet.Tables("RecommendationsTbl").Columns.Add("Underlier", GetType(String))
            myDataSet.Tables("RecommendationsTbl").Columns.Add("Vol", GetType(Double))
            myDataSet.Tables("RecommendationsTbl").Columns.Add("FamDelta", GetType(Double))
            myDataSet.Tables("RecommendationsTbl").Columns.Add("RecType", GetType(String))
            myDataSet.Tables("RecommendationsTbl").Columns.Add("RecSymbol", GetType(String))
            myDataSet.Tables("RecommendationsTbl").Columns.Add("RecQty", GetType(Double))
        End If
        For i = 0 To 11
            myRow = myDataSet.Tables("RecommendationsTbl").Rows.Add()
            myRow("Underlier") = myDataSet.Tables("TickersTbl").Rows(i)("Ticker")
            myRow("Vol") = GetVol(myRow("Underlier"))
            myRow("FamDelta") = 0
            myRow("RecType") = "Hold"
            myRow("RecSymbol") = "-"
            myRow("RecQty") = 0
        Next
        'DisplayRecommendations()
    End Sub

    Public Sub DisplayRecommendations()
        Dim myRow As DataRow
        For i = 0 To 11
            myRow = myDataSet.Tables("RecommendationsTbl").Rows(i)
            Globals.Dashboard.UnderlierRange.Rows(i + 1) = myRow("Underlier")
            Globals.Dashboard.VolatilityRange.Rows(i + 1) = myRow("Vol")
            Globals.Dashboard.FamilyDeltaRange.Rows(i + 1) = myRow("FamDelta")
            Globals.Dashboard.RecTypeRange.Rows(i + 1) = myRow("RecType")
            Globals.Dashboard.RecSymbolRange.Rows(i + 1) = myRow("RecSymbol")
            Globals.Dashboard.RecQtyRange.Rows(i + 1) = myRow("RecQty")
        Next
    End Sub

    Public Sub CalcRecommendations(targetDate As Date)
        Dim ticker As String
        For Each myRecRow As DataRow In myDataSet.Tables("RecommendationsTbl").Rows
            ticker = myRecRow("Underlier").trim()
            myRecRow("FamDelta") = CalcFamilyDelta(ticker, targetDate)
            myRecRow("RecType") = "Hold"
            myRecRow("RecSymbol") = ""
            myRecRow("RecQty") = 0
            If HedgingToday(myRecRow, targetDate) = True AndAlso NeedToHedge(myRecRow, targetDate) = True Then
                CalcCandidateRecScores(myRecRow, targetDate)
                FindBestHedgeForFamily(myRecRow, targetDate)
            End If
        Next
    End Sub

    Public Sub SmartHedgeAll()
        For i As Integer = 0 To 11
            CalcRecommendation(i, currentDate)
            trType = myDataSet.Tables("RecommendationsTbl").Rows(i)("RecType").Trim()
            trQty = myDataSet.Tables("RecommendationsTbl").Rows(i)("RecQty")
            trSymbol = myDataSet.Tables("RecommendationsTbl").Rows(i)("RecSymbol").Trim()
            If trType = "Hold" Then
            Else
                ComputeTransactionProperties()
                If IsTransactionValid(trType, trSymbol, trQty) Then
                    ExecuteTransaction()
                    CalcFinancialMetrics(currentDate)
                    If traderMode <> "Sim" Then
                        Application.DoEvents()
                        DisplayFinancialMetrics(currentDate)
                        CalcRecommendations(currentDate)
                        DisplayRecommendations()
                        Application.DoEvents()
                    End If
                Else
                    Dim msg As String
                    msg = String.Format("I could not execute '{0}' '{1}' {2}'.", trType, trQty, trSymbol)
                    MessageBox.Show(msg)
                End If
            End If
        Next
    End Sub



    Public Sub CalcCandidateRecScores(myRecRow As DataRow, targetDate As Date)
        ResetCandidateRecommendations()
        If myRecRow("FamDelta") > 0 Then
            EvaluateSellingStock(800, myRecRow, targetDate)
            EvaluateSellingCall(700, myRecRow, targetDate)
            EvaluateBuyingPut(600, myRecRow, targetDate)
            EvaluateBuyingBackPut(400, myRecRow, targetDate)
            EvaluateSellingShortStock(100, myRecRow, targetDate)
            EvaluateSellingShortCall(200, myRecRow, targetDate)
        Else
            EvaluateSellingPut(800, myRecRow, targetDate)
            EvaluateBuyingBackStock(700, myRecRow, targetDate)
            EvaluateBuyingCall(600, myRecRow, targetDate)
            EvaluateBuyingBackCall(500, myRecRow, targetDate)
            EvaluateSellingShortPut(300, myRecRow, targetDate)
            EvaluateBuyingStock(100, myRecRow, targetDate)
        End If
    End Sub

    Public Sub ResetCandidateRecommendations()
        If myDataSet.Tables.Contains("CandidateRecommendationsTbl") Then
            myDataSet.Tables("CandidateRecommendationsTbl").Clear()
        Else
            myDataSet.Tables.Add("CandidateRecommendationsTbl")
            myDataSet.Tables("CandidateRecommendationsTbl").Columns.Add("Symbol", GetType(String))
            myDataSet.Tables("CandidateRecommendationsTbl").Columns.Add("Type", GetType(String))
            myDataSet.Tables("CandidateRecommendationsTbl").Columns.Add("Qty", GetType(Double))
            myDataSet.Tables("CandidateRecommendationsTbl").Columns.Add("Score", GetType(Double))
        End If
    End Sub

    Public Function HedgingToday(myRecRow As DataRow, targetDate As Date) As Boolean
        If targetDate.DayOfWeek = DayOfWeek.Saturday Or
                targetDate.DayOfWeek = DayOfWeek.Sunday Then
            Return False
        End If
        Return True
    End Function

    Public Function NeedToHedge(RecRow As DataRow, targetDate As Date) As Boolean
        If Math.Abs(RecRow("FamDelta")) < 1000 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function TooCloseToMaxMargins() As Boolean
        If ((maxMargin - Math.Abs(margin)) < 2000000) Then
            Return True
        Else
            Return False

        End If
    End Function

    Public Function MaxShortWithinContraints(sym As String, tdate As Date) As Double
        Dim q As Double = 0
        Dim maxAllowableIncreaseInMargins As Double = 0
        If TooCloseToMaxMargins() = True Then
            Return 0
        Else
            maxAllowableIncreaseInMargins = (maxMargin - Math.Abs(margin)) - 1000000
            If maxAllowableIncreaseInMargins <= 0 Then
                Return 0
            Else
                q = maxAllowableIncreaseInMargins / GetBid(sym, tdate)
                Return Math.Truncate(q)
            End If
        End If
    End Function

    Public Function AvailableCashInLow() As Boolean
        Dim availableCash As Double = CAccount - (Math.Abs(margin) * 0.3)
        If availableCash < 1000000 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function MaxPurchasePossible(sym As String, tdate As Date) As Double
        Dim ask As Double = 0
        Dim q As Double = 0
        Dim availableCash As Double = CAccount - (Math.Abs(margin) * 0.3)
        availableCash = availableCash - 1000000
        ask = GetAsk(sym, tdate)
        If availableCash > 0 And ask > 0 Then
            q = availableCash / ask
            Return Math.Truncate(q)
        Else
            Return 0
        End If
    End Function

    Public Sub FindBestHedgeForFamily(myRecRow As DataRow, targdate As Date)
        myRecRow("recType") = "Hold"
        myRecRow("recQty") = 0
        myRecRow("recSymbol") = ""
        Dim bestscore As Double = 0

        If myDataSet.Tables("CandidateRecommendationsTbl").Rows.Count = 0 Then
            Exit Sub
        End If
        For Each myCandidateRecRow As DataRow In myDataSet.Tables("CandidateRecommendationsTbl").Rows
            If myCandidateRecRow("Score") > bestscore Then
                myRecRow("recType") = myCandidateRecRow("Type")
                myRecRow("recQty") = myCandidateRecRow("Qty")
                myRecRow("recSymbol") = myCandidateRecRow("Symbol")
                bestscore = myCandidateRecRow("Score")
            End If
        Next
    End Sub

    Public Sub CalcRecommendation(i As Integer, targetDate As Date)
        Dim ticker As String
        Dim myRecRow As DataRow
        myRecRow = myDataSet.Tables("RecommendationsTbl").Rows(i)
        ticker = myRecRow("Underlier").trim()
        myRecRow("FamDelta") = CalcFamilyDelta(ticker, targetDate)
        myRecRow("RecType") = "Hold"
        myRecRow("RecSymbol") = ""
        myRecRow("RecQty") = 0
        If HedgingToday(myRecRow, targetDate) = True AndAlso NeedToHedge(myRecRow, targetDate) = True Then
            CalcCandidateRecScores(myRecRow, targetDate)
            FindBestHedgeForFamily(myRecRow, targetDate)
        End If
        For j As Integer = 0 To 20
            Application.DoEvents()
        Next
    End Sub


End Module
