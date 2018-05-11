Module ScoringRules

    '  NOTE: this code is given to you As-Is. It executes, but its parameters 
    '  are intentionally set to suboptimal values.  You need to understand it
    '  and make all the necessary changes.

    Public Sub EvaluateSellingStock(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyDelta = recRow("FamDelta")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim adjustment As Double = 0
        If IsInIP(underlier) Then
            Exit Sub   ' cannot sell if in IP
        Else
            underlierCurrPos = GetCurrentPositionInAP(underlier)
            If underlierCurrPos <= 0 Then ' we cannot sell since we are not long
                Exit Sub
            Else
                hedgeQty = CalcQtyNeededToHedge(underlier, familyDelta, tDate)
                If hedgeQty = 0 Then
                    Exit Sub  ' nothing to do
                Else
                    adjustment = (1 - (underlierCurrPos / hedgeQty)) * -50
                    If hedgeQty > underlierCurrPos Then     ' you have fewer than needed
                        hedgeQty = underlierCurrPos         ' sell all you have
                        'adjustment = -50                    ' arbitrary adjustment!
                    End If
                    newRow = myDataSet.Tables("CandidateRecommendationsTbl").Rows.Add()
                    newRow("Type") = "Sell"
                    newRow("Symbol") = underlier
                    newRow("Qty") = hedgeQty
                    newRow("Score") = baseScore + adjustment
                End If
            End If
        End If
    End Sub

    Public Sub EvaluateSellingCall(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyDelta = recRow("FamDelta")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0

        For Each APRow As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            adjustment = 0
            sym = APRow("Symbol").Trim()
            If IsACall(sym) AndAlso IsInTheFamily(sym, underlier) Then
                symCurrPosition = APRow("Units")
                If symCurrPosition > 0 Then
                    hedgeQty = CalcQtyNeededToHedge(sym, familyDelta, tDate)
                    If hedgeQty > 0 Then
                        If symCurrPosition < hedgeQty Then
                            hedgeQty = symCurrPosition ' sell all you have
                            adjustment = (1 - (symCurrPosition / hedgeQty)) * -50           ' because incomplete hedge
                        End If
                        newRow = myDataSet.Tables("CandidateRecommendationsTbl").Rows.Add()
                        newRow("Type") = "Sell"
                        newRow("Symbol") = sym
                        newRow("Qty") = hedgeQty
                        newRow("Score") = baseScore + adjustment
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub EvaluateSellingShortCall(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyDelta = recRow("FamDelta")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0

        If TooCloseToMaxMargins() Then
            Exit Sub ' we have no more credit
        End If
        ' only these options will be considered, in this order
        ' you might add/subtract To/from the list, for example adding the JULY otions
        For Each partialSymbol As String In {"_COCTE", "_COCTD", "_COCTC", "_COCTB", "_COCTA"}
            adjustment = 0
            sym = underlier + partialSymbol
            If Not IsInIP(sym) Then
                symCurrPosition = GetCurrentPositionInAP(sym)
                If symCurrPosition <= 0 Then   ' because if long cannot sell short
                    hedgeQty = CalcQtyNeededToHedge(sym, familyDelta, tDate)
                    maxShort = MaxShortWithinContraints(sym, tDate)
                    If hedgeQty > maxShort Then
                        hedgeQty = maxShort
                        adjustment = (1 - (maxShort / hedgeQty)) * -50
                    End If
                    If hedgeQty > 0 Then
                        Select Case partialSymbol
                            Case "_COCTE"
                                adjustment = adjustment + 10
                            Case "_COCTD"
                                adjustment = adjustment + 5
                            Case "_COCTC"
                                adjustment = adjustment + 3
                            Case "_COCTB"
                                adjustment = adjustment + 2
                            Case "_COCTA"
                                adjustment = adjustment + 1
                        End Select
                        newRow = myDataSet.Tables("CandidateRecommendationsTbl").Rows.Add()
                        newRow("Type") = "SellShort"
                        newRow("Symbol") = sym
                        newRow("Qty") = hedgeQty
                        newRow("Score") = baseScore + adjustment
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub EvaluateBuyingBackPut(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyDelta = recRow("FamDelta")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        Dim maxBuy As Integer = 0
        If AvailableCashInLow() Then
            Exit Sub
        End If
        For Each dr As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            adjustment = 0
            sym = dr("Symbol").ToString().Trim()
            If IsAStock(sym) Or sym = "CAccount" Then
                ' skip
            Else
                If IsAPut(sym) Then
                    If GetUnderlier(sym) = underlier Then
                        symCurrPosition = dr("Units")
                        If symCurrPosition < 0 Then
                            hedgeQty = CalcQtyNeededToHedge(sym, familyDelta, tDate)
                            If hedgeQty > Math.Abs(symCurrPosition) Then
                                hedgeQty = Math.Abs(symCurrPosition) ' buy back all that you have
                                'adjustment = -50
                                ' how much can you afford?
                                maxBuy = MaxPurchasePossible(sym, tDate)
                                If maxBuy < hedgeQty Then
                                    adjustment = (1 - (maxBuy / hedgeQty)) * -50
                                    hedgeQty = maxBuy
                                    'adjustment = adjustment - 50
                                End If
                                If hedgeQty > 0 Then
                                    newRow = myDataSet.Tables("CandidateRecommendationsTbl").Rows.Add()
                                    newRow("Type") = "Buy"
                                    newRow("Symbol") = sym
                                    newRow("Qty") = hedgeQty
                                    newRow("Score") = baseScore + adjustment
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub EvaluateBuyingPut(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyDelta = recRow("FamDelta")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        Dim maxBuy As Double = 0
        If AvailableCashInLow() Then
            Exit Sub
        End If
        ' arbitrarily only considers OCT options in this order - you can change that
        For Each partialSymbol As String In {"_POCTA", "_POCTB", "_POCTC", "_POCTD", "_POCTE"}
            adjustment = 0
            sym = underlier + partialSymbol
            If Not IsInIP(sym) Then
                symCurrPosition = GetCurrentPositionInAP(sym)
                If symCurrPosition >= 0 Then ' if short it is a buyback
                    hedgeQty = CalcQtyNeededToHedge(sym, familyDelta, tDate)
                    maxBuy = MaxPurchasePossible(sym, tDate)  ' how much can you afford?
                    If maxBuy < hedgeQty Then
                        adjustment = (1 - (maxBuy / hedgeQty)) * -50
                        hedgeQty = maxBuy
                        'adjustment = -50
                    End If
                    If hedgeQty > 0 Then
                        Select Case partialSymbol
                            Case "_POCTE"
                                adjustment = adjustment + 1
                            Case "_POCTD"
                                adjustment = adjustment + 3
                            Case "_POCTC"
                                adjustment = adjustment + 5
                            Case "_POCTB"
                                adjustment = adjustment + 4
                            Case "_POCTA"
                                adjustment = adjustment + 2
                        End Select
                        newRow = myDataSet.Tables("CandidateRecommendationsTbl").Rows.Add()
                        newRow("Type") = "Buy"
                        newRow("Symbol") = sym
                        newRow("Qty") = hedgeQty
                        newRow("Score") = baseScore + adjustment
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub EvaluateSellingShortStock(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyDelta = recRow("FamDelta")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        If TooCloseToMaxMargins() Then
            Exit Sub
        End If
        If Not IsInIP(underlier) Then
            underlierCurrPos = GetCurrentPositionInAP(underlier)
            If underlierCurrPos <= 0 Then ' if long we cannot sell short
                hedgeQty = CalcQtyNeededToHedge(underlier, familyDelta, tDate)
                maxShort = MaxShortWithinContraints(underlier, tDate)
                If hedgeQty > maxShort Then
                    adjustment = (1 - (maxShort / hedgeQty)) * -50
                    hedgeQty = maxShort
                    'adjustment = -50
                End If
                If hedgeQty > 0 Then
                    newRow = myDataSet.Tables("CandidateRecommendationsTbl").Rows.Add()
                    newRow("Type") = "SellShort"
                    newRow("Symbol") = underlier
                    newRow("Qty") = hedgeQty
                    newRow("Score") = baseScore + adjustment
                End If
            End If
        End If
    End Sub

    Public Sub EvaluateSellingPut(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyDelta = recRow("FamDelta")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPosition As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        For Each dr As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            adjustment = 0
            sym = dr("symbol").ToString().Trim()
            If IsAStock(sym) Or sym = "CAccount" Then
                ' skip
            Else
                If (IsAPut(sym)) Then
                    If GetUnderlier(sym) = underlier Then
                        symCurrPosition = GetCurrentPositionInAP(sym)
                        If symCurrPosition > 0 Then
                            hedgeQty = CalcQtyNeededToHedge(sym, familyDelta, tDate)
                            If symCurrPosition < hedgeQty Then
                                adjustment = (1 - (symCurrPosition / hedgeQty)) * -50
                                hedgeQty = symCurrPosition
                                'adjustment = -50
                            End If
                            If hedgeQty > 0 Then
                                newRow = myDataSet.Tables("CandidateRecommendationsTbl").Rows.Add()
                                newRow("Type") = "Sell"
                                newRow("Symbol") = sym
                                newRow("Qty") = hedgeQty
                                newRow("Score") = baseScore + adjustment
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub EvaluateSellingShortPut(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyDelta = recRow("FamDelta")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        If TooCloseToMaxMargins() Then
            Exit Sub
        End If
        ' arbitrary order, arbitrary exclusion of jul options
        For Each partialSymbol As String In {"_POCTA", "_POCTB", "_POCTC", "_POCTD", "_POCTE"}
            adjustment = 0
            sym = underlier + partialSymbol
            If Not IsInIP(sym) Then
                symCurrPosition = GetCurrentPositionInAP(sym)
                If symCurrPosition <= 0 Then
                    hedgeQty = CalcQtyNeededToHedge(sym, familyDelta, tDate)
                    maxShort = MaxShortWithinContraints(sym, tDate)
                    If maxShort < hedgeQty Then
                        adjustment = (1 - (maxShort / hedgeQty)) * -50
                        hedgeQty = maxShort
                        ' adjustment = -50
                    End If
                    If hedgeQty > 0 Then
                        Select Case partialSymbol
                            Case "_POCTE"
                                adjustment = adjustment + 1
                            Case "_POCTD"
                                adjustment = adjustment + 2
                            Case "_POCTC"
                                adjustment = adjustment + 7
                            Case "_POCTB"
                                adjustment = adjustment + 5
                            Case "_POCTA"
                                adjustment = adjustment + 3
                        End Select
                        newRow = myDataSet.Tables("CandidateRecommendationsTbl").Rows.Add()
                        newRow("Type") = "SellShort"
                        newRow("Symbol") = sym
                        newRow("Qty") = hedgeQty
                        newRow("Score") = baseScore + adjustment
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub EvaluateBuyingBackCall(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyDelta = recRow("FamDelta")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        Dim maxBuy As Integer = 0
        If AvailableCashInLow() Then
            Exit Sub
        End If
        For Each dr As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            sym = dr("Symbol").ToString().Trim()
            If IsAStock(sym) Or sym = "CAccount" Then
                ' skip
            Else
                If IsACall(sym) Then
                    If GetUnderlier(sym) = underlier Then
                        symCurrPosition = dr("Units")
                        If symCurrPosition < 0 Then
                            hedgeQty = CalcQtyNeededToHedge(sym, familyDelta, tDate)
                            If Math.Abs(symCurrPosition) < hedgeQty Then
                                adjustment = (1 - (maxBuy / hedgeQty)) * -50
                                hedgeQty = Math.Abs(symCurrPosition) ' buy back all that you have
                                'adjustment = -50
                            End If
                            maxBuy = MaxPurchasePossible(sym, tDate)
                            If maxBuy < hedgeQty Then
                                hedgeQty = maxBuy
                                adjustment = -50
                            End If
                            If hedgeQty > 0 Then
                                newRow = myDataSet.Tables("CandidateRecommendationsTbl").Rows.Add()
                                newRow("Type") = "Buy"
                                newRow("Symbol") = sym
                                newRow("Qty") = hedgeQty
                                newRow("Score") = baseScore + adjustment
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub EvaluateBuyingBackStock(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyDelta = recRow("FamDelta")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        Dim maxBuy As Integer = 0
        If AvailableCashInLow() Then
            Exit Sub
        End If
        If Not IsInIP(underlier) Then
            symCurrPosition = GetCurrentPositionInAP(underlier)
            If symCurrPosition < 0 Then
                hedgeQty = CalcQtyNeededToHedge(underlier, familyDelta, tDate)
                If Math.Abs(symCurrPosition) < hedgeQty Then
                    adjustment = (1 - (symCurrPosition / hedgeQty)) * -50
                    hedgeQty = Math.Abs(symCurrPosition) ' buy back all that you have
                    'adjustment = -50
                End If
                maxBuy = MaxPurchasePossible(underlier, tDate) ' how much can you afford?
                If maxBuy < hedgeQty Then
                    adjustment = (1 - (maxBuy / hedgeQty)) * -50
                    hedgeQty = maxBuy
                    'adjustment = adjustment - 10
                End If
                If hedgeQty > 0 Then
                    newRow = myDataSet.Tables("CandidateRecommendationsTbl").Rows.Add()
                    newRow("Type") = "Buy"
                    newRow("Symbol") = underlier
                    newRow("Qty") = hedgeQty
                    newRow("Score") = baseScore + adjustment
                End If
            End If
        End If
    End Sub

    Public Sub EvaluateBuyingCall(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyDelta = recRow("FamDelta")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        Dim maxBuy As Integer = 0
        If AvailableCashInLow() Then
            Exit Sub
        End If
        ' only considers OCT options - you can change this
        For Each partialSymbol As String In {"_COCTE", "_COCTD", "_COCTC", "_COCTB", "_COCTA"}
            adjustment = 0
            sym = underlier + partialSymbol
            If Not IsInIP(sym) Then
                symCurrPosition = GetCurrentPosition(sym)

                If symCurrPosition >= 0 Then                  ' if short is a buyback
                    hedgeQty = CalcQtyNeededToHedge(sym, familyDelta, tDate)
                    maxBuy = MaxPurchasePossible(sym, tDate)   ' how much can you afford?
                    If maxBuy < hedgeQty Then
                        adjustment = (1 - (maxBuy / hedgeQty)) * -50
                        hedgeQty = maxBuy
                        'adjustment = -50
                    End If
                    If hedgeQty > 0 Then
                        Select Case partialSymbol
                            Case "_COCTE"
                                adjustment = adjustment + 1
                            Case "_COCTD"
                                adjustment = adjustment + 2
                            Case "_COCTC"
                                adjustment = adjustment + 10
                            Case "_COCTB"
                                adjustment = adjustment + 5
                            Case "_COCTA"
                                adjustment = adjustment + 4
                        End Select
                        newRow = myDataSet.Tables("CandidateRecommendationsTbl").Rows.Add()
                        newRow("Type") = "Buy"
                        newRow("Symbol") = sym
                        newRow("Qty") = hedgeQty
                        newRow("Score") = baseScore + adjustment
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub EvaluateBuyingStock(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyDelta = recRow("FamDelta")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        Dim maxBuy As Integer = 0

        If AvailableCashInLow() Then
            Exit Sub
        End If
        If IsInIP(underlier) Then
            Exit Sub
        End If
        symCurrPosition = GetCurrentPositionInAP(underlier)
        If symCurrPosition < 0 Then     ' if short then we need a buyback
            Exit Sub
        End If
        hedgeQty = CalcQtyNeededToHedge(underlier, familyDelta, tDate)
        ' how much can you afford?
        maxBuy = MaxPurchasePossible(underlier, tDate)
        If maxBuy < hedgeQty Then
            adjustment = (1 - (maxBuy / hedgeQty)) * -50
            hedgeQty = maxBuy
            'adjustment = -50
        End If
        If hedgeQty > 0 Then
            newRow = myDataSet.Tables("CandidateRecommendationsTbl").Rows.Add()
            newRow("Type") = "Buy"
            newRow("Symbol") = underlier
            newRow("Qty") = hedgeQty
            newRow("Score") = baseScore + adjustment
        End If
    End Sub

    Public Sub EvaluateSellingStockG(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyGamma = recRow("FamilyGamma")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim adjustment As Double = 0
        If IsInIP(underlier) Then
            Exit Sub   ' cannot sell if in IP
        Else
            underlierCurrPos = GetCurrentPositionInAP(underlier)
            If underlierCurrPos <= 0 Then ' we cannot sell since we are not long
                Exit Sub
            Else
                hedgeQty = CalcQtyNeededToHedgeG(underlier, familyGamma, tDate)
                If hedgeQty = 0 Then
                    Exit Sub  ' nothing to do
                Else
                    adjustment = (1 - (underlierCurrPos / hedgeQty)) * -50
                    If hedgeQty > underlierCurrPos Then     ' you have fewer than needed
                        hedgeQty = underlierCurrPos         ' sell all you have
                        'adjustment = -(underlierCurrPos / hedgeQty) * 50               ' arbitrary adjustment!
                    End If
                    newRow = myDataSet.Tables("CandidateRecommendationsGTbl").Rows.Add()
                    newRow("Type") = "Sell"
                    newRow("Symbol") = underlier
                    newRow("Qty") = hedgeQty
                    newRow("Score") = baseScore + adjustment
                End If
            End If
        End If
    End Sub

    Public Sub EvaluateSellingCallG(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyGamma = recRow("FamilyGamma")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0

        For Each APRow As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            adjustment = 0
            sym = APRow("Symbol").Trim()
            If IsACall(sym) AndAlso IsInTheFamily(sym, underlier) Then
                symCurrPosition = APRow("Units")
                If symCurrPosition > 0 Then
                    hedgeQty = CalcQtyNeededToHedgeG(sym, familyGamma, tDate)
                    If hedgeQty > 0 Then
                        If symCurrPosition < hedgeQty Then
                            adjustment = (1 - (symCurrPosition / hedgeQty)) * -50
                            hedgeQty = symCurrPosition ' sell all you have
                            'adjustment = -50           ' because incomplete hedge
                        End If
                        newRow = myDataSet.Tables("CandidateRecommendationsGTbl").Rows.Add()
                        newRow("Type") = "Sell"
                        newRow("Symbol") = sym
                        newRow("Qty") = hedgeQty
                        newRow("Score") = baseScore + adjustment
                    End If
                End If
            End If
        Next
    End Sub
    Public Sub EvaluateSellingShortCallG(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyGamma = recRow("FamilyGamma")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0

        If TooCloseToMaxMargins() Then
            Exit Sub ' we have no more credit
        End If
        ' only these options will be considered, in this order
        ' you might add/subtract To/from the list, for example adding the JULY otions
        For Each partialSymbol As String In {"_COCTE", "_COCTD", "_COCTC", "_COCTB", "_COCTA"}
            adjustment = 0
            sym = underlier + partialSymbol
            If Not IsInIP(sym) Then
                symCurrPosition = GetCurrentPositionInAP(sym)
                If symCurrPosition <= 0 Then   ' because if long cannot sell short
                    hedgeQty = CalcQtyNeededToHedgeG(sym, familyGamma, tDate)
                    maxShort = MaxShortWithinConstraints(sym, tDate)
                    If hedgeQty > maxShort Then
                        adjustment = (1 - (maxShort / hedgeQty)) * -50
                        hedgeQty = maxShort
                        'adjustment = -50
                    End If
                    If hedgeQty > 0 Then
                        Select Case partialSymbol
                            Case "_COCTE"
                                adjustment = adjustment + 10
                            Case "_COCTD"
                                adjustment = adjustment + 5
                            Case "_COCTC"
                                adjustment = adjustment + 3
                            Case "_COCTB"
                                adjustment = adjustment + 2
                            Case "_COCTA"
                                adjustment = adjustment + 1
                        End Select
                        newRow = myDataSet.Tables("CandidateRecommendationsGTbl").Rows.Add()
                        newRow("Type") = "SellShort"
                        newRow("Symbol") = sym
                        newRow("Qty") = hedgeQty
                        newRow("Score") = baseScore + adjustment
                    End If
                End If
            End If
        Next
    End Sub
    Public Sub EvaluateBuyingBackPutG(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyGamma = recRow("FamilyGamma")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        Dim maxBuy As Integer = 0
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        For Each dr As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            adjustment = 0
            sym = dr("Symbol").ToString().Trim()
            If IsAStock(sym) Or sym = "CAccount" Then
                ' skip
            Else
                If IsAPut(sym) Then
                    If GetUnderlier(sym) = underlier Then
                        symCurrPosition = dr("Units")
                        If symCurrPosition < 0 Then
                            hedgeQty = CalcQtyNeededToHedgeG(sym, familyGamma, tDate)
                            If hedgeQty > Math.Abs(symCurrPosition) Then
                                adjustment = (1 - (maxBuy / hedgeQty)) * -10
                                hedgeQty = Math.Abs(symCurrPosition) ' buy back all that you have
                                'adjustment = -50
                                ' how much can you afford?
                                maxBuy = MaxPurchasePossible(sym, tDate)
                                If maxBuy < hedgeQty Then
                                    adjustment = (1 - (maxBuy / hedgeQty)) * -10

                                    hedgeQty = maxBuy
                                    'adjustment = adjustment - 50
                                End If
                                If hedgeQty > 0 Then
                                    newRow = myDataSet.Tables("CandidateRecommendationsGTbl").Rows.Add()
                                    newRow("Type") = "Buy"
                                    newRow("Symbol") = sym
                                    newRow("Qty") = hedgeQty
                                    newRow("Score") = baseScore + adjustment
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub
    Public Sub EvaluateBuyingPutG(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyGamma = recRow("FamilyGamma")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        Dim maxBuy As Double = 0
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        ' arbitrarily only considers OCT options in this order - you can change that
        For Each partialSymbol As String In {"_POCTA", "_POCTB", "_POCTC", "_POCTD", "_POCTE"}
            adjustment = 0
            sym = underlier + partialSymbol
            If Not IsInIP(sym) Then
                symCurrPosition = GetCurrentPositionInAP(sym)
                If symCurrPosition >= 0 Then ' if short it is a buyback
                    hedgeQty = CalcQtyNeededToHedgeG(sym, familyGamma, tDate)
                    maxBuy = MaxPurchasePossible(sym, tDate)  ' how much can you afford?
                    If maxBuy < hedgeQty Then
                        adjustment = (1 - (maxBuy / hedgeQty)) * -50
                        hedgeQty = maxBuy
                        'adjustment = -50
                    End If
                    If hedgeQty > 0 Then
                        Select Case partialSymbol
                            Case "_POCTE"
                                adjustment = adjustment + 1
                            Case "_POCTD"
                                adjustment = adjustment + 3
                            Case "_POCTC"
                                adjustment = adjustment + 5
                            Case "_POCTB"
                                adjustment = adjustment + 4
                            Case "_POCTA"
                                adjustment = adjustment + 2
                        End Select
                        newRow = myDataSet.Tables("CandidateRecommendationsGTbl").Rows.Add()
                        newRow("Type") = "Buy"
                        newRow("Symbol") = sym
                        newRow("Qty") = hedgeQty
                        newRow("Score") = baseScore + adjustment
                    End If
                End If
            End If
        Next
    End Sub
    Public Sub EvaluateSellingShortStockG(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyGamma = recRow("FamilyGamma")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        If TooCloseToMaxMargins() Then
            Exit Sub
        End If
        If Not IsInIP(underlier) Then
            underlierCurrPos = GetCurrentPositionInAP(underlier)
            If underlierCurrPos <= 0 Then ' if long we cannot sell short
                hedgeQty = CalcQtyNeededToHedgeG(underlier, familyGamma, tDate)
                maxShort = MaxShortWithinConstraints(underlier, tDate)
                If hedgeQty > maxShort Then
                    adjustment = (1 - (maxShort / hedgeQty)) * -50
                    hedgeQty = maxShort
                    'adjustment = -50
                End If
                If hedgeQty > 0 Then
                    newRow = myDataSet.Tables("CandidateRecommendationsGTbl").Rows.Add()
                    newRow("Type") = "SellShort"
                    newRow("Symbol") = underlier
                    newRow("Qty") = hedgeQty
                    newRow("Score") = baseScore + adjustment
                End If
            End If
        End If
    End Sub
    Public Sub EvaluateSellingPutG(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyGamma = recRow("FamilyGamma")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPosition As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        For Each dr As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            adjustment = 0
            sym = dr("symbol").ToString().Trim()
            If IsAStock(sym) Or sym = "CAccount" Then
                ' skip
            Else
                If (IsAPut(sym)) Then
                    If GetUnderlier(sym) = underlier Then
                        symCurrPosition = GetCurrentPositionInAP(sym)
                        If symCurrPosition > 0 Then
                            hedgeQty = CalcQtyNeededToHedgeG(sym, familyGamma, tDate)
                            If symCurrPosition < hedgeQty Then
                                adjustment = (1 - (symCurrPosition / hedgeQty)) * -50
                                hedgeQty = symCurrPosition
                                'adjustment = -50
                            End If
                            If hedgeQty > 0 Then
                                newRow = myDataSet.Tables("CandidateRecommendationsGTbl").Rows.Add()
                                newRow("Type") = "Sell"
                                newRow("Symbol") = sym
                                newRow("Qty") = hedgeQty
                                newRow("Score") = baseScore + adjustment
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub
    Public Sub EvaluateSellingShortPutG(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyGamma = recRow("FamilyGamma")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        If TooCloseToMaxMargins() Then
            Exit Sub
        End If
        ' arbitrary order, arbitrary exclusion of jul options
        For Each partialSymbol As String In {"_POCTA", "_POCTB", "_POCTC", "_POCTD", "_POCTE"}
            adjustment = 0
            sym = underlier + partialSymbol
            If Not IsInIP(sym) Then
                symCurrPosition = GetCurrentPositionInAP(sym)
                If symCurrPosition <= 0 Then
                    hedgeQty = CalcQtyNeededToHedgeG(sym, familyGamma, tDate)
                    maxShort = MaxShortWithinConstraints(sym, tDate)
                    If maxShort < hedgeQty Then
                        adjustment = (1 - (maxShort / hedgeQty)) * -50
                        hedgeQty = maxShort
                        'adjustment = -50
                    End If
                    If hedgeQty > 0 Then
                        Select Case partialSymbol
                            Case "_POCTE"
                                adjustment = adjustment + 1
                            Case "_POCTD"
                                adjustment = adjustment + 2
                            Case "_POCTC"
                                adjustment = adjustment + 7
                            Case "_POCTB"
                                adjustment = adjustment + 5
                            Case "_POCTA"
                                adjustment = adjustment + 3
                        End Select
                        newRow = myDataSet.Tables("CandidateRecommendationsGTbl").Rows.Add()
                        newRow("Type") = "SellShort"
                        newRow("Symbol") = sym
                        newRow("Qty") = hedgeQty
                        newRow("Score") = baseScore + adjustment
                    End If
                End If
            End If
        Next
    End Sub
    Public Sub EvaluateBuyingBackCallG(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyGamma = recRow("FamilyGamma")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        Dim maxBuy As Integer = 0
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        For Each dr As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            sym = dr("Symbol").ToString().Trim()
            If IsAStock(sym) Or sym = "CAccount" Then
                ' skip
            Else
                If IsACall(sym) Then
                    If GetUnderlier(sym) = underlier Then
                        symCurrPosition = dr("Units")
                        If symCurrPosition < 0 Then
                            hedgeQty = CalcQtyNeededToHedgeG(sym, familyGamma, tDate)
                            If Math.Abs(symCurrPosition) < hedgeQty Then
                                adjustment = (1 - (symCurrPosition / hedgeQty)) * -10
                                hedgeQty = Math.Abs(symCurrPosition) ' buy back all that you have
                                'adjustment = -50
                            End If
                            maxBuy = MaxPurchasePossible(sym, tDate)
                            If maxBuy < hedgeQty Then
                                adjustment = (1 - (maxBuy / hedgeQty)) * -10
                                hedgeQty = maxBuy

                                'adjustment = -50
                            End If
                            If hedgeQty > 0 Then
                                newRow = myDataSet.Tables("CandidateRecommendationsGTbl").Rows.Add()
                                newRow("Type") = "Buy"
                                newRow("Symbol") = sym
                                newRow("Qty") = hedgeQty
                                newRow("Score") = baseScore + adjustment
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub
    Public Sub EvaluateBuyingBackStockG(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyGamma = recRow("FamilyGamma")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        Dim maxBuy As Integer = 0
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        If Not IsInIP(underlier) Then
            symCurrPosition = GetCurrentPositionInAP(underlier)
            If symCurrPosition < 0 Then
                hedgeQty = CalcQtyNeededToHedgeG(underlier, familyGamma, tDate)
                If Math.Abs(symCurrPosition) < hedgeQty Then
                    adjustment = (1 - (symCurrPosition / hedgeQty)) * -10
                    hedgeQty = Math.Abs(symCurrPosition) ' buy back all that you have
                    ' adjustment = -50
                End If
                maxBuy = MaxPurchasePossible(underlier, tDate) ' how much can you afford?
                If maxBuy < hedgeQty Then
                    adjustment = (1 - (maxBuy / hedgeQty)) * -50
                    hedgeQty = maxBuy
                    'adjustment = adjustment - 10
                End If
                If hedgeQty > 0 Then
                    newRow = myDataSet.Tables("CandidateRecommendationsGTbl").Rows.Add()
                    newRow("Type") = "Buy"
                    newRow("Symbol") = underlier
                    newRow("Qty") = hedgeQty
                    newRow("Score") = baseScore + adjustment
                End If
            End If
        End If
    End Sub
    Public Sub EvaluateBuyingCallG(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyGamma = recRow("FamilyGamma")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim sym As String
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        Dim maxBuy As Integer = 0
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        ' only considers OCT options - you can change this
        For Each partialSymbol As String In {"_COCTE", "_COCTD", "_COCTC", "_COCTB", "_COCTA"}
            adjustment = 0
            sym = underlier + partialSymbol
            If Not IsInIP(sym) Then
                symCurrPosition = GetCurrentPositionInAP(sym)
                If symCurrPosition >= 0 Then                  ' if short is a buyback
                    hedgeQty = CalcQtyNeededToHedgeG(sym, familyGamma, tDate)
                    maxBuy = MaxPurchasePossible(sym, tDate)   ' how much can you afford?
                    If maxBuy < hedgeQty Then
                        adjustment = (1 - (maxBuy / hedgeQty)) * -50
                        hedgeQty = maxBuy
                        'adjustment = -50
                    End If
                    If hedgeQty > 0 Then
                        Select Case partialSymbol
                            Case "_COCTE"
                                adjustment = adjustment + 1
                            Case "_COCTD"
                                adjustment = adjustment + 2
                            Case "_COCTC"
                                adjustment = adjustment + 10
                            Case "_COCTB"
                                adjustment = adjustment + 5
                            Case "_COCTA"
                                adjustment = adjustment + 4
                        End Select
                        newRow = myDataSet.Tables("CandidateRecommendationsGTbl").Rows.Add()
                        newRow("Type") = "Buy"
                        newRow("Symbol") = sym
                        newRow("Qty") = hedgeQty
                        newRow("Score") = baseScore + adjustment
                    End If
                End If
            End If
        Next
    End Sub
    Public Sub EvaluateBuyingStockG(baseScore As Integer, recRow As DataRow, tDate As Date)
        Dim familyGamma = recRow("FamilyGamma")
        Dim underlier As String = recRow("Underlier").trim()
        Dim underlierCurrPos As Double = 0
        Dim hedgeQty As Double = 0
        Dim newRow As DataRow
        Dim adjustment As Double = 0
        Dim symCurrPosition As Double = 0
        Dim maxShort As Double = 0
        Dim maxBuy As Integer = 0

        If AvailableCashIsLow() Then
            Exit Sub
        End If
        If IsInIP(underlier) Then
            Exit Sub
        End If
        symCurrPosition = GetCurrentPositionInAP(underlier)
        If symCurrPosition < 0 Then     ' if short then we need a buyback
            Exit Sub
        End If
        hedgeQty = CalcQtyNeededToHedgeG(underlier, familyGamma, tDate)
        ' how much can you afford?
        maxBuy = MaxPurchasePossible(underlier, tDate)
        If maxBuy < hedgeQty Then
            adjustment = (1 - (maxBuy / hedgeQty)) * -25
            hedgeQty = maxBuy
            'adjustment = -50
        End If
        If hedgeQty > 0 Then
            newRow = myDataSet.Tables("CandidateRecommendationsGTbl").Rows.Add()
            newRow("Type") = "Buy"
            newRow("Symbol") = underlier
            newRow("Qty") = hedgeQty
            newRow("Score") = baseScore + adjustment
        End If
    End Sub

End Module
