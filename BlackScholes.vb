Module BlackScholes
    Public Function CalcDelta(symbol As String, targetDate As Date) As Double
        Dim sigma As Double
        Dim K As Double
        Dim S As Double
        Dim r As Double = iRate
        Dim t As Double
        Dim ts As TimeSpan
        Dim underlier As String
        Dim d1 As Double

        If symbol = "CAccount" Then
            Return 0
        End If
        If IsAStock(symbol) Then
            Return 1
        End If
        If targetDate.Date >= GetExpiration(symbol).Date Then
            Return 0
        End If
        If GetAsk(symbol, targetDate) = 0 Then
            Return 0
        End If

        underlier = GetUnderlier(symbol)
        sigma = GetVol(underlier)
        K = GetStrike(symbol)
        S = CalcMTM(underlier, targetDate)
        ts = GetExpiration(symbol).Date - targetDate.Date
        t = ts.Days / 365.25

        d1 = (Math.Log(S / K) + (r + sigma * sigma / 2) * t) / (sigma * Math.Sqrt(t))
        If IsACall(symbol) Then
            Return Globals.ThisWorkbook.Application.WorksheetFunction.Norm_S_Dist(d1, True)
        End If
        If IsAPut(symbol) Then
            Return (Globals.ThisWorkbook.Application.WorksheetFunction.Norm_S_Dist(d1, True) - 1)
        End If
        Return 0
    End Function

    Public Function CalcFamilyDelta(t As String, targetDate As Date) As Double
        Dim tempFamDelta As Double = 0
        Dim delta As Double = 0
        Dim sym As String
        t = t.Trim()

        For Each dr As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            sym = dr("Symbol").ToString().Trim()
            If IsInTheFamily(sym, t) Then
                delta = CalcDelta(sym, currentDate)
                tempFamDelta = tempFamDelta + delta * dr("Units")
            End If
        Next

        If myDataSet.Tables.Contains("InitialPositionsTbl") Then
            For Each dr As DataRow In myDataSet.Tables("InitialPositionsTbl").Rows
                sym = dr("Symbol").ToString().Trim()
                If IsInTheFamily(sym, t) Then
                    delta = CalcDelta(sym, currentDate)
                    tempFamDelta = tempFamDelta + delta * dr("Units")
                End If
            Next
        End If
        Return tempFamDelta
    End Function

    Public Function GetVol(symbol As String) As Double
        symbol = symbol.Trim()
        Select Case symbol
            Case "AAPL"
                Return 0.295
            Case "AMT"
                Return 0.243
            Case "AMZ"
                Return 0.311
            Case "BIDU"
                Return 0.387
            Case "FB"
                Return 0.371
            Case "GOOG"
                Return 0.347
            Case "LUV"
                Return 0.2615
            Case "MSFT"
                Return 0.313
            Case "SBUX"
                Return 0.203
            Case "SNAP"
                Return 1.111
            Case "TESLA"
                Return 0.412
            Case "WMT"
                Return 0.336
            Case Else
                Return 0.2
        End Select
    End Function

    Public Function CalcQtyNeededToHedge(sym As String, familyDelta As Double, targetDate As Date) As Integer
        Dim familyDeltaTarget As Double = 0

        Dim q As Double
        Dim delta As Double = 0
        delta = CalcDelta(sym, targetDate)
        If Math.Abs(delta) < 0.125 Then
            Return 0
        Else
            q = (familyDeltaTarget - familyDelta) / delta
            Return Math.Abs(Math.Round(q))
        End If
    End Function
End Module
