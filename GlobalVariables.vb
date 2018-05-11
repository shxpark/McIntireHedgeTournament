Module GlobalVariables

    Public myConnection As SqlClient.SqlConnection
    Public myCommand As SqlClient.SqlCommand
    Public myDataAdapter As SqlClient.SqlDataAdapter
    Public myDataSet As DataSet

    Public activeDB As String = ""
    Public traderMode As String = ""

    Public teamID As String = "09"
    Public teamPortfolioTableName As String = "PortfolioTeam" + teamID
    Public confirmationTicketTableName As String = "ConfirmationTicketTeam" + teamID

    Public initialCAccount As Double = 0
    Public iRate As Double = 0
    Public startDate As Date = "1/1/1"
    Public currentDate As Date = "1/1/1"
    Public endDate As Date = "1/1/1"
    Public maxMargin As Double = 0

    Public TPVatStart As Double = 0
    Public lastPriceDownloadDate As Date = "1/1/1"

    Public IPValue As Double = 0
    Public APValue As Double = 0

    Public trType As String = ""
    Public trSecurityType As String = ""
    Public trSymbol As String = ""
    Public trQty As Double = 0
    Public trPrice As Double = 0
    Public trCost As Double = 0
    Public trTotValue As Double = 0
    Public trStrike As Double = 0
    Public CAccount As Double = 0
    Public CAccountAT As Double = 0
    Public lastTransactionDate As Date = "1/1/1"
    Public interestSLT As Double = 0

    '---Start part 4---
    Public turnOffIP As Boolean = False
    Public margin As Double = 0
    Public TE As Double = 0
    Public TEpercent As Double = 0
    Public sumTE As Double = 0
    Public TPV As Double = 0
    Public TaTPV As Double = 0

    '----5
    Public marginAT As Double = 0

    '---6
    Public trDelta As Double = 0

    Public lastTEUpdate As Date = "1/1/1"
End Module
