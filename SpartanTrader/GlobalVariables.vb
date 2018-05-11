Module GlobalVariables

    Public myConnection As SqlClient.SqlConnection
    Public myCommand As SqlClient.SqlCommand
    Public myDataAdapter As SqlClient.SqlDataAdapter
    Public myDataSet As DataSet

    Public activeDB As String = ""
    Public traderMode As String = ""

    Public teamID As String = "30"
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


End Module
