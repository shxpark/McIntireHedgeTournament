Module DBProcedures

    ' This sub sets up the ADO objects and connects to the DB indicated in myConnectionString

    Public Sub ConnectToDB(myConnString As String)
        'refer to the slide deck to see the big pic
        myConnection = New SqlClient.SqlConnection      'create connection
        myConnection.ConnectionString = myConnString    'sets connection string
        myCommand = New SqlClient.SqlCommand            'the command represents the SQL query
        myCommand.Connection = myConnection             'links commands and connections
        myDataAdapter = New SqlClient.SqlDataAdapter    'adapter
        myDataAdapter.SelectCommand = myCommand         'links adpater and command
        myDataSet = New DataSet
        myConnection.Open()
    End Sub

    Public Sub CloseDBConnection()
        myConnection.Close()
    End Sub

    Public Sub GetDataTableFromDB(SQLQuery As String, NameOfOutputDataTable As String)
        If myDataSet.Tables.Contains(NameOfOutputDataTable) Then
            myDataSet.Tables(NameOfOutputDataTable).Clear()
        End If
        myCommand.CommandText = SQLQuery
        myDataAdapter.Fill(myDataSet, NameOfOutputDataTable)
    End Sub


    'ST Part 1

    Public Sub ConnectToActiveDB()
        Select Case activeDB
            Case "Alpha"
                ConnectToDB("Data Source=f-sg6m-s4.comm.virginia.edu;" +
                            "Initial Catalog=HedgeTournamentALPHA;Integrated Security=True")
            Case "Beta"
                ConnectToDB("Data Source=f-sg6m-s4.comm.virginia.edu;" +
                            "Initial Catalog=HedgeTournamentBETA;Integrated Security=True")
            Case "Gamma"
                ConnectToDB("Data Source=f-sg6m-s4.comm.virginia.edu;" +
                           "Initial Catalog=HedgeTournamentGAMMA;Integrated Security=True")
            Case Else
                MessageBox.Show("Holy Cow! No active database selected.",
                                "Spartan Trader", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Select
    End Sub

    Public Sub DownloadPricesForOneDay(targetDate As Date)
        Dim mySQL As String
        mySQL = "Select * from StockMarket where Date = '" + targetDate.ToShortDateString() + "';"
        GetDataTableFromDB(mySQL, "StockMarketOneDayTbl")

        mySQL = "Select * from OptionMarket where Date = '" + targetDate.ToShortDateString() + "';"
        GetDataTableFromDB(mySQL, "OptionMarketOneDayTbl")
        lastPriceDownloadDate = targetDate
    End Sub

    Public Function DownloadCurrentDate() As Date
        Dim temp As String = ""
        myCommand.CommandText = "Select Value from EnvironmentVariable where Name = 'CurrentDate'"
        temp = myCommand.ExecuteScalar()
        Return Date.Parse(temp)
    End Function

End Module
