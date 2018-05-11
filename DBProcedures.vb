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

    Public Function DownloadLastTransactionDate(targetDate As Date) As Date
        Dim temp As String = ""
        myCommand.CommandText = String.Format("Select max(date) from TransactionQueue where teamid = {0} and date <= '{1}'",
                                              teamID, targetDate.ToShortDateString())
        Try
            temp = myCommand.ExecuteScalar()
            Return Date.Parse(temp)
        Catch myException As Exception
            MessageBox.Show("Holy Cow! Last transaction not found. Set LastTransactionDate to StartDate.",
                            "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return startDate
        End Try
    End Function

    Public Sub RunQuery(SQLQuery As String)
        myCommand.CommandText = SQLQuery
        myCommand.ExecuteNonQuery()
    End Sub

    '--- ST Part 5 ---

    Public Sub ClearTeamPortfolioOnDB()
        RunQuery("Delete From " + teamPortfolioTableName)
    End Sub

    Public Sub UploadPosition(sym As String, newValue As Double)
        Try
            Dim sql As String
            newValue = Math.Round(newValue, 2)
            sym = sym.Trim()

            sql = "Delete from " + teamPortfolioTableName + " where Symbol = '" + sym + "';"
            RunQuery(sql)

            If (newValue = 0) And (sym <> "CAccount") Then
            Else
                sql = String.Format("Insert into {0} Values ('{1}', {2}, 0)",
                                    teamPortfolioTableName, sym, newValue)
                RunQuery(sql)
            End If
        Catch myException As Exception
            MessageBox.Show("I could not upload" + sym + "." +
                            "Maybe this will help: " + myException.Message)
        End Try
    End Sub

    Public Function DownloadCapitalAccount() As Double
        Dim temp As String = ""
        myCommand.CommandText = "Select Units from " + teamPortfolioTableName + " where Symbol = 'CAccount'"
        temp = myCommand.ExecuteScalar()
        If temp = Nothing Then
            UploadPosition("CAccount", initialCAccount)
            Return initialCAccount
        Else
            Return Double.Parse(temp)
        End If


    End Function

    Public Function ThereIsData() As Boolean
        Dim temp As Integer = 0
        myCommand.CommandText = "Select top(1) ask from StockMarket"
        temp = myCommand.ExecuteScalar()
        If temp > 0 Then
            Return True
        Else
            MessageBox.Show("There is no data.", "No data", MessageBoxButtons.OK)
            Return False
        End If
    End Function

End Module
