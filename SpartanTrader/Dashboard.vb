
Public Class Dashboard

    Private Sub Sheet4_Startup() Handles Me.Startup

    End Sub

    Private Sub Sheet4_Shutdown() Handles Me.Shutdown

    End Sub

    Public Sub InitializeDisplay()
        InitialPositionsLO.DataBodyRange.Interior.Color = System.Drawing.Color.MidnightBlue
        InitialPositionsLO.ListColumns(1).Range.Font.Color = System.Drawing.Color.White
        InitialPositionsLO.ListColumns(2).Range.NumberFormat = "[White]#,##0;[Red]-#,##0"
        InitialPositionsLO.ListColumns(3).Range.NumberFormat = "[Green]$ #,##0;[Red]$ -#,##0"
        InitialPositionsLO.Range.Columns.ColumnWidth = 13

        AcquiredPositionsLO.DataBodyRange.Interior.Color = System.Drawing.Color.MidnightBlue
        AcquiredPositionsLO.ListColumns(1).Range.Font.Color = System.Drawing.Color.White
        AcquiredPositionsLO.ListColumns(2).Range.NumberFormat = "[White]#,##0;[Red]-#,##0"
        AcquiredPositionsLO.ListColumns(3).Range.NumberFormat = "[Green]$ #,##0;[Red]$ -#,##0"
        AcquiredPositionsLO.Range.Columns.ColumnWidth = 13
    End Sub
End Class
