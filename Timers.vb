Module Timers

    Public WithEvents spyTimer As Timer
    Public WithEvents secondsTimer As Timer
    Public secondsLeft As Integer

    Public Sub StartTimers()
        spyTimer = New Timer
        spyTimer.Interval = 2000 ' do not change these settings
        secondsTimer = New Timer
        secondsTimer.Interval = 1000
        ShowSeconds(0)  ' reset the screen countdown
        secondsTimer.Start()
        spyTimer.Start()
    End Sub

    Private Sub secondsTimer_Tick() Handles secondsTimer.Tick
        secondsLeft = secondsLeft - 1
        ShowSeconds(secondsLeft)
    End Sub

    Public Sub ShowSeconds(secs As Integer)
        Try
            If Math.Abs(secs) > 60 Then
                secs = 0
            End If
            secondsLeft = secs
            If Globals.ThisWorkbook.ActiveSheet.Name = "Dashboard" Then
                Globals.Dashboard.SecondsCell.Value = Math.Abs(secs)
                Select Case secs
                    Case Is < 0
                        Globals.Dashboard.SecondsCell.Font.Color = System.Drawing.Color.Orange
                    Case Is <= 5
                        Globals.Dashboard.SecondsCell.Font.Color = System.Drawing.Color.Red
                    Case Is <= 10
                        Globals.Dashboard.SecondsCell.Font.Color = System.Drawing.Color.Orange
                    Case Else
                        Globals.Dashboard.SecondsCell.Font.Color = System.Drawing.Color.SteelBlue
                End Select
            End If
        Catch
            ' skip
        End Try
        Application.DoEvents()
    End Sub

    Public Sub StopTimers()
        If IsNothing(spyTimer) Then
            ' skip
        Else
            spyTimer.Stop()
            secondsTimer.Stop()
            ShowSeconds(60) ' reset the screen timer
        End If
    End Sub

    Private Sub spyTimer_Tick() Handles spyTimer.Tick
        Dim tempNewDate As Date
        tempNewDate = DownloadCurrentDate()
        If tempNewDate.Date <> currentDate.Date Then ' it is a new day!
            currentDate = tempNewDate
            ShowSeconds(60)
            RunDailyRoutine()
        End If
    End Sub

End Module
