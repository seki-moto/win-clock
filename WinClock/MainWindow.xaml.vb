Imports System.Windows.Threading

Class MainWindow

    Private WithEvents ClockTimer As DispatcherTimer
    Private dateFormat As String
    Private timeFormat As String

    Private Sub MainWindow_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        dateFormat = Me.txtDate.Text
        timeFormat = Me.txtTime.Text

        ClockTimer = New DispatcherTimer()
        ClockTimer.Interval = New TimeSpan(0, 0, 0, 0, 200)
        Call RefreshClock()

        Me.Top = My.Settings.PosTop
        Me.Left = My.Settings.PosLeft
        Me.Background.Opacity = My.Settings.Opacity
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ClockTimer.Start()
    End Sub

    Private Sub MainWindow_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        My.Settings.PosTop = Me.Top
        My.Settings.PosLeft = Me.Left
        My.Settings.Opacity = Me.Background.Opacity
        My.Settings.Save()
    End Sub

    Private Sub MainWindow_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseLeftButtonDown
        Me.DragMove()
    End Sub

    Private Sub MainWindow_MouseWheel(sender As Object, e As MouseWheelEventArgs) Handles Me.MouseWheel
        If Keyboard.Modifiers = ModifierKeys.Control + ModifierKeys.Shift Then
            If e.Delta > 0 AndAlso Me.Background.Opacity > 0.1 Then
                Me.Background.Opacity -= 0.02
            ElseIf e.Delta < 0 AndAlso Me.Background.Opacity < 1 Then
                Me.Background.Opacity += 0.02
            End If
        End If
    End Sub

    Private Sub MainWindow_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseDoubleClick
        If e.ChangedButton.HasFlag(MouseButton.Left) Then
            Me.Close()
        End If
    End Sub

    Private Sub ClockTimer_Tick(sender As Object, e As EventArgs) Handles ClockTimer.Tick
        Call RefreshClock()
    End Sub

    Private Sub RefreshClock()
        Dim _now As Date = Now()
        Me.txtDate.Text = _now.ToString(dateFormat)
        Me.txtTime.Text = _now.ToString(timeFormat)
    End Sub
End Class
