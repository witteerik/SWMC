


''' <summary>
''' A Windows-forms control that can be used to indicate the progress of time consuming operations. In order to inactivate this control
''' (for instance when run on a web-server), the global variable BlockProgressForm can be set to TRUE.
''' </summary>
Public Class ProgressDisplay

    Public Property MaxValue As Double
    Public Property MinValue As Double
    Public Property Title As String
    'Public WithEvents UpdateTimer As New Timers.Timer
    Private IsInitialized As Boolean = False
    Public Property PrivateUpdateInterval As Integer = 1
    Private SpeedUpdateCount As Integer = 0
    Private LastSpeedUpdateTime As DateTime = DateTime.Now
    Private LastUpdateProcessedItemsCount As Integer = 0

    Private Sub ProgressDisplay_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Shadows Sub Show()

        If BlockProgressForm = True Then Exit Sub
        MyBase.Show()

    End Sub


    Public Shadows Sub Close()

        If BlockProgressForm = True Then Exit Sub
        MyBase.Close()

    End Sub

    Public Sub UpdateProgress_old(ByRef NewProgressValue As Double, ByRef SetMaxValue As Double, Optional ByRef SetMinValue As Double = 0, Optional SetTitle As String = "Working...")

        If BlockProgressForm = True Then Exit Sub

        ProgressBar1.Minimum = MinValue
        ProgressBar1.Maximum = MaxValue
        ProgressBar1.Value = NewProgressValue

        ProgressLabel.Text = "Processing item " & NewProgressValue & " of " & MaxValue - MinValue
        ProgressLabel.Refresh()
        Me.Text = SetTitle

    End Sub

    Public Sub UpdateProgress(NewProgressValue As Double, Optional ByRef NewMaxValue As Double = Nothing, Optional AddToTitle As String = "",
                              Optional AutoCorrectRange As Boolean = True)

        If BlockProgressForm = True Then Exit Sub

        If IsInitialized = False Then
            Throw New Exception("The ProgressDisplay is not initialized. Make sure to call method Initialize before using UpdateProgress.")
        End If

        If NewProgressValue Mod PrivateUpdateInterval = 0 Then

            If Not NewMaxValue = Nothing Then
                MaxValue = NewMaxValue
                ProgressBar1.Maximum = NewMaxValue
            End If

            'Auto correcting values outside of range
            If AutoCorrectRange = True Then
                If NewProgressValue > ProgressBar1.Maximum Then ProgressBar1.Maximum = NewProgressValue
                If NewProgressValue < ProgressBar1.Minimum Then ProgressBar1.Minimum = NewProgressValue
            End If

            ProgressBar1.Value = NewProgressValue
            ProgressLabel.Text = "Processing item " & NewProgressValue & " of " & MaxValue - MinValue & " " & AddToTitle

            ProgressLabel.Refresh()

            Me.Invalidate()
            Me.Refresh()

        End If

    End Sub

    Public Sub UpdateSpeedLabel(ByVal ProcessedItemCount As Integer, ByRef DescriptionUpdateInterval As Integer,
                                Optional ByRef ItemDescription As String = "items", Optional ShowEstimatedProcessingTime As Boolean = True,
                                Optional ByRef NewMaxValue As Double = Nothing)

        If BlockProgressForm = True Then Exit Sub

        If IsInitialized = False Then
            Throw New Exception("The ProgressDisplay is not initialized. Make sure to call method Initialize before using UpdateSpeedLabel.")
        End If

        If SpeedUpdateCount Mod DescriptionUpdateInterval = 0 Then

            'Updating the maxvalue
            If Not NewMaxValue = Nothing Then
                MaxValue = NewMaxValue
            End If

            'Getting the temporal interval
            Dim CurrentTimeSpan As TimeSpan = DateTime.Now - LastSpeedUpdateTime
            Dim TemporalInterval As Double = CurrentTimeSpan.TotalSeconds

            'Getting the number of processed items since last update
            If LastUpdateProcessedItemsCount > ProcessedItemCount Then LastUpdateProcessedItemsCount = 0 'Resetting LastUpdateProcessedItemsCount if a new count has started, and LastUpdateProcessedItemsCount is not reset manually
            Dim ProcessedItems As Integer = ProcessedItemCount - LastUpdateProcessedItemsCount
            Dim ProcessingSpeed As Double = Math.Round(ProcessedItems / TemporalInterval, 0)


            If ShowEstimatedProcessingTime = False Or ProcessingSpeed = 0 Or MaxValue = Nothing Then
                SpeedLabel.Text = "Processing " & ProcessingSpeed & " " & ItemDescription & " / second"
            Else
                Dim EstimatedTimeLeft As Integer = (MaxValue - ProcessedItemCount) / ProcessingSpeed
                SpeedLabel.Text = "Processing " & ProcessingSpeed & " " & ItemDescription & " / second. Estimated time left: " & Int(EstimatedTimeLeft / 60) & " min " & EstimatedTimeLeft - (60 * Int(EstimatedTimeLeft / 60)) & " sec"
            End If

            SpeedLabel.Refresh()

            'Storing last update info
            LastSpeedUpdateTime = DateTime.Now
            LastUpdateProcessedItemsCount = ProcessedItemCount

        End If

        SpeedUpdateCount += 1

        Me.Invalidate()
        Me.Refresh()

    End Sub

    ''' <summary>
    ''' Hi-jacking the speed label to show additional information.
    ''' </summary>
    ''' <param name="Information"></param>
    Public Sub UpdateExtraInfoLabel(ByRef Information As String)

        SpeedLabel.Visible = True
        SpeedLabel.Text = Information
        SpeedLabel.Refresh()

    End Sub


    Public Sub ResetProcessingSpeed()

        If BlockProgressForm = True Then Exit Sub

        LastUpdateProcessedItemsCount = 0
        LastSpeedUpdateTime = DateTime.Now

    End Sub

    Public Sub Initialize(ByRef SetMaxValue As Double, Optional ByRef SetMinValue As Double = 0,
                          Optional SetTitle As String = "Working...", Optional UpdateInterval As Integer = 1,
                          Optional SetDescriptionText As String = "",
                          Optional TopTopMostForm As Boolean = True)

        If BlockProgressForm = True Then Exit Sub

        Me.TopMost = TopTopMostForm

        If SetMaxValue < 0 Then SetMaxValue = 0
        If SetMinValue < 0 Then SetMinValue = 0

        Title = SetTitle
        MinValue = SetMinValue
        MaxValue = SetMaxValue

        Me.Text = Title
        ProgressBar1.Minimum = MinValue
        ProgressBar1.Maximum = MaxValue

        ProgressIndicator = 0

        ProgressLabel.Text = ""
        SpeedLabel.Text = ""

        PrivateUpdateInterval = UpdateInterval

        IsInitialized = True

    End Sub


End Class

