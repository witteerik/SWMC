
Imports System.IO
Imports System.Windows.Forms
Imports System.Threading


''' <summary>
'''Hard coded information on the size of text corpus behind the AfcList frequency data.
''' </summary>
Public Module AfcListCorpusInfo
    Public Const AfcListCorpusTokenCount As Integer = 536866005
    Public Const AfcListCorpusWordTypeCount As Integer = 3591552
    Public Const AfcListCorpusSentenceCount As Integer = 34280136
    Public Const AfcListCorpusDocumentCount As Integer = 11154

    ''' <summary>
    ''' Returns the string used as the first line in the WordGroup function to export .txt file.
    ''' </summary>
    ''' <returns></returns>
    Public Function GetAfcListCorpusDataString() As String
        Return "AfcListCorpusTokenCount=" & AfcListCorpusTokenCount & vbTab &
            "AfcListCorpusWordTypeCount=" & AfcListCorpusWordTypeCount & vbTab &
            "AfcListCorpusSentenceCount=" & AfcListCorpusSentenceCount & vbTab &
            "AfcListCorpusDocumentCount=" & AfcListCorpusDocumentCount
    End Function

End Module


Public Module AfcListMySqlDatabaseInfo

    Public AfcListTableName As String = "AfcList"
    Public AfcListMySqlConnectionString As String

    ''' <summary>
    ''' Call this method to create an AfcListMySqlConnectionString from a .txt file containing the database connection information
    ''' </summary>
    ''' <param name="DatabaseInfoFilePath"></param>
    Public Function LoadAfcListMySqlDatabaseInfo(ByVal DatabaseInfoFilePath As String) As String

        Dim InputData() As String = System.IO.File.ReadAllLines(DatabaseInfoFilePath, Text.Encoding.UTF8)

        Dim InputDictionary As New SortedList(Of String, String)
        For LineIndex = 0 To InputData.Length - 1
            Dim CurrentLine As String = InputData(LineIndex).Trim
            If CurrentLine = "" Then Continue For
            Dim CurrentLineSplit() As String = CurrentLine.Split("=")

            Dim CurrentKey As String = CurrentLineSplit(0).Trim
            If CurrentKey = "" Then Continue For

            If CurrentLineSplit.Length < 2 Then Continue For

            Dim CurrentValue As String = CurrentLineSplit(1).Trim
            If CurrentValue = "" Then Continue For

            If Not InputDictionary.ContainsKey(CurrentKey) Then
                InputDictionary.Add(CurrentKey, CurrentValue)
            End If
        Next

        Dim mysql_host As String = InputDictionary("mysql_host")
        Dim mysql_user As String = InputDictionary("mysql_user")
        Dim mysql_password As String = InputDictionary("mysql_password")
        Dim mysql_database As String = InputDictionary("mysql_database")

        'AfcListMySqlConnectionString = "server=" + mysql_host + ";uid=" + mysql_user + ";pwd=" + mysql_password + ";database=" + mysql_database + ";"

        AfcListMySqlConnectionString = "server=" + mysql_host + ";uid=" + mysql_user + ";pwd=" + mysql_password + ";database=" + mysql_database + "; Character Set=utf8" & ";"

        Return AfcListMySqlConnectionString

    End Function


    ''' <summary>
    ''' Call this method to set (and return) the default AfcListMySqlConnectionString hard-coded within the library
    ''' </summary>
    Public Function LoadDefaultAfcListMySqlDatabaseInfo() As String

        Dim mysql_server As String = "mysql682.loopia.se"
        Dim mysql_uid As String = "swmwp@s258672"
        Dim mysql_pwd As String = "hJe37s20Vb"
        Dim mysql_database As String = "swedishwordmetrics_com_db_1"

        'Set the global variable AfcListMySqlConnectionString to prepares for server connection
        AfcListMySqlConnectionString = "server=" + mysql_server + ";uid=" + mysql_uid + ";pwd=" + mysql_pwd + ";database=" + mysql_database + "; Character Set=utf8" & ";"

        Return AfcListMySqlConnectionString

    End Function

End Module



Public Module GlobalObjects

    'Variables that can be used with the Windows forms type ProgressDisplay 
    Public ProgressIndicator As Integer
    Public BlockProgressForm As Boolean = False

    'Variables used for logging support
    ''' <summary>
    ''' Can be used to block all logging, for example when run on a web-server. But rather than blocking from within SendInfoToLog, logging should be blocked
    ''' from the calling code.
    ''' </summary>
    Public GeneralLogIsActive As Boolean = True
    Public logFilePath As String = "C:\AfcMetricsLog\"
    Public showErrors As Boolean = True
    Public logErrors As Boolean = True
    Public logIsInMultiThreadApplication As Boolean = False
    Public LoggingSpinLock As New Threading.SpinLock


    ' Objects that hold OLD and PLD comparison-corpus data
    Public OLD_corpus() As String
    Public PLD_IPA_Corpus() As String

    ' Special characters
    Public ZeroPhoneme As String = "∅"
    Public PhoneticLength As String = "ː"
    Public IpaMainStress As String = "ˈ"
    Public IpaMainSwedishAccent2 As String = "²"
    Public IpaSecondaryStress As String = "ˌ"
    Public IpaSyllableBoundary As String = "."
    Public SwedishStressList As New SortedSet(Of String) From {IpaMainStress, IpaMainSwedishAccent2, IpaSecondaryStress}
    Public AmbiguosOnsetMarker As String = "!"
    Public AmbiguosCodaMarker As String = "*"
    Public SwedishOrthographicCharacters As New List(Of String) From {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "å", "ä", "ö", "é"} ' borde <à> vara med? Troligen inte!?
    Public SwedishConsonants_IPA As New List(Of String) From {“ɕ", "ŋ", "ɧ", "s", "v", "m", "ɱ", "n", "d", "ɡ", "k", "r", "ʝ", "b", "f", "h", "l", "p", "t", "ʂ", "ʈ", "ɖ", "ɭ", "ɳ", "ŋː", "ɧː", "sː", "vː", "mː", "ɱː", "nː", "dː", "ɡː", "kː", "rː", "ʝː", "bː", "fː", "lː", "pː", "tː", "ʂː", "ʈː", "ɖː", "ɭː", "ɳː"}
    Public SwedishVowels_IPA As New List(Of String) From {“iː", "ɪ", "yː", "ʏ", "eː", "e̞", "ə", "øː", "ø̞", "ɛː", "ɛ̝", "ʉː", "ɵ", "ʉ̞", "ɑː", "a", "uː", "ʊ", "oː", "ɔ", "a͡uː", "e͡ʉː", “i", "y", "e", "ø", "ɛ", "ʉ", "ɑ", "u", "o", "a͡u", "e͡ʉ"}
    Public VowelAllophoneReductionList As New Dictionary(Of String, String) From {{"i", "ɪ"}, {"y", "ʏ"}, {"ø", "ø̞"}, {"e", "e̞"}, {"ɛ", "ɛ̝"}, {"u", "ʊ"}, {"o", "ɔ"}, {"ʉ", "ʉ̞"}}
    Public SwedishShortVowels_IncludingLongReduced_IPA As New List(Of String) From {"ɪ", "ʏ", "e̞", "ə", "ø̞", "œ", "ɛ̝", "æ", "ɵ", "ʉ̞", "a", "ʊ", "ɔ", “i", "y", "e", "ø", "œ", "ɛ", "æ", "ʉ", "ɑ", "u", "o", "a͡u", "e͡ʉ"}
    Public SwedishShortVowels_IPA As New List(Of String) From {"ɪ", "ʏ", "e̞", "ə", "ø̞", "œ", "ɛ̝", "æ", "ɵ", "ʉ̞", "a", "ʊ", "ɔ"}
    Public SwedishLongVowels_IPA As New List(Of String) From {“iː", "yː", "eː", "øː", "œː", "ɛː", "æː", "ʉː", "ɑː", "uː", "oː", "a͡uː", "e͡ʉː"}
    Public AllSuprasegmentalIPACharacters As New List(Of String) From {"_", "¤", "ˈ", "²", "ˌ", "."}

End Module


Public Module MathMethods

    ''' <summary>
    ''' Calculates Zipf Value from raw word type frequency.
    ''' </summary>
    ''' <param name="RawWordTypeFrequency">The total number of times a tokens exists in the corpus used.</param>
    ''' <param name="CorpusTotalTokenCount">The total number of tokens in the corpus used to set the raw word type frequency.</param>
    ''' <param name="CorpusTotalWordTypeCount">The total number of word types in the corpus used to set the raw word type frequency.</param>
    ''' <param name="PositionTerm"></param>
    Public Function CalculateZipfValue(ByRef RawWordTypeFrequency As Long, ByVal CorpusTotalTokenCount As Long, ByVal CorpusTotalWordTypeCount As Integer, Optional ByVal PositionTerm As Integer = 3)

        Return Math.Log10((RawWordTypeFrequency + 1) / ((CorpusTotalTokenCount + CorpusTotalWordTypeCount) / 1000000)) + PositionTerm

    End Function


    Public Enum StandardDeviationTypes
        Population
        Sample
    End Enum

    ''' <summary>
    ''' Standardizes the values in the input list.
    ''' </summary>
    ''' <param name="InputValueType">Default calculation type (Population) uses N in the variance calculation denominator. If Sample type is used, the denominator is N-1.</param>
    ''' <param name="SetMeanTo">An optional term that can be used to adjust the mean to a desired value. Can be used to avoid negative values within the distribution.</param>
    ''' <param name="ExcludeNegativeValues">If set to true, all negative values will be ignored when calculating mean and standard deviation. All negative values will be standardized based on the mean and standard deviation of non negative values.</param>
    ''' <param name="InputListOfDouble"></param>
    ''' <param name="SetToZeroOnNoVariance">Sets all values in the input list to zero if no variance is detected.</param>
    Public Sub Standardization(ByRef InputListOfDouble As List(Of Double),
                              Optional ByRef SetMeanTo As Double = 0,
                               Optional ByRef ExcludeNegativeValues As Boolean = False,
                               Optional ByRef InputValueType As StandardDeviationTypes = StandardDeviationTypes.Population,
                               Optional ByVal SetToZeroOnNoVariance As Boolean = True)


        'Notes the number of values in the input list
        'Dim n As Integer = InputListOfDouble.Count

        'Calculates the sum of the values in the input list
        Dim Sum As Double = 0
        Dim n As Integer = 0
        For i = 0 To InputListOfDouble.Count - 1
            If ExcludeNegativeValues = True Then
                If InputListOfDouble(i) >= 0 Then
                    Sum += InputListOfDouble(i)
                    n += 1
                End If
            Else
                Sum += InputListOfDouble(i)
                n += 1
            End If
        Next

        'Calculates the arithemtric mean of the values in the input list
        Dim ArithmetricMean As Double = Sum / n

        'Calculates the sum of squares of the values in the input list
        Dim SumOfSquares As Double = 0
        Dim n_SumOfSquares As Integer = 0
        For i = 0 To InputListOfDouble.Count - 1
            If ExcludeNegativeValues = True Then
                If InputListOfDouble(i) >= 0 Then
                    SumOfSquares += (InputListOfDouble(i) - ArithmetricMean) ^ 2
                End If
            Else
                SumOfSquares += (InputListOfDouble(i) - ArithmetricMean) ^ 2
            End If
        Next

        'Calculates the variance of the values in the input list
        Dim Variance As Double
        Select Case InputValueType
            Case StandardDeviationTypes.Population
                Variance = (1 / (n)) * SumOfSquares
            Case StandardDeviationTypes.Sample
                Variance = (1 / (n - 1)) * SumOfSquares
        End Select

        'Setting to zero on no variance, then exits sub (The reason this is needed is that we will get the square root of 0 in the next step if variance is 0.)
        If Variance = 0 Then
            If SetToZeroOnNoVariance = True Then
                For n = 0 To InputListOfDouble.Count - 1
                    InputListOfDouble(n) = 0
                Next
                Exit Sub
            End If
        End If

        'Calculates, the standard deviation of the values in the input list
        Dim StandardDeviation As Double = Math.Sqrt(Variance)

        'Standardizes the values in the input list
        For n = 0 To InputListOfDouble.Count - 1
            InputListOfDouble(n) = ((InputListOfDouble(n) - ArithmetricMean) / StandardDeviation) + SetMeanTo
        Next

    End Sub


    Public Enum roundingMethods
        getClosestValue
        alwaysDown
        alwaysUp
        donNotRound
    End Enum

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="inputValue"></param>
    ''' <param name="roundingMethod"></param>
    ''' <param name="DecimalsInReturnsString"></param>
    ''' <param name="SkipRounding">If set to true, the rounding function is inactivated and the input value is returned unalterred.</param>
    ''' <returns></returns>
    Public Function Rounding(ByVal inputValue As Object, Optional ByVal roundingMethod As roundingMethods = roundingMethods.getClosestValue,
                             Optional DecimalsInReturnsString As Integer? = Nothing, Optional ByVal SkipRounding As Boolean = False,
                             Optional MinimumNonDecimalsInReturnString As Integer? = Nothing)

        'Returns the input value, if SkipRounding is true
        If SkipRounding = True Then Return inputValue

        Try

            Dim ReturnValue As Double = inputValue

            Select Case roundingMethod
                Case roundingMethods.alwaysDown
                    ReturnValue = Int(ReturnValue)

                Case roundingMethods.alwaysUp
                    If Not inputValue - Int(inputValue) = 0 Then
                        ReturnValue = Int(ReturnValue) + 1
                    Else
                        ReturnValue = ReturnValue
                    End If

                Case roundingMethods.donNotRound
                    ReturnValue = ReturnValue

                Case roundingMethods.getClosestValue

                    If DecimalsInReturnsString Is Nothing Then
                        ReturnValue = (Math.Round(ReturnValue))
                        'If not midpoint rounding is done below
                    End If

                Case Else
                    Throw New Exception("The " & roundingMethod & " rounding method enumerator is not valid.")
                    Return Nothing
            End Select

            Dim RetString As String = ""
            If DecimalsInReturnsString IsNot Nothing Or MinimumNonDecimalsInReturnString IsNot Nothing Then

                If DecimalsInReturnsString < 0 Then Throw New ArgumentException("DecimalsInReturnsString cannot be lower than 0.")
                If MinimumNonDecimalsInReturnString < 0 Then Throw New ArgumentException("MinimumNonDecimalsInReturnString cannot be lower than 0.")

                'Adding decimals to format
                Dim NumberFormat As String = "0"
                If DecimalsInReturnsString IsNot Nothing Then
                    For n = 0 To DecimalsInReturnsString - 1
                        If n = 0 Then NumberFormat &= "."
                        NumberFormat &= "0"
                    Next
                End If

                'Adding non-decimals to format
                If MinimumNonDecimalsInReturnString IsNot Nothing Then
                    For n = 0 To MinimumNonDecimalsInReturnString - 2 ' -2 as one 0 has already been added above
                        NumberFormat = "0" & NumberFormat
                    Next
                End If

                RetString = ReturnValue.ToString(NumberFormat).TrimEnd("0").Trim(".").Trim(",")
                If RetString = "" Then RetString = "0"

                Return RetString

            End If

            Return ReturnValue

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return Nothing
        End Try

    End Function




    ''' <summary>
    ''' Returns the geometric mean in a list of Doubles.
    ''' </summary>
    ''' <param name="InputList"></param>
    ''' <returns>Returns the geometric mean of the vaules in the input list.</returns>
    Public Function GeometricMean(InputList As List(Of Double))

        Dim n As Integer = InputList.Count

        Dim ListProduct As Double = 1
        For i = 0 To InputList.Count - 1
            ListProduct *= InputList(i)
        Next

        Return ListProduct ^ (1 / n)

    End Function



    ''' <summary>
    ''' Calculates edit distance using a Levenshtein type algorithm, based on code example in Jurafsky and Martin (p 33), and slightly modified to increase performance.
    ''' </summary>
    ''' <param name="Target"></param>
    ''' <param name="Source"></param>
    ''' <returns></returns>
    Function LevenshteinDistance(ByVal Target As List(Of String), ByVal Source As List(Of String)) As Integer


        'Setting up the distance matrix
        Dim n As Integer = Target.Count
        Dim m As Integer = Source.Count
        Dim DistanceMatrix(n + 1, m + 1) As Integer
        DistanceMatrix(0, 0) = 0

        Dim InsertionCost As Integer = 1
        Dim SubstitutionCost As Integer = 1
        Dim DeletionCost As Integer = 1

        For Column_i = 1 To n
            DistanceMatrix(Column_i, 0) = DistanceMatrix(Column_i - 1, 0) + InsertionCost
        Next
        For Row_j = 1 To m
            DistanceMatrix(0, Row_j) = DistanceMatrix(0, Row_j - 1) + DeletionCost
        Next

        Dim nMinus1 As Integer = n - 1
        Dim mMinus1 As Integer = m - 1

        For i = 0 To nMinus1
            For j = 0 To mMinus1

                'Sets the SubstitutionCost to 0 if the compared characters are the same
                If Source(j) = Target(i) Then
                    SubstitutionCost = 0
                Else
                    SubstitutionCost = 1
                End If

                DistanceMatrix(i + 1, j + 1) = Math.Min(Math.Min(DistanceMatrix(i, j + 1) + InsertionCost,
                            DistanceMatrix(i, j) + SubstitutionCost),
                            DistanceMatrix(i + 1, j) + DeletionCost)
            Next
        Next

        Return (DistanceMatrix(n, m))

    End Function


    Function LevenshteinDistance(ByVal String1 As String, ByVal String2 As String) As Integer

        'The VB.NET code in this sub originates from http://rosettacode.org/wiki/Levenshtein_distance 
        'It is however corrected by me in the last line which was originally "Return Matrix(String1.Length -1, String2.Length -1)"
        'It is also slightly modified to increase performance

        Dim Matrix(String1.Length, String2.Length) As Integer

        Matrix(0, 0) = 0
        For Key = 1 To String1.Length
            Matrix(Key, 0) = Key
        Next
        For Key = 1 To String2.Length
            Matrix(0, Key) = Key
        Next

        For Key1 As Integer = 1 To String2.Length
            For Key2 As Integer = 1 To String1.Length
                If String1(Key2 - 1) = String2(Key1 - 1) Then
                    Matrix(Key2, Key1) = Matrix(Key2 - 1, Key1 - 1) 'no operation
                Else
                    Matrix(Key2, Key1) = Math.Min(Matrix(Key2 - 1, Key1) + 1,
                                                          Math.Min(Matrix(Key2, Key1 - 1) + 1,
                                                                   Matrix(Key2 - 1, Key1 - 1) + 1))
                End If
            Next
        Next

        Return Matrix(String1.Length, String2.Length)
    End Function


End Module






Public Module LoggingMethods


    Public Sub SendInfoToLog(ByVal message As String,
                             Optional ByVal LogFileNameWithoutExtension As String = "",
                             Optional LogFileTemporaryPath As String = "",
                             Optional ByVal OmitDateInsideLog As Boolean = False,
                             Optional ByVal OmitDateInFileName As Boolean = False)
        'Optional ByRef SpinLock As Threading.SpinLock = Nothing)

        Dim SpinLockTaken As Boolean = False

        Try

            'Blocks logging if GeneralLogIsActive is False
            If GeneralLogIsActive = False Then Exit Sub

            'Attempts to enter a spin lock to avoid multiple thread conflicts when saving to the same file
            LoggingSpinLock.Enter(SpinLockTaken)

            If LogFileTemporaryPath = "" Then LogFileTemporaryPath = logFilePath

            Dim FileNameToUse As String = ""

            If OmitDateInFileName = False Then
                If LogFileNameWithoutExtension = "" Then
                    FileNameToUse = "log-" & DateTime.Now.ToShortDateString.Replace("/", "-") & ".txt"
                Else
                    FileNameToUse = LogFileNameWithoutExtension & "-" & DateTime.Now.ToShortDateString.Replace("/", "-") & ".txt"
                End If
            Else
                If LogFileNameWithoutExtension = "" Then
                    FileNameToUse = "log.txt"
                Else
                    FileNameToUse = LogFileNameWithoutExtension & ".txt"
                End If

            End If

            Dim OutputFilePath As String = Path.Combine(LogFileTemporaryPath, FileNameToUse)

            'Adds a thread ID if in multi thread app
            If logIsInMultiThreadApplication = True Then
                Dim TreadName As String = Thread.CurrentThread.ManagedThreadId
                OutputFilePath &= "ThreadID_" & TreadName
            End If

            Try
                'If File.Exists(logFilePathway) Then File.Delete(logFilePathway)
                If Not Directory.Exists(LogFileTemporaryPath) Then Directory.CreateDirectory(LogFileTemporaryPath)
                Dim samplewriter As New StreamWriter(OutputFilePath, FileMode.Append)
                If OmitDateInsideLog = False Then
                    samplewriter.WriteLine(DateTime.Now.ToString & vbCrLf & message)
                Else
                    samplewriter.WriteLine(message)
                End If
                samplewriter.Close()

            Catch ex As Exception
                Errors(ex.ToString, "Error saving to log file!")
            End Try

        Finally

            'Releases any spinlock
            If SpinLockTaken = True Then LoggingSpinLock.Exit()
        End Try

    End Sub

    Public Sub Errors(ByVal errorText As String, Optional ByVal errorTitle As String = "Error")

        If showErrors = True Then
            MsgBox(errorText, MsgBoxStyle.Critical, errorTitle)
        End If

        If logErrors = True Then
            SendInfoToLog("The following error occurred: " & vbCrLf & errorTitle & errorText, "Errors")
        End If

    End Sub


End Module


Public Module WindowsFormsCustomFileDialogs


    ''' <summary>
    ''' Asks the user to supply a file path by using a save file dialog box.
    ''' </summary>
    ''' <param name="directory">Optional initial directory.</param>
    ''' <param name="fileName">Optional suggested file name</param>
    ''' <param name="fileExtensions">Optional possible extensions</param>
    ''' <param name="BoxTitle">The message/title on the file dialog box</param>
    ''' <returns>Returns the file path, or nothing if a file path could not be created.</returns>
    Public Function GetSaveFilePath(Optional ByRef directory As String = "",
                                Optional ByRef fileName As String = "",
                                    Optional fileExtensions() As String = Nothing,
                                    Optional BoxTitle As String = "") As String

        Dim filePath As String = ""
        'Asks the user for a file path using the SaveFileDialog box.
SavingFile: Dim sfd As New SaveFileDialog

        'Creating a filterstring
        If fileExtensions IsNot Nothing Then
            Dim filter As String = ""
            For ext = 0 To fileExtensions.Length - 1

                filter &= fileExtensions(ext).Trim(".") & " files (*." & fileExtensions(ext).Trim(".") & ")|*." & fileExtensions(ext).Trim(".") & "|"
            Next
            filter = filter.TrimEnd("|")
            sfd.Filter = filter
        End If

        If Not directory = "" Then sfd.InitialDirectory = directory
        If Not fileName = "" Then sfd.FileName = fileName
        If Not BoxTitle = "" Then sfd.Title = BoxTitle

        Dim result As DialogResult = sfd.ShowDialog()
        If result = DialogResult.OK Then
            filePath = sfd.FileName

            'Updats input variables to contain the selected
            directory = Path.GetDirectoryName(filePath)
            fileName = Path.GetFileName(filePath)

            Return filePath
        Else
            Dim errorSaving As MsgBoxResult = MsgBox("An error occurred choosing file name.", MsgBoxStyle.RetryCancel, "Warning!")
            If errorSaving = MsgBoxResult.Retry Then
                GoTo SavingFile
            Else
                Return Nothing
            End If
        End If

    End Function

    ''' <summary>
    ''' Asks the user to supply a file path by using an open file dialog box.
    ''' </summary>
    ''' <param name="directory">Optional initial directory.</param>
    ''' <param name="fileName">Optional suggested file name</param>
    ''' <param name="fileExtensions">Optional possible extensions</param>
    ''' <param name="BoxTitle">The message/title on the file dialog box</param>
    ''' <returns>Returns the file path, or nothing if a file path could not be created.</returns>
    Public Function GetOpenFilePath(Optional directory As String = "",
                                    Optional fileName As String = "",
                                    Optional fileExtensions() As String = Nothing,
                                    Optional BoxTitle As String = "",
                                    Optional ReturnEmptyStringOnCancel As Boolean = False) As String

        Dim filePath As String = ""

SavingFile: Dim ofd As New OpenFileDialog
        'Creating a filterstring
        If fileExtensions IsNot Nothing Then
            Dim filter As String = ""
            For ext = 0 To fileExtensions.Length - 1
                filter &= fileExtensions(ext).Trim(".") & " files (*." & fileExtensions(ext).Trim(".") & ")|*." & fileExtensions(ext).Trim(".") & "|"
            Next
            filter = filter.TrimEnd("|")
            ofd.Filter = filter
        End If

        If Not directory = "" Then ofd.InitialDirectory = directory
        If Not fileName = "" Then ofd.FileName = fileName
        If Not BoxTitle = "" Then ofd.Title = BoxTitle

        Dim result As DialogResult = ofd.ShowDialog()
        If result = DialogResult.OK Then
            filePath = ofd.FileName
            Return filePath
        Else
            'Returns en empty string if cancel was pressed and ReturnEmptyStringOnCancel = True 
            If ReturnEmptyStringOnCancel = True Then Return ""

            Dim boxResult As MsgBoxResult = MsgBox("An error occurred choosing file name.", MsgBoxStyle.RetryCancel, "Warning!")
            If boxResult = MsgBoxResult.Retry Then
                GoTo SavingFile
            Else
                Return Nothing
            End If
        End If

    End Function


End Module













