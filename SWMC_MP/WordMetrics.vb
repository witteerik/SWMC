
Imports System.Data
Imports MySqlConnector


Public Class WordMetricsCalculation

    Private CustumSpeechDataLocation As String
    Private ValidPhoneticCharacters As List(Of String)


    ''' <summary>
    ''' Creates a new instance of WordMetricsCalculation, and prepares for SQL server connection.
    ''' </summary>
    ''' <param name="CustomSpeechDataLocation">The location of custom speech data read from text files. Leave to "" to use the embedded speech data.</param>
    Public Sub New(ByVal mysql_server As String, ByVal mysql_uid As String, ByVal mysql_pwd As String, ByVal mysql_database As String,
                   Optional ByVal CustomSpeechDataLocation As String = "")

        Me.CustumSpeechDataLocation = CustomSpeechDataLocation

        'Set the global variable AfcListMySqlConnectionString to prepares for server connection
        AfcListMySqlConnectionString = "server=" + mysql_server + ";uid=" + mysql_uid + ";pwd=" + mysql_pwd + ";database=" + mysql_database + "; Character Set=utf8" & ";"

        'Creating list of valid phonetic characters
        ValidPhoneticCharacters = CreateListOfValidPhoneticCharactersForSwedish()

    End Sub

    ''' <summary>
    ''' Creates a new instance of WordMetricsCalculation, and prepares for the default SQL server connection.
    ''' </summary>
    ''' <param name="CustomSpeechDataLocation">The location of custom speech data read from text files. Leave to "" to use the embedded speech data.</param>
    Public Sub New(Optional ByVal SetupDefaultSqlConnection As Boolean = True, Optional ByVal CustomSpeechDataLocation As String = "")

        Me.CustumSpeechDataLocation = CustomSpeechDataLocation

        'Set the global variable AfcListMySqlConnectionString to prepares for server connection
        If SetupDefaultSqlConnection = True Then LoadDefaultAfcListMySqlDatabaseInfo()

        'Creating list of valid phonetic characters
        ValidPhoneticCharacters = CreateListOfValidPhoneticCharactersForSwedish()

    End Sub

    ''' <summary>
    ''' Calculates the selected word metrics from a string array of orthographic words, or orthographic word and phonetic transcriptions. 
    ''' If a phonetic transcription is not supplied an attempt is made at getting the corresponding AFC list phonetic transcription/s.
    ''' </summary>
    ''' <param name="InputWords"></param>
    ''' <param name="CalculateOrthographicTransparency"></param>
    ''' <param name="UseTempSpellingChanges"></param>
    ''' <param name="TemporarySpellingChangeListArray"></param>
    ''' <param name="CalculatePhonotacticProbability"></param>
    ''' <param name="CalculatePhoneticNeighborhoodDensity"></param>
    ''' <param name="ListOfInvalidPhoneticCharacterWords"></param>
    ''' <param name="OtherPhoneticErrorWords"></param>
    ''' <param name="CorrectDoubleSpacesInTranscription"></param>
    ''' <param name="CheckPhonemeValidity"></param>
    ''' <param name="CheckTranscriptionStructure"></param>
    ''' <param name="p2g_Settings"></param>
    ''' <param name="Unresolved_p2g_Character"></param>
    ''' <param name="FatalErrorDescription"></param>
    ''' <param name="DoAfcListLookup"></param>
    ''' <param name="DontExportAnything"></param>
    ''' <param name="IsRunFromServer"></param>
    ''' <param name="UseLocalAccdb">Set to true to use a local Access database file. The file must be in the format . accdb, reside in the AccdbFileFolder, and have the AccdbFileName file name.</param>
    ''' <param name="AccdbFileFolder">The folder where a local Access database file is stored.</param>
    ''' <param name="AccdbFileName">The file name (including the extension) of a local Access database file is stored.</param>
    ''' <returns></returns>
    Public Function CalculateWordMetrics(ByVal InputWords() As String,
                                         Optional ByVal CalculateOrthographicTransparency As Boolean = True,
                                         Optional ByVal UseTempSpellingChanges As Boolean = False,
                                         Optional ByVal TemporarySpellingChangeListArray() As String = Nothing,
                                         Optional ByVal CalculatePhonotacticProbability As Boolean = True,
                                         Optional ByVal CalculatePhoneticNeighborhoodDensity As Boolean = True,
                                         Optional ByVal CalculateOrthographicNeighborhoodDensity As Boolean = True,
                                         Optional ByVal CalculatePLDx As Boolean = True,
                                         Optional ByVal PLDxCount As Integer = 20,
                                         Optional ByVal CalculateOLDx As Boolean = True,
                                         Optional ByVal OLDxCount As Integer = 20,
                                         Optional ByVal CalculateOrthographicIsolationPoints As Boolean = True,
                                         Optional ByVal CalculatePhoneticIsolationPoints As Boolean = True,
                                         Optional ByVal AddInputWordsToComparisonLists As Boolean = True,
                                         Optional ByRef ListOfInvalidPhoneticCharacterWords As List(Of String) = Nothing,
                                         Optional ByRef OtherPhoneticErrorWords As List(Of String) = Nothing,
                                         Optional ByVal CorrectDoubleSpacesInTranscription As Boolean = True,
                                         Optional ByVal CheckPhonemeValidity As Boolean = True,
                                         Optional ByVal CheckTranscriptionStructure As Boolean = True,
                                         Optional ByVal p2g_Settings As p2gParameters = Nothing,
                                         Optional ByVal Unresolved_p2g_Character As String = "!",
                                         Optional ByRef FatalErrorDescription As String = "",
                                         Optional ByVal DoAfcListLookup As Boolean = True,
                                         Optional ByVal DontExportAnything As Boolean = True,
                                         Optional ByVal IsRunFromServer As Boolean = False) As WordGroup


        'Creates a new list of Tuple of strings, containing spelling and phonetic transcription
        Dim InputSpellingAndTranscription As New List(Of Tuple(Of String, String, Single?))

        For Each InputWord In InputWords
            'Skips to next if InputWord string is empty
            If InputWord = "" Then Continue For

            'Getting the zipf value enterred by the user
            Dim ZipfValue_Word As Single? = Nothing
            If InputWord.Contains("]") Then
                Dim ZipfSplit() As String = InputWord.Split("]")

                'Changing the InputWord string to contain only the content before the ] (first) character
                InputWord = ZipfSplit(0).Trim.Trim(vbTab).Trim 'Trims off both blank spaces and tabs

                Dim ZipfValueString As String = ZipfSplit(ZipfSplit.Length - 1).Trim
                If IsNumeric(ZipfValueString) Then 'N.B. It's not clear here if we need a replacement for commas/dots. I guess it may depend on the configuration of the local computer, and the user should make sure that the correct values are read by looking at the output data.
                    ZipfValue_Word = ZipfValueString
                End If
            End If

            'Getting spelling and transcription, or only spelling
            'Determines if the input string contains a phonetic transcription by detecting a "[" charachter
            Dim PhoneticTranscription As String = ""
            Dim Spelling As String = ""
            If InputWord.Contains("[") Then

                'Gets the transcription from the input string
                Dim InputSplit() As String = InputWord.Split("[")
                Spelling = InputSplit(0).Trim.Trim(vbTab).Trim.ToLower 'Trims off both blank spaces and tabs, and converts to lower
                PhoneticTranscription = InputSplit(1).Trim.TrimEnd("]")

            Else
                'Assuming that the input string is only spelling
                Spelling = InputWord.Trim.ToLower 'converts to lower

            End If

            'Ignores the word if the spelling is empty
            If Spelling = "" Then Continue For

            'Adds the new spelling/transcription/ZipfValue combination
            InputSpellingAndTranscription.Add(New Tuple(Of String, String, Single?)(Spelling, PhoneticTranscription, ZipfValue_Word))

        Next

        Dim Result = CalculateWordMetrics(InputSpellingAndTranscription.ToArray,
                                         CalculateOrthographicTransparency,
                                         UseTempSpellingChanges,
                                         TemporarySpellingChangeListArray,
                                         CalculatePhonotacticProbability,
                                         CalculatePhoneticNeighborhoodDensity,
                                         CalculateOrthographicNeighborhoodDensity,
                                          CalculatePLDx, PLDxCount,
                                          CalculateOLDx, OLDxCount,
                                         CalculateOrthographicIsolationPoints, CalculatePhoneticIsolationPoints,
                                          AddInputWordsToComparisonLists,
                                         ListOfInvalidPhoneticCharacterWords,
                                         OtherPhoneticErrorWords,
                                         CorrectDoubleSpacesInTranscription,
                                         CheckPhonemeValidity,
                                         CheckTranscriptionStructure,
                                         p2g_Settings,
                                         Unresolved_p2g_Character,
                                         FatalErrorDescription,
                                          DoAfcListLookup,
                                          DontExportAnything,
                                          IsRunFromServer)

        Return Result

    End Function


    ''' <summary>
    ''' Calculates the selected word metrices from an array of orthographic words, phonetic transcriptions and ZipfValues.
    ''' </summary>
    ''' <param name="InputWords"></param>
    ''' <param name="CalculateOrthographicTransparency"></param>
    ''' <param name="UseTempSpellingChanges"></param>
    ''' <param name="TemporarySpellingChangeListArray"></param>
    ''' <param name="CalculatePhonotacticProbability"></param>
    ''' <param name="CalculatePhoneticNeighborhoodDensity"></param>
    ''' <param name="ListOfInvalidPhoneticCharacterWords"></param>
    ''' <param name="OtherPhoneticErrorWords"></param>
    ''' <param name="CorrectDoubleSpacesInTranscription"></param>
    ''' <param name="CheckPhonemeValidity"></param>
    ''' <param name="CheckTranscriptionStructure"></param>
    ''' <param name="p2g_Settings"></param>
    ''' <param name="Unresolved_p2g_Character"></param>
    ''' <param name="FatalErrorDescription"></param>
    ''' <param name="DoAfcListLookup"></param>
    ''' <param name="DontExportAnything"></param>
    ''' <param name="IsRunFromServer"></param>
    ''' <param name="UseLocalAccdb">Set to true to use a local Access database file. The file must be in the format . accdb, reside in the AccdbFileFolder, and have the AccdbFileName file name.</param>
    ''' <param name="AccdbFileFolder">The folder where a local Access database file is stored.</param>
    ''' <param name="AccdbFileName">The file name (including the extension) of a local Access database file is stored.</param>
    ''' <returns></returns>
    Public Function CalculateWordMetrics(ByVal InputWords() As Tuple(Of String, String, Single?),
                                         Optional ByVal CalculateOrthographicTransparency As Boolean = True,
                                         Optional ByVal UseTempSpellingChanges As Boolean = False,
                                         Optional ByRef TemporarySpellingChangeListArray() As String = Nothing,
                                         Optional ByVal CalculatePhonotacticProbability As Boolean = True,
                                         Optional ByVal CalculatePhoneticNeighborhoodDensity As Boolean = True,
                                         Optional ByVal CalculateOrthographicNeighborhoodDensity As Boolean = True,
                                         Optional ByVal CalculatePLDx As Boolean = True,
                                         Optional ByVal PLDxCount As Integer = 20,
                                         Optional ByVal CalculateOLDx As Boolean = True,
                                         Optional ByVal OLDxCount As Integer = 20,
                                         Optional ByVal CalculateOrthographicIsolationPoints As Boolean = True,
                                         Optional ByVal CalculatePhoneticIsolationPoints As Boolean = True,
                                         Optional ByVal AddInputWordsToComparisonLists As Boolean = True,
                                         Optional ByRef ListOfInvalidPhoneticCharacterWords As List(Of String) = Nothing,
                                         Optional ByRef OtherPhoneticErrorWords As List(Of String) = Nothing,
                                         Optional ByRef CorrectDoubleSpacesInTranscription As Boolean = True,
                                         Optional ByRef CheckPhonemeValidity As Boolean = True,
                                         Optional ByRef CheckTranscriptionStructure As Boolean = True,
                                         Optional ByRef p2g_Settings As p2gParameters = Nothing,
                                         Optional ByRef Unresolved_p2g_Character As String = "!",
                                         Optional ByRef FatalErrorDescription As String = "",
                                         Optional ByRef DoAfcListLookup As Boolean = True,
                                         Optional ByRef DontExportAnything As Boolean = True,
                                         Optional ByVal IsRunFromServer As Boolean = False) As WordGroup


        'Making sure we have a list of valid phonetic characters
        If CheckPhonemeValidity = True And ValidPhoneticCharacters Is Nothing Then
            ValidPhoneticCharacters = CreateListOfValidPhoneticCharactersForSwedish()
        End If

        'Resetting the error lists
        ListOfInvalidPhoneticCharacterWords = New List(Of String) From {"The following words contained invalid phonetic characters:"}
        OtherPhoneticErrorWords = New List(Of String) From {"The following errors were found in the phonetic transcriptions:"}

        'Creating a new word group from the input words
        'Looks up the phonetic transcriptions and creates a new word group
        Dim NewWordGroup As New WordGroup
        For Each InputWord In InputWords

            'Skips to next if spelling is empty
            Dim Spelling As String = InputWord.Item1
            If Spelling = "" Then Continue For

            'Creats a new word
            Dim NewWord As New Word With {.OrthographicForm = Spelling}

            'Adds any ZipfValue, added by the user
            Dim ZipfValue As Single? = InputWord.Item3
            If ZipfValue IsNot Nothing Then
                NewWord.ZipfValue_Word = ZipfValue
            End If

            'Adds the transcription if it's not empty
            Dim PhoneticTranscription As String = InputWord.Item2
            If PhoneticTranscription <> "" Then
                'Adds the transcription
                Dim ContainsInvalidPhoneticCharacters As Boolean = False
                Dim CorrectedDoubleSpacesInPhoneticForm As Boolean = False

                NewWord.ParseInputPhoneticString(PhoneticTranscription, ValidPhoneticCharacters,
                                                               ContainsInvalidPhoneticCharacters,
                                                               CorrectDoubleSpacesInTranscription, CorrectedDoubleSpacesInPhoneticForm,
                                                               CheckPhonemeValidity)

                'WordLists.WordListsIO.ParseInputPhoneticString(PhoneticTranscription, NewWord, ValidPhoneticCharacters,
                '                                               ContainsInvalidPhoneticCharacters,
                '                                               CorrectDoubleSpacesInTranscription, CorrectedDoubleSpacesInPhoneticForm,
                '                                               CheckPhonemeValidity)

                'Checking for valid phonetic characters
                If ContainsInvalidPhoneticCharacters = True Then
                    ListOfInvalidPhoneticCharacterWords.Add(String.Join(" ", NewWord.BuildExtendedIpaArray) & vbTab & NewWord.OrthographicForm)
                End If

                'Determining input syllable structure and openness
                Dim TotalErrors As Integer = 0
                TotalErrors = NewWord.DetermineSyllableIndices() 'Determining internal syllable structure
                NewWord.DetermineSyllableOpenness()  'Detecting syllable openness

                'Checking for transcription errors
                If CheckTranscriptionStructure = True Then
                    TotalErrors += NewWord.MarkSyllableWeightErrors()
                    TotalErrors += NewWord.MarkPhoneticLengthInWrongPlace()

                    'Inly adding errors if CheckTranscriptionStructure = true
                    If TotalErrors > 0 Then OtherPhoneticErrorWords.Add(String.Join(" ", NewWord.BuildExtendedIpaArray) & vbTab & NewWord.OrthographicForm & vbTab & "ErrorMessages: " & vbTab & String.Join(", ", NewWord.ManualEvaluations))

                End If

            End If

            'Adds the new word combination
            NewWordGroup.MemberWords.Add(NewWord)

        Next

        'Looking up all available data in the AFC list, filling up data that hasn't been added by the user
        If DoAfcListLookup Then
            NewWordGroup = AfcListLookup(NewWordGroup, FatalErrorDescription)
            If NewWordGroup Is Nothing Then Return Nothing
        End If

        'Calculating some data for words that were not in the Afclist
        CalculateBasicDataForWordNotInAfcList(NewWordGroup, FatalErrorDescription)

        If CalculateWordMetrics(NewWordGroup, CalculateOrthographicTransparency, UseTempSpellingChanges,
                                TemporarySpellingChangeListArray, CalculatePhonotacticProbability,
                                CalculatePhoneticNeighborhoodDensity, CalculateOrthographicNeighborhoodDensity,
                                CalculatePLDx, PLDxCount, CalculateOLDx, OLDxCount, CalculateOrthographicIsolationPoints,
                                CalculatePhoneticIsolationPoints, AddInputWordsToComparisonLists,
                                p2g_Settings, Unresolved_p2g_Character, FatalErrorDescription,
                                DontExportAnything, IsRunFromServer) = False Then Return Nothing

        'Returning the word group where all results are stored
        Return NewWordGroup

    End Function


    ''' <summary>
    ''' Calculates word metrics on the input word group words. Returns True if successful or False if an error occurred.
    ''' </summary>
    ''' <param name="InputWordGroup"></param>
    ''' <param name="CalculateOrthographicTransparency"></param>
    ''' <param name="UseTempSpellingChanges"></param>
    ''' <param name="TemporarySpellingChangeListArray"></param>
    ''' <param name="CalculatePhonotacticProbability"></param>
    ''' <param name="CalculatePhoneticNeighborhoodDensity"></param>
    ''' <param name="CalculatePLDx"></param>
    ''' <param name="PLDxCount"></param>
    ''' <param name="CalculateOLDx"></param>
    ''' <param name="OLDxCount"></param>
    ''' <param name="p2g_Settings"></param>
    ''' <param name="Unresolved_p2g_Character"></param>
    ''' <param name="FatalErrorDescription"></param>
    ''' <param name="DontExportAnything"></param>
    ''' <param name="IsRunOnServer"></param>
    ''' <returns></returns>
    Public Function CalculateWordMetrics(ByRef InputWordGroup As WordGroup,
                                         Optional ByVal CalculateOrthographicTransparency As Boolean = True,
                                         Optional ByVal UseTempSpellingChanges As Boolean = False,
                                         Optional ByRef TemporarySpellingChangeListArray() As String = Nothing,
                                         Optional ByVal CalculatePhonotacticProbability As Boolean = True,
                                         Optional ByVal CalculatePhoneticNeighborhoodDensity As Boolean = True,
                                         Optional ByVal CalculateOrthographicNeighborhoodDensity As Boolean = True,
                                         Optional ByVal CalculatePLDx As Boolean = True,
                                         Optional ByVal PLDxCount As Integer = 20,
                                         Optional ByVal CalculateOLDx As Boolean = True,
                                         Optional ByVal OLDxCount As Integer = 20,
                                         Optional ByVal CalculateOrthographicIsolationPoints As Boolean = True,
                                         Optional ByVal CalculatePhoneticIsolationPoints As Boolean = True,
                                         Optional ByVal AddInputWordsToComparisonLists As Boolean = True,
                                         Optional ByRef p2g_Settings As p2gParameters = Nothing,
                                         Optional ByRef Unresolved_p2g_Character As String = "!",
                                         Optional ByRef FatalErrorDescription As String = "",
                                         Optional ByRef DontExportAnything As Boolean = True,
                                         Optional ByVal IsRunOnServer As Boolean = False) As Boolean

        'Forcing calculation of ND if LDx is set to true
        If CalculatePLDx = True Then CalculatePhoneticNeighborhoodDensity = True
        If CalculateOLDx = True Then CalculateOrthographicNeighborhoodDensity = True

        'Word metrics calculations

        'Orthographic Transparency
        If CalculateOrthographicTransparency = True Then
            Try

                'Loading OT functions
                If p2g_Settings Is Nothing Then
                    If CustumSpeechDataLocation = "" Then
                        p2g_Settings = New p2gParameters(, , UseTempSpellingChanges, DontExportAnything, IsRunOnServer)
                    Else
                        p2g_Settings = New p2gParameters(, System.IO.Path.Combine(CustumSpeechDataLocation, "p2g_Rules.txt"), UseTempSpellingChanges, DontExportAnything, IsRunOnServer)
                    End If
                    p2g_Settings.ExportAttemptedSpellingSegmentations = False
                    p2g_Settings.ExportOnlyFailedAttemptedSpellingSegmentations = False
                    p2g_Settings.MaximumNumberOfExampleWords = 0
                    p2g_Settings.UseExampleWordSampling = False
                    p2g_Settings.WarnForIllegalCharactersInSpelling = False
                    p2g_Settings.LengthSensitiveLookUpPhonemes = True
                End If
                If p2g_Settings.UseTemporarySpellingChange = True Then
                    p2g_Settings.TemporarySpellingChangeListArray = TemporarySpellingChangeListArray
                End If

                'Running identification of sonographs
                InputWordGroup.GetSonographs(p2g_Settings)

                'Orthographic transparency
                Dim g2p_Probability As GraphemeToPhonemes = Nothing
                Dim GIL2P_Probability As KeyInitialSegmentToValueProbability = Nothing
                Dim PIP2G_Probability As KeyInitialSegmentToValueProbability = Nothing

                If CustumSpeechDataLocation = "" Then
                    g2p_Probability = GraphemeToPhonemes.Load_g2p_DataFromFile(,, Unresolved_p2g_Character, WordGroup.WordFrequencyUnit.WordType)

                    GIL2P_Probability = KeyInitialSegmentToValueProbability.Load_KIS2V_FromTxtFile(KeyInitialSegmentToValueProbability.OrthographicTransparencyTypes.GIL2P)

                    PIP2G_Probability = KeyInitialSegmentToValueProbability.Load_KIS2V_FromTxtFile(KeyInitialSegmentToValueProbability.OrthographicTransparencyTypes.PIP2G)
                Else
                    g2p_Probability = GraphemeToPhonemes.Load_g2p_DataFromFile(,, Unresolved_p2g_Character, WordGroup.WordFrequencyUnit.WordType,
                                                                                             System.IO.Path.Combine(CustumSpeechDataLocation, "g2p_Data.txt"))

                    GIL2P_Probability = KeyInitialSegmentToValueProbability.Load_KIS2V_FromTxtFile(KeyInitialSegmentToValueProbability.OrthographicTransparencyTypes.GIL2P,
                                                                                                                 System.IO.Path.Combine(CustumSpeechDataLocation, "GIL2P_Data.txt"))

                    PIP2G_Probability = KeyInitialSegmentToValueProbability.Load_KIS2V_FromTxtFile(KeyInitialSegmentToValueProbability.OrthographicTransparencyTypes.PIP2G,
                                                                                                                 System.IO.Path.Combine(CustumSpeechDataLocation, "PIP2G_Data.txt"))
                End If



                'Generating loading problem messages
                If g2p_Probability Is Nothing Then
                    FatalErrorDescription = "Detected problems reading the g2p_Data data file. Cannot procede!"
                    Return False
                End If
                If GIL2P_Probability Is Nothing Then
                    FatalErrorDescription = "Detected problems reading the GIL2P_Data data file. Cannot procede!"
                    Return False
                End If
                If PIP2G_Probability Is Nothing Then
                    FatalErrorDescription = "Detected problems reading the PIP2G_Data data file. Cannot procede!"
                    Return False
                End If

                'Calculating G2P_Probability
                g2p_Probability.Apply_G2P_Data(InputWordGroup)

                'Calculating GIL2P_Probability
                GIL2P_Probability.Apply_KIS2V_OT_Data(InputWordGroup)

                'Calculating PIP2G_Probability
                PIP2G_Probability.Apply_KIS2V_OT_Data(InputWordGroup)

            Catch ex As Exception
                FatalErrorDescription = ex.ToString
                Return False
            End Try
        End If


        'Phonotactic probability
        If CalculatePhonotacticProbability = True Then
            Try

                Dim SSPP_Data As PhonoTactics
                Dim PSP_Data As Positional_PhonoTactics
                Dim PSBP_Data As Positional_PhonoTactics
                Dim SyllabificationTool As Syllabification
                Dim RoundingDecimals As Integer = 6

                'Initiating the resyllabification tool, and loading cluster data
                SyllabificationTool = New Syllabification(,, True)
                If CustumSpeechDataLocation = "" Then
                    If SyllabificationTool.LoadClusterDataFromFile() = False Then
                        FatalErrorDescription = "Detected problems reading the SyllabificationClusters data file. Cannot procede!"
                        Return False
                    End If
                Else
                    If SyllabificationTool.LoadClusterDataFromFile(System.IO.Path.Combine(CustumSpeechDataLocation, "SyllabificationClusters.txt")) = False Then
                        FatalErrorDescription = "Detected problems reading the SyllabificationClusters data file. Cannot procede!"
                        Return False
                    End If
                End If


                'Initiating Witte type phonotactic probability calculator
                SSPP_Data = New PhonoTactics(PhonoTactics.PhonoTacticCalculationTypes.PhonotacticPredictability,,,
                                                                      True, False, False, PhonoTactics.FrequencyUnits.PhonemeCount)

                If CustumSpeechDataLocation = "" Then
                    SSPP_Data.LoadProbabilityDataFromFile()
                Else
                    SSPP_Data.LoadProbabilityDataFromFile(System.IO.Path.Combine(CustumSpeechDataLocation, "SSPP_Matrix_FullLines.txt"))
                End If




                'Initiating the Positional_PhonoTactics calculator - monogram probabilities
                PSP_Data = New Positional_PhonoTactics(,, False, False, False,
                                                                                   Positional_PhonoTactics.PhonemeCombinationLengths.MonoGramCalculation,
                                                                                     Positional_PhonoTactics.FrequencyUnits.WordCount,)
                If CustumSpeechDataLocation = "" Then
                    PSP_Data.LoadProbabilityDataFromFile()
                Else
                    PSP_Data.LoadProbabilityDataFromFile(System.IO.Path.Combine(CustumSpeechDataLocation, "PSP_Matrix_FullLines.txt"))
                End If


                'Initiating the Positional_PhonoTactics calculator - bigram probabilities
                PSBP_Data = New Positional_PhonoTactics(,, False, False, False,
                                                                                        Positional_PhonoTactics.PhonemeCombinationLengths.BiGramCalculation,
                                                                                         Positional_PhonoTactics.FrequencyUnits.WordCount,)

                If CustumSpeechDataLocation = "" Then
                    PSBP_Data.LoadProbabilityDataFromFile()
                Else
                    PSBP_Data.LoadProbabilityDataFromFile(System.IO.Path.Combine(CustumSpeechDataLocation, "PSBP_Matrix_FullLines.txt"))
                End If

                'Generating loading problem messages
                If SSPP_Data Is Nothing Then
                    FatalErrorDescription = "Detected problems reading the SSPP_Matrix_FullLines data file. Cannot procede!"
                    Return False
                End If
                If PSP_Data Is Nothing Then
                    FatalErrorDescription = "Detected problems reading the PSP_Matrix_FullLines data file. Cannot procede!"
                    Return False
                End If
                If PSBP_Data Is Nothing Then
                    FatalErrorDescription = "Detected problems reading the PSBP_Matrix_FullLines data file. Cannot procede!"
                    Return False
                End If

                'TODO: Are RemoveDoublePhoneticCharacters and this ChangeAmbiSyllabicLongConsonantPositions needed, or is it also done by the syllabificationTool? They are left, in case not.
                'Removing double consonants which may have come about in the last step
                InputWordGroup.RemoveDoublePhoneticCharacters(True, False, True, False, True, True)
                'Standardizing the locations of long ambisyllablic consonants to the coda before the syllable boundary
                InputWordGroup.ChangeAmbiSyllabicLongConsonantPositions(WordGroup.LongConsonantPositions.SyllableCoda, False)

                'Doing resyllabification into the alternate syllable structures (only used for the SSPP type)
                If SyllabificationTool.Syllabify(InputWordGroup, True) = True Then

                    'Calculating SSPP phonotactic probability
                    SSPP_Data.TransitionData(PhonotacticTransitionDataDirections.RetrievePhonotacticProbabilities, InputWordGroup,,, False)

                    'Calculating positional phonotactics - monogram and biphone probabilities

                    'Normal type
                    PSP_Data.TransitionData(PhonotacticTransitionDataDirections.RetrievePhonotacticProbabilities, InputWordGroup, , ,,,, False)
                    PSBP_Data.TransitionData(PhonotacticTransitionDataDirections.RetrievePhonotacticProbabilities, InputWordGroup, , ,,,, False)

                    'Z-transformed type
                    PSP_Data.TransitionData(PhonotacticTransitionDataDirections.RetrievePhonotacticProbabilities, InputWordGroup, , ,,,, True)
                    PSBP_Data.TransitionData(PhonotacticTransitionDataDirections.RetrievePhonotacticProbabilities, InputWordGroup, , ,,,, True)

                Else

                    FatalErrorDescription &= "Syllabification failed for one or more words. Therefore unable to calculate phonotactic probability. Look in the column ManualEvaluations to find out for which words this problem occurred."

                End If

            Catch ex As Exception
                FatalErrorDescription = ex.ToString
                Return False
            End Try
        End If


        'Phonetic neighborhood density
        Dim PLDComparisonCorpus As WordGroup.PLDComparisonCorpus = Nothing
        If CalculatePhoneticNeighborhoodDensity = True Or CalculatePhoneticIsolationPoints = True Then
            Try
                'Loading the comparisoncorpus if it has not been loaded

                If CustumSpeechDataLocation = "" Then
                    PLDComparisonCorpus = WordGroup.PLDComparisonCorpus.LoadComparisonCorpus()
                Else
                    PLDComparisonCorpus = WordGroup.PLDComparisonCorpus.LoadComparisonCorpus(System.IO.Path.Combine(CustumSpeechDataLocation, "PLDComparisonCorpus_AfcList.txt"))
                End If

                If PLDComparisonCorpus Is Nothing Then
                    FatalErrorDescription = "Detected problems reading the AfcList PLDComparisonCorpus. Cannot procede!"
                    Return False
                End If

                'Adding words lacking in the AFC-list to the present comparison corpus.
                If AddInputWordsToComparisonLists = True Then

                    'Dim NewWordsAddedByUser As New List(Of WordLists.Word)
                    For Each CurrentWord In InputWordGroup.MemberWords
                        'Creating a PDL1 transcription for the current word
                        Dim CurrentWordPLD1_Trandscription As List(Of String) = CurrentWord.BuildPLD1TypeTranscription
                        Dim CurrentWordPLD1_TrandscriptionString As String = String.Join(" ", CurrentWordPLD1_Trandscription)

                        'Adding the PDL1 transcription and Zipf-scale value of the word to the ComparisonCorpus, only if the PLD transcription does not already exist in the ComparisonCorpus.
                        'Adding the syllable length key, if it doesn't already exist
                        If Not PLDComparisonCorpus.ContainsKey(CurrentWord.Syllables.Count) Then
                            PLDComparisonCorpus.Add(CurrentWord.Syllables.Count, New WordGroup.PLD_SyllableLengthSpecificComparisonCorpusData)
                        End If

                        'Checking if the PLD1 transcription is not already included in the ComparisonCorpus (otherwise ignoring it)
                        If Not PLDComparisonCorpus(CurrentWord.Syllables.Count).ContainsKey(CurrentWordPLD1_TrandscriptionString) Then

                            'Adds the PLD1 transcription and it's user specified Zipf-scale value to ComparisonCorpus
                            PLDComparisonCorpus(CurrentWord.Syllables.Count).Add(CurrentWordPLD1_TrandscriptionString, New WordGroup.PLD_ComparisonCorpusData With {.PLD1Transcription = CurrentWordPLD1_Trandscription, .ZipfValue = CurrentWord.ZipfValue_Word})

                        End If
                    Next
                End If

            Catch ex As Exception
                FatalErrorDescription &= ex.ToString
                Return False
            End Try
        End If

        If CalculatePhoneticNeighborhoodDensity = True Then
            Try

                'Calculating PLD1 and optionally PLDx, based on the loaded comparison corpus
                Dim tempFatalError As String = ""
                InputWordGroup.Calculate_PLD_UsingComparisonCorpus(PLDComparisonCorpus, ,, tempFatalError, CalculatePLDx, PLDxCount)

                FatalErrorDescription &= tempFatalError
                If tempFatalError <> "" Then
                    Return False
                End If

                'Calculating FWPN_DensityProbability
                InputWordGroup.Calculate_FWPN_DensityProbability(,,, True)

                'Updating PLD1NeighbourCount
                InputWordGroup.Update_PLD1NeighbourCount()

            Catch ex As Exception
                FatalErrorDescription &= ex.ToString
                Return False
            End Try
        End If

        'Calculating calculatePhoneticIsolationPoint
        If CalculatePhoneticIsolationPoints = True Then
            Try
                InputWordGroup.CalculatePhoneticIsolationPoints(PLDComparisonCorpus, FatalErrorDescription, Not IsRunOnServer)
            Catch ex As Exception
                FatalErrorDescription &= ex.ToString
                Return False
            End Try
        End If


        'Loading OLD comparison corpus from file
        Dim OLDComparisonCorpus As WordGroup.OrthographicComparisonCorpus = Nothing
        If CalculateOrthographicNeighborhoodDensity = True Or CalculateOrthographicIsolationPoints = True Then

            If CustumSpeechDataLocation = "" Then
                OLDComparisonCorpus = WordGroup.OrthographicComparisonCorpus.LoadComparisonCorpus(, FatalErrorDescription)
            Else
                OLDComparisonCorpus = WordGroup.OrthographicComparisonCorpus.LoadComparisonCorpus(System.IO.Path.Combine(CustumSpeechDataLocation, "OLDComparisonCorpus_AfcList.txt"), FatalErrorDescription)
            End If

            If OLDComparisonCorpus Is Nothing Then
                FatalErrorDescription &= "Detected problems reading the AfcList OLDComparisonCorpus. Cannot procede!"
                Return False
            End If

            If AddInputWordsToComparisonLists = True Then

                'Adding words lacking in the AFC-list to the present comparison corpus.
                If AddInputWordsToComparisonLists = True Then

                    'Dim NewWordsAddedByUser As New List(Of WordLists.Word)
                    For Each CurrentWord In InputWordGroup.MemberWords

                        'Adding the spelling and the Zipf-scale value of the word to the ComparisonCorpus, only if the spelling does not already exist in the ComparisonCorpus.
                        If Not OLDComparisonCorpus.ContainsKey(CurrentWord.OrthographicForm) Then

                            'Adds the spelling and it's user specified Zipf-scale value to ComparisonCorpus
                            OLDComparisonCorpus.Add(CurrentWord.OrthographicForm, CurrentWord.ZipfValue_Word)

                        End If
                    Next
                End If
            End If
        End If

        If CalculateOrthographicNeighborhoodDensity = True Then
            Try

                'Calculating OLD1 and optionally OLDx
                InputWordGroup.Calculate_OLD_UsingComparisonCorpus(OLDComparisonCorpus, FatalErrorDescription, CalculateOLDx, OLDxCount, IsRunOnServer)

                'Calculating FWON_DensityProbability
                InputWordGroup.Calculate_FWON_DensityProbability(,,, True)

                InputWordGroup.Update_OLD1NeighbourCount()

            Catch ex As Exception
                FatalErrorDescription &= ex.ToString
                Return False
            End Try
        End If

        'TODO: Here we should add calculation of OLD1_Count

        'Calculating calculatePhoneticIsolationPoint
        If CalculateOrthographicIsolationPoints = True Then
            Try
                InputWordGroup.CalculateOrthographicIsolationPoints(OLDComparisonCorpus, FatalErrorDescription, IsRunOnServer)
            Catch ex As Exception
                FatalErrorDescription &= ex.ToString
                Return False
            End Try
        End If


        Return True

    End Function

    ''' <summary>
    ''' Fills up the output word group with all available data from the AFC list. 
    ''' For input words that contain only spelling, all AFC list words with that spelling are added.
    ''' For input words with both a spelling and a transcription specified, only data from AFC list words with exact matches in spelling and transcriptions are added.
    ''' User supplied ZipfValues are only used in cases when the input word does not exist in the AFC list.
    ''' </summary>
    ''' <param name="InputWordGroup"></param>
    ''' <param name="FatalErrorDescription"></param>
    ''' <returns></returns>
    Public Function AfcListLookup(ByRef InputWordGroup As WordGroup,
                                  Optional ByRef FatalErrorDescription As String = "")

        Try

            Dim AfcListConnection = New MySqlConnection(AfcListMySqlConnectionString)
            If AfcListConnection Is Nothing Then
                FatalErrorDescription = "Unable to establish a connection with the AFC-list database."
                Return Nothing
            End If

            'Tries to open the connection
            Try
                AfcListConnection.Open()
            Catch ex As Exception
                FatalErrorDescription = ex.ToString
                Return Nothing
            End Try

            'Remove duplicates from the input list?

            'Creates an output word group
            Dim OutputWordGroup As New WordGroup
            OutputWordGroup.GetCorpusInfoFromOtherWordgroup(InputWordGroup)

            For WordIndex = 0 To InputWordGroup.MemberWords.Count - 1

                Dim CurrentInputWord = InputWordGroup.MemberWords(WordIndex)

                Dim CurrentOrthForm As String = CurrentInputWord.OrthographicForm
                Dim CurrentPhoneticForm As String = String.Join(" ", CurrentInputWord.BuildExtendedIpaArray).Trim
                Dim CurrentZipfValue As String = CurrentInputWord.ZipfValue_Word

                'Skips if CurrentOrthForm is empty
                If CurrentOrthForm = "" Then Continue For

                'Checks if we have a phonetic form
                Dim Query As String = ""
                If CurrentPhoneticForm = "" Then
                    'If not, look up all phonetic forms with the current orthographic form
                    Query = "SELECT * FROM " & AfcListTableName & vbCr & vbLf &
                        "WHERE OrthographicForm='" & CurrentOrthForm & "';"
                Else
                    'If the user has input a phonetic form, check if the AfcList contains the spelling/pronunciation combination, and add available data
                    Query = "SELECT * FROM " & AfcListTableName & vbCr & vbLf &
                        "WHERE OrthographicForm='" & CurrentOrthForm &
                        "' AND PhoneticForm='" & CurrentPhoneticForm & "';"
                End If

                Dim AfcListDataAdapter As MySqlDataAdapter = New MySqlDataAdapter(Query, AfcListConnection)
                Dim CurrentWordsTable As New DataTable
                AfcListDataAdapter.Fill(CurrentWordsTable)


                If CurrentWordsTable.Rows.Count = 0 Then
                    'The input word was lacking in the AFC list. Adds the input word without any added AFC list data.
                    OutputWordGroup.MemberWords.Add(CurrentInputWord)

                    'Marks the word as lacking word list data (some columns will be blank in the website output)
                    CurrentInputWord.ContainsWordListData = False

                Else

                    'Adds the words found in the AfcList look-up
                    For n = 0 To CurrentWordsTable.Rows.Count - 1

                        'Creates a new word
                        Dim NewWord As New Word
                        NewWord.OrthographicForm = CurrentOrthForm 'CurrentWordsTable(n)("OrthographicForm")
                        NewWord.ProportionStartingWithUpperCase = CurrentWordsTable(n)("UpperCase")

                        'Only reading Homographs if the Homograph input string is not empty
                        If IsDBNull(CurrentWordsTable(n)("Homographs")) = False Then
                            Dim CurrentHomograps As String = CurrentWordsTable(n)("Homographs").trim
                            If Not CurrentHomograps = "" Then
                                Dim InputForms() As String = CurrentHomograps.Split("|")
                                For CurrentIndex = 0 To InputForms.Length - 1
                                    Dim newInputForm As String = InputForms(CurrentIndex).Trim
                                    NewWord.LanguageHomographs = New List(Of String)
                                    If newInputForm <> "" Then
                                        NewWord.LanguageHomographs.Add(newInputForm)
                                    End If
                                Next
                            End If
                        End If
                        NewWord.OrthographicFormContainsSpecialCharacter = CurrentWordsTable(n)("SpecialCharacter")
                        NewWord.RawWordTypeFrequency = CurrentWordsTable(n)("RawWordTypeFrequency")
                        NewWord.RawDocumentCount = CurrentWordsTable(n)("RawDocumentCount")


                        'N.B. Only reading phonetic form if the phonetic input string is not empty (which it should never be). 
                        Dim PhoneticInputString As String = CurrentWordsTable(n)("PhoneticForm").trim
                        If Not PhoneticInputString = "" Then
                            Dim ContainsInvalidPhoneticCharacter As Boolean
                            Dim CorrectedDoubleSpacesInPhoneticForm As Boolean
                            'WordLists.WordListsIO.ParseInputPhoneticString(PhoneticInputString, NewWord, ValidPhoneticCharacters,
                            '                         ContainsInvalidPhoneticCharacter, True,
                            '                         CorrectedDoubleSpacesInPhoneticForm, True)

                            NewWord.ParseInputPhoneticString(PhoneticInputString, ValidPhoneticCharacters,
                                                     ContainsInvalidPhoneticCharacter, True,
                                                     CorrectedDoubleSpacesInPhoneticForm, True)


                            Dim TotalErrors As Integer = NewWord.DetermineSyllableIndices() 'Determining internal syllable structure
                            NewWord.DetermineSyllableOpenness()  'Detecting syllable openness

                            If ContainsInvalidPhoneticCharacter = True Or CorrectedDoubleSpacesInPhoneticForm = True Then 'Or TotalErrors > 0 Then
                                FatalErrorDescription = "Incorrect transcription detected in the AFC-list!"
                                Return Nothing
                            End If

                        End If

                        NewWord.PhonotacticType = CurrentWordsTable(n)("PhonotacticType")

                        'Only reading Homophones if the Homophones input string is not empty
                        If IsDBNull(CurrentWordsTable(n)("Homophones")) = False Then
                            Dim CurrentHomophones As String = CurrentWordsTable(n)("Homophones").trim
                            If Not CurrentHomophones = "" Then
                                Dim InputForms() As String = CurrentHomophones.Split("|")
                                For CurrentIndex = 0 To InputForms.Length - 1
                                    Dim newInputForm As String = InputForms(CurrentIndex).Trim
                                    NewWord.LanguageHomophones = New List(Of String)
                                    If Not newInputForm = "" Then
                                        NewWord.LanguageHomophones.Add(newInputForm)
                                    End If
                                Next
                            End If
                        End If
                        If IsDBNull(CurrentWordsTable(n)("AllPoS")) = False Then

                            Dim AllPossiblePoS_InputString As String = CurrentWordsTable(n)("AllPoS").trim

                            'Only reads PoS if the string is not empty
                            If Not AllPossiblePoS_InputString = "" Then

                                Dim AllPoS() As String = AllPossiblePoS_InputString.Split("|")
                                For PoS = 0 To AllPoS.Length - 1

                                    Dim CurrentPoSSplit() As String = AllPoS(PoS).Trim.Split(":")
                                    Dim tempPos As String = CurrentPoSSplit(0).Trim

                                    Dim AlreadyAddedPoSs As New SortedSet(Of String)
                                    If Not AlreadyAddedPoSs.Contains(tempPos) Then
                                        AlreadyAddedPoSs.Add(tempPos)

                                        If CurrentPoSSplit.Length = 1 Then
                                            NewWord.AllPossiblePoS.Add(New Tuple(Of String, Double)(tempPos, 0))
                                        ElseIf CurrentPoSSplit.Length > 1 Then
                                            NewWord.AllPossiblePoS.Add(New Tuple(Of String, Double)(tempPos, CurrentPoSSplit(1).Trim))
                                        End If
                                    Else
                                        'Just ignores any erroneous duplicates here
                                    End If
                                Next
                            End If
                        End If
                        If IsDBNull(CurrentWordsTable(n)("AllLemmas")) = False Then

                            Dim AllOccurringLemmas_InputString As String = CurrentWordsTable(n)("AllLemmas").trim

                            'Only reads AllLemmas if the string is not empty
                            If Not AllOccurringLemmas_InputString = "" Then

                                Dim AllLemmas() As String = AllOccurringLemmas_InputString.Split("|")
                                For Lemma = 0 To AllLemmas.Length - 1

                                    Dim CurrentLemmaSplit() As String = AllLemmas(Lemma).Trim.Split(":")
                                    Dim temp_Lemma As String = CurrentLemmaSplit(0).Trim

                                    Dim AlreadyAddedLemmas As New SortedSet(Of String)
                                    If Not AlreadyAddedLemmas.Contains(temp_Lemma) Then
                                        AlreadyAddedLemmas.Add(temp_Lemma)

                                        If CurrentLemmaSplit.Length = 1 Then
                                            NewWord.AllOccurringLemmas.Add(New Tuple(Of String, Double)(temp_Lemma, 0))
                                        ElseIf CurrentLemmaSplit.Length > 1 Then
                                            NewWord.AllOccurringLemmas.Add(New Tuple(Of String, Double)(temp_Lemma, CurrentLemmaSplit(1).Trim))
                                        End If

                                    Else
                                        'Just ignores any erroneous duplicates here
                                    End If
                                Next
                            End If
                        End If

                        If IsDBNull(CurrentWordsTable(n)("NumberOfSenses")) = False Then
                            NewWord.NumberOfSenses = CurrentWordsTable(n)("NumberOfSenses")
                        End If

                        NewWord.Abbreviation = CurrentWordsTable(n)("Abbreviation")
                        NewWord.Acronym = CurrentWordsTable(n)("Acronym")
                        NewWord.ForeignWord = CurrentWordsTable(n)("ForeignWord")
                        NewWord.ZipfValue_Word = CurrentWordsTable(n)("ZipfValue")

                        'Acctually some AFC-list columns are not read here. However, that data is (or could be) generated from the columns read.

                        'Adding the word
                        OutputWordGroup.MemberWords.Add(NewWord)

                    Next
                End If

            Next

            AfcListConnection.Close()

            Return OutputWordGroup

        Catch ex As Exception
            FatalErrorDescription = ex.ToString
            Return Nothing
        End Try

    End Function


    Private Sub CalculateBasicDataForWordNotInAfcList(ByRef InputWordGroup As WordGroup,
                                  Optional ByRef FatalErrorDescription As String = "")


        For word = 0 To InputWordGroup.MemberWords.Count - 1

            'Skips if the word has AFC-list data
            If InputWordGroup.MemberWords(word).ContainsWordListData = True Then Continue For

            'Skips to next if there is no phonetic form
            Dim CurrentPhoneticForm As String = String.Join(" ", InputWordGroup.MemberWords(word).BuildExtendedIpaArray).Trim
            If CurrentPhoneticForm = "" Then Continue For

            'Calculates phonotactic type
            InputWordGroup.MemberWords(word).SetWordPhonotacticType()
        Next


        'The following section to look up homophones and homographs is skipped.
        'TODO should we activate it again?
        Exit Sub

        Dim AfcListConnection = New MySqlConnection(AfcListMySqlConnectionString)

        If AfcListConnection Is Nothing Then
            FatalErrorDescription = "Unable to establish a connection with the AFC-list database."
            Exit Sub
        End If

        'Tries to open the connection
        Try
            AfcListConnection.Open()
        Catch ex As Exception
            FatalErrorDescription = ex.ToString
            Exit Sub
        End Try

        For word = 0 To InputWordGroup.MemberWords.Count - 1

            'Skips if the word has AFC-list data
            If InputWordGroup.MemberWords(word).ContainsWordListData = True Then Continue For

            'Skips to next if there is no phonetic form
            Dim CurrentPhoneticForm As String = String.Join(" ", InputWordGroup.MemberWords(word).BuildExtendedIpaArray).Trim
            If CurrentPhoneticForm = "" Then Continue For

            'Calculates phonotactic type
            InputWordGroup.MemberWords(word).SetWordPhonotacticType()


            'N.B. Only homographs and homophones already present in the AFC-list will be detected for any new word. Homographs and homophones that exist among other new words will not be detected.
            'New words will however not be added as homographs to the AFC-list words

            'Gets homographs

            'Getting a list of AFC-list words with the same spelling
            Dim Query As String = "SELECT * FROM " & AfcListTableName & vbCr & vbLf &
                    "WHERE OrthographicForm='" & InputWordGroup.MemberWords(word).OrthographicForm & "';"

            Dim AfcListDataAdapter As MySqlDataAdapter = New MySqlDataAdapter(Query, AfcListConnection)

            Dim CurrentWordsTable As New DataTable
            AfcListDataAdapter.Fill(CurrentWordsTable)

            'Creates a reduced transcription for the current input word
            Dim ReducedTranscription As String = String.Join(" ", InputWordGroup.MemberWords(word).BuildReducedIpaArray(False))

            If CurrentWordsTable.Rows.Count > 0 Then

                For n = 0 To CurrentWordsTable.Rows.Count - 1
                    If CurrentWordsTable(n)("ReducedTranscription") <> ReducedTranscription Then

                        'Adding the homograph
                        InputWordGroup.MemberWords(word).LanguageHomographs.Add(CurrentWordsTable(n)("ReducedTranscription"))

                    End If
                Next
            End If


            'Determines homophones

            'Getting a list of AFC-list words with the same ReducedTranscription
            Query = "SELECT * FROM " & AfcListTableName & vbCr & vbLf &
                    "WHERE ReducedTranscription='" & ReducedTranscription & "';"

            AfcListDataAdapter = New MySqlDataAdapter(Query, DirectCast(AfcListConnection, MySqlConnection))

            CurrentWordsTable = New DataTable

            AfcListDataAdapter.Fill(CurrentWordsTable)

            If CurrentWordsTable.Rows.Count > 0 Then
                For n = 0 To CurrentWordsTable.Rows.Count - 1
                    If CurrentWordsTable(n)("OrthographicForm") <> InputWordGroup.MemberWords(word).OrthographicForm Then

                        'Adding the homophone
                        InputWordGroup.MemberWords(word).LanguageHomophones.Add(CurrentWordsTable(n)("OrthographicForm"))

                    End If
                Next
            End If

        Next


    End Sub

End Class

Public Module AfcListSearch

    Public Function SearchAfcList(ByVal SqlQuery As String,
                                  ByRef TotalNumberOfHits As Integer,
                                  Optional ByRef FatalErrorDescription As String = "",
                                  ByVal Optional IndexOfFirstHit As Integer = 0,
                                  ByVal Optional MaxNumberOfHits As Integer = Integer.MaxValue) As List(Of TextOnlyWord)

        Try

            If AfcListMySqlConnectionString = "" Then AfcListMySqlDatabaseInfo.LoadDefaultAfcListMySqlDatabaseInfo()

            Dim AfcListConnection = New MySqlConnection(AfcListMySqlConnectionString)
            If AfcListConnection Is Nothing Then
                FatalErrorDescription = "Unable to establish a connection with the AFC-list database."
                Return Nothing
            End If

            'Tries to open the connection
            Try
                AfcListConnection.Open()
            Catch ex As Exception
                FatalErrorDescription = ex.ToString
                Return Nothing
            End Try

            Dim AfcListDataAdapter As MySqlDataAdapter = New MySqlDataAdapter(SqlQuery, AfcListConnection)
            Dim CurrentWordsTable As New DataTable
            AfcListDataAdapter.Fill(CurrentWordsTable)

            'String the total number of hits
            TotalNumberOfHits = CurrentWordsTable.Rows.Count

            Dim OutputList = New List(Of TextOnlyWord)

            Dim TempStartIndex As Integer = IndexOfFirstHit
            If TempStartIndex < 0 Then TempStartIndex = 0

            Dim TempLastIndex As Integer = TempStartIndex + MaxNumberOfHits - 1
            TempLastIndex = Math.Min(TempLastIndex, CurrentWordsTable.Rows.Count - TempStartIndex - 1)

            'Adds the words found in the AfcList look-up (limited to the selected index range)
            For n = TempStartIndex To TempLastIndex

                'For c = 0 To CurrentWordsTable.Columns.Count - 1
                '    Console.WriteLine(CurrentWordsTable.Columns(c).ColumnName)
                'Next

                'Creating a new TextOnlyWord
                Dim NewTextOnlyWord As New TextOnlyWord

                Dim CurrentRow = CurrentWordsTable(n)

                Try

                    '(Since the incoming AFC-list column order is fixed, the copying can be hard coded here.) (However, we need to check for dBNull, but perhaps not for all AFC-list columns)
                    Integer.TryParse(CurrentRow.ItemArray(0), NewTextOnlyWord.Id)
                    If Not IsDBNull(CurrentRow.ItemArray(0)) Then NewTextOnlyWord.Id = CurrentRow.ItemArray(0)
                    If Not IsDBNull(CurrentRow.ItemArray(1)) Then NewTextOnlyWord.OrthographicForm = CurrentRow.ItemArray(1)
                    If Not IsDBNull(CurrentRow.ItemArray(2)) Then NewTextOnlyWord.GIL2P_OT_Average = CurrentRow.ItemArray(2)
                    If Not IsDBNull(CurrentRow.ItemArray(3)) Then NewTextOnlyWord.GIL2P_OT_Min = CurrentRow.ItemArray(3)
                    If Not IsDBNull(CurrentRow.ItemArray(4)) Then NewTextOnlyWord.PIP2G_OT_Average = CurrentRow.ItemArray(4)
                    If Not IsDBNull(CurrentRow.ItemArray(5)) Then NewTextOnlyWord.PIP2G_OT_Min = CurrentRow.ItemArray(5)
                    If Not IsDBNull(CurrentRow.ItemArray(6)) Then NewTextOnlyWord.G2P_OT_Average = CurrentRow.ItemArray(6)
                    If Not IsDBNull(CurrentRow.ItemArray(7)) Then NewTextOnlyWord.UpperCase = CurrentRow.ItemArray(7)
                    If Not IsDBNull(CurrentRow.ItemArray(8)) Then NewTextOnlyWord.Homographs = CurrentRow.ItemArray(8)
                    If Not IsDBNull(CurrentRow.ItemArray(9)) Then NewTextOnlyWord.HomographCount = CurrentRow.ItemArray(9)
                    If Not IsDBNull(CurrentRow.ItemArray(10)) Then NewTextOnlyWord.SpecialCharacter = CurrentRow.ItemArray(10)
                    If Not IsDBNull(CurrentRow.ItemArray(11)) Then NewTextOnlyWord.RawWordTypeFrequency = CurrentRow.ItemArray(11)
                    If Not IsDBNull(CurrentRow.ItemArray(12)) Then NewTextOnlyWord.RawDocumentCount = CurrentRow.ItemArray(12)
                    If Not IsDBNull(CurrentRow.ItemArray(13)) Then NewTextOnlyWord.PhoneticForm = CurrentRow.ItemArray(13)
                    If Not IsDBNull(CurrentRow.ItemArray(14)) Then NewTextOnlyWord.ReducedTranscription = CurrentRow.ItemArray(14)
                    If Not IsDBNull(CurrentRow.ItemArray(15)) Then NewTextOnlyWord.PhonotacticType = CurrentRow.ItemArray(15)
                    If Not IsDBNull(CurrentRow.ItemArray(16)) Then NewTextOnlyWord.SSPP_Average = CurrentRow.ItemArray(16)
                    If Not IsDBNull(CurrentRow.ItemArray(17)) Then NewTextOnlyWord.SSPP_Min = CurrentRow.ItemArray(17)
                    If Not IsDBNull(CurrentRow.ItemArray(18)) Then NewTextOnlyWord.PSP_Sum = CurrentRow.ItemArray(18)
                    If Not IsDBNull(CurrentRow.ItemArray(19)) Then NewTextOnlyWord.PSBP_Sum = CurrentRow.ItemArray(19)
                    If Not IsDBNull(CurrentRow.ItemArray(20)) Then NewTextOnlyWord.S_PSP_Average = CurrentRow.ItemArray(20)
                    If Not IsDBNull(CurrentRow.ItemArray(21)) Then NewTextOnlyWord.S_PSBP_Average = CurrentRow.ItemArray(21)
                    If Not IsDBNull(CurrentRow.ItemArray(22)) Then NewTextOnlyWord.Homophones = CurrentRow.ItemArray(22)
                    If Not IsDBNull(CurrentRow.ItemArray(23)) Then NewTextOnlyWord.HomophoneCount = CurrentRow.ItemArray(23)
                    If Not IsDBNull(CurrentRow.ItemArray(24)) Then NewTextOnlyWord.PNDP = CurrentRow.ItemArray(24)
                    If Not IsDBNull(CurrentRow.ItemArray(25)) Then NewTextOnlyWord.ONDP = CurrentRow.ItemArray(25)
                    If Not IsDBNull(CurrentRow.ItemArray(26)) Then NewTextOnlyWord.Sonographs = CurrentRow.ItemArray(26)
                    If Not IsDBNull(CurrentRow.ItemArray(27)) Then NewTextOnlyWord.AllPoS = CurrentRow.ItemArray(27)
                    If Not IsDBNull(CurrentRow.ItemArray(28)) Then NewTextOnlyWord.AllLemmas = CurrentRow.ItemArray(28)
                    If Not IsDBNull(CurrentRow.ItemArray(29)) Then NewTextOnlyWord.NumberOfSenses = CurrentRow.ItemArray(29)
                    If Not IsDBNull(CurrentRow.ItemArray(30)) Then NewTextOnlyWord.Abbreviation = CurrentRow.ItemArray(30)
                    If Not IsDBNull(CurrentRow.ItemArray(31)) Then NewTextOnlyWord.Acronym = CurrentRow.ItemArray(31)
                    If Not IsDBNull(CurrentRow.ItemArray(32)) Then NewTextOnlyWord.ForeignWord = CurrentRow.ItemArray(32)
                    If Not IsDBNull(CurrentRow.ItemArray(33)) Then NewTextOnlyWord.ZipfValue = CurrentRow.ItemArray(33)
                    If Not IsDBNull(CurrentRow.ItemArray(34)) Then NewTextOnlyWord.LetterCount = CurrentRow.ItemArray(34)
                    If Not IsDBNull(CurrentRow.ItemArray(35)) Then NewTextOnlyWord.GraphemeCount = CurrentRow.ItemArray(35)
                    If Not IsDBNull(CurrentRow.ItemArray(36)) Then NewTextOnlyWord.DiGraphCount = CurrentRow.ItemArray(36)
                    If Not IsDBNull(CurrentRow.ItemArray(37)) Then NewTextOnlyWord.TriGraphCount = CurrentRow.ItemArray(37)
                    If Not IsDBNull(CurrentRow.ItemArray(38)) Then NewTextOnlyWord.LongGraphemesCount = CurrentRow.ItemArray(38)
                    If Not IsDBNull(CurrentRow.ItemArray(39)) Then NewTextOnlyWord.SyllableCount = CurrentRow.ItemArray(39)
                    If Not IsDBNull(CurrentRow.ItemArray(40)) Then NewTextOnlyWord.Tone = CurrentRow.ItemArray(40)
                    If Not IsDBNull(CurrentRow.ItemArray(41)) Then NewTextOnlyWord.MainStressSyllable = CurrentRow.ItemArray(41)
                    If Not IsDBNull(CurrentRow.ItemArray(42)) Then NewTextOnlyWord.SecondaryStressSyllable = CurrentRow.ItemArray(42)
                    If Not IsDBNull(CurrentRow.ItemArray(43)) Then NewTextOnlyWord.PhoneCount = CurrentRow.ItemArray(43)
                    If Not IsDBNull(CurrentRow.ItemArray(44)) Then NewTextOnlyWord.PLD1WordCount = CurrentRow.ItemArray(44)
                    If Not IsDBNull(CurrentRow.ItemArray(45)) Then NewTextOnlyWord.OLD1WordCount = CurrentRow.ItemArray(45)
                    If Not IsDBNull(CurrentRow.ItemArray(46)) Then NewTextOnlyWord.PossiblePoSCount = CurrentRow.ItemArray(46)
                    If Not IsDBNull(CurrentRow.ItemArray(47)) Then NewTextOnlyWord.PossibleLemmaCount = CurrentRow.ItemArray(47)

                Catch ex As Exception
                    Console.WriteLine(ex.ToString)
                End Try

                ''Or else reflection could be used to set values as below (which is probably slower, but I havn't tested it...)
                ''Getting all properties of a TextOnlyWord 
                'Dim PropNameList As New List(Of String)
                'Dim AfcListProperyInfo() As System.Reflection.PropertyInfo = TextOnlyWord.AfcListColumnNamesProperties
                'For Each pi In AfcListProperyInfo
                '    Try
                '        Select Case pi.Name
                '            Case "Id"
                '                GetType(TextOnlyWord).GetProperty(pi.Name).SetValue(NewTextOnlyWord, CurrentWordsTable(n)(pi.Name))
                '            Case Else
                '                GetType(TextOnlyWord).GetProperty(pi.Name).SetValue(NewTextOnlyWord, CurrentWordsTable(n)(pi.Name).ToString)
                '        End Select
                '    Catch ex As Exception
                '        Console.WriteLine(pi.Name)
                '    End Try
                'Next

                OutputList.Add(NewTextOnlyWord)

            Next

            AfcListConnection.Close()

            Return OutputList

        Catch ex As Exception
            FatalErrorDescription = ex.ToString
            Return Nothing
        End Try

    End Function



End Module
