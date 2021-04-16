
Imports System.IO

Public Class Syllabification

        Public ReadOnly Property WordStartMarker As String
        Public ReadOnly Property WordEndMarker As String
        Public ReadOnly Property DuplicateAmbiSyllabicLongConsonants As Boolean

        Public WordInitialOnSetClusters As Dictionary(Of String, Double)
        Public WordFinalCodaClusters As Dictionary(Of String, Double)
        Public WordMedialClusters As Dictionary(Of String, CodaOnsetCombination)


        Public Sub New(Optional ByRef SetWordStartMarker As String = "*", Optional ByRef SetWordEndMarker As String = "_",
                           Optional ByRef SetDuplicateAmbiSyllabicLongConsonants As Boolean = True)
            WordStartMarker = SetWordStartMarker
            WordEndMarker = SetWordEndMarker
            DuplicateAmbiSyllabicLongConsonants = SetDuplicateAmbiSyllabicLongConsonants
        End Sub


        ''' <summary>
        ''' Reads the Syllables stored in WorgGroup.memberWords.Word and collects statistical data that can be used for syllabification. 
        ''' The method uses the WordGroup.MemberWords.Word.Phonemes as intermediate containers.
        ''' The syllabification is based on the probabilities of word initial and word final consonant clusters.
        ''' </summary>
        ''' <param name="InputWordGroup"></param>
        Public Sub CollectClusterProbabilities(ByRef InputWordGroup As WordGroup,
                                                   Optional ByRef FrequencyWeightingUnit As WordGroup.WordFrequencyUnit = WordGroup.WordFrequencyUnit.RawFrequency,
                                                   Optional ByRef IgnoreZeroFrequencyWords As Boolean = True,
                                                   Optional ByRef InferLengthLessConsonants As Boolean = True,
                                                   Optional ByRef InferSwedishSegmentalProcesses As Boolean = True)

            SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'Starting a progress window
            Dim myProgress As New ProgressDisplay
            myProgress.Initialize(InputWordGroup.MemberWords.Count - 1, 0, "Preparing for syllabification...", 100)
            myProgress.Show()

            'Creating phonetic forms
            For word = 0 To InputWordGroup.MemberWords.Count - 1

                'Updating progress
                myProgress.UpdateProgress(word)
                InputWordGroup.MemberWords(word).Phonemes = InputWordGroup.MemberWords(word).BuildExtendedIpaArray(, True,,,, False, False)

            Next

            'Closing the progress display
            myProgress.Close()

            'Collecting word initial and final clusters to use for parsing medial clusters later
            CollectWordInitialAndFinalClusters(InputWordGroup, FrequencyWeightingUnit, IgnoreZeroFrequencyWords, InferLengthLessConsonants, InferSwedishSegmentalProcesses)

            'Creating most possible medial cluster combinations used for parsing medial clusters into coda+onset
            CreateMedialClusterCombinations()

            'Exporting clusters to Txt file
            ExportClustersToFile(logFilePath)


        End Sub

        ''' <summary>
        ''' Resyllabifies the words in the input wordlist.
        ''' </summary>
        ''' <param name="InputWordGroup"></param>
        ''' <param name="PutNewSyllabificationInAlternateSyllabification">If set to False, the new syllabification will replace the old WordGroup.MemberWords.Word.Syllables. If set to True the new syllabifications are stored in WordGroups.MemberWords.Syllables_AlternateSyllabification.</param>
        Public Sub Syllabify(ByRef InputWordGroup As WordGroup, Optional ByRef PutNewSyllabificationInAlternateSyllabification As Boolean = False)

            Dim myProgress As New ProgressDisplay
            'Starting a progress window
            myProgress.Initialize(InputWordGroup.MemberWords.Count - 1, 0, "Updating transcriptions ", 100)
            myProgress.Show()

            'Creating phonetic forms
            For word = 0 To InputWordGroup.MemberWords.Count - 1

                'Updating progress
                myProgress.UpdateProgress(word)

                InputWordGroup.MemberWords(word).Phonemes = InputWordGroup.MemberWords(word).BuildExtendedIpaArray(,,,,, False, False)

            Next
            myProgress.Close()


            'Starting a progress window
            myProgress = New ProgressDisplay
            myProgress.Initialize(InputWordGroup.MemberWords.Count - 1, 0, "Performing syllabification ", 100)
            myProgress.Show()

            'Syllabification
            For word = 0 To InputWordGroup.MemberWords.Count - 1

                'Updating progress
                myProgress.UpdateProgress(word)

                'Creating an empty Syllable Skipping words that have no phonetic transcription (added 2017-11-07, in order to handle words added on the website without phonetic transcription)
                If InputWordGroup.MemberWords(word).Phonemes.Count = 0 Then

                    If PutNewSyllabificationInAlternateSyllabification = False Then
                        InputWordGroup.MemberWords(word).Syllables = New Word.ListOfSyllables
                    Else
                        InputWordGroup.MemberWords(word).Syllables_AlternateSyllabification = New Word.ListOfSyllables
                    End If

                    Continue For
                End If


                Dim ContainsSegmentationGuessing As Boolean = False
                Dim ContainsUnparsableCluster As Boolean = False


                'Doing syllabification
                If PutNewSyllabificationInAlternateSyllabification = False Then

                    Dim OriginalSyllables As Word.ListOfSyllables = InputWordGroup.MemberWords(word).Syllables.CreateCopy

                    InputWordGroup.MemberWords(word).Syllables = CreateSyllabification(InputWordGroup.MemberWords(word).Phonemes, ContainsSegmentationGuessing, ContainsUnparsableCluster)

                    'Updating suprasegmentals to the new values
                    InputWordGroup.MemberWords(word).Tone = InputWordGroup.MemberWords(word).Syllables.Tone
                    InputWordGroup.MemberWords(word).MainStressSyllableIndex = InputWordGroup.MemberWords(word).Syllables.MainStressSyllableIndex
                    InputWordGroup.MemberWords(word).SecondaryStressSyllableIndex = InputWordGroup.MemberWords(word).Syllables.SecondaryStressSyllableIndex

                    'Checking if the two types of syllabification has the same length
                    If Not OriginalSyllables.Count = InputWordGroup.MemberWords(word).Syllables.Count Then
                        SendInfoToLog("New syllabification: " & vbTab & String.Join(" ", InputWordGroup.MemberWords(word).BuildExtendedIpaArray), "SyllabificationsWithAlteredLengths")
                    End If

                    'Exporting "guessings"
                    If ContainsSegmentationGuessing = True Then
                        SendInfoToLog(String.Join(" ", InputWordGroup.MemberWords(word).Phonemes.ToArray) & vbTab &
                                      String.Join(" ", InputWordGroup.MemberWords(word).BuildExtendedIpaArray()), "SyllablificationGuessingClusters")
                    End If

                    'Exporting unparsable cluster words
                    If ContainsUnparsableCluster = True Then
                        SendInfoToLog(String.Join(" ", InputWordGroup.MemberWords(word).Phonemes.ToArray) & vbTab &
                                      String.Join(" ", InputWordGroup.MemberWords(word).BuildExtendedIpaArray(,,,, True)), "SyllablificationUnparsedCluster")
                    End If

                Else

                    InputWordGroup.MemberWords(word).Syllables_AlternateSyllabification = CreateSyllabification(InputWordGroup.MemberWords(word).Phonemes, ContainsSegmentationGuessing, ContainsUnparsableCluster)

                    'Checking if the two types of syllabification has the same length
                    If Not InputWordGroup.MemberWords(word).Syllables.Count = InputWordGroup.MemberWords(word).Syllables_AlternateSyllabification.Count Then
                        SendInfoToLog("Original: " & vbTab & String.Join(" ", InputWordGroup.MemberWords(word).BuildExtendedIpaArray) & vbTab &
                                          "Alternate:" & vbTab & String.Join(" ", InputWordGroup.MemberWords(word).BuildExtendedIpaArray(,,,, True)), "SyllabificationsWithAlteredLengths")
                    End If

                    'Exporting "guessings"
                    If ContainsSegmentationGuessing = True Then
                        SendInfoToLog(String.Join(" ", InputWordGroup.MemberWords(word).Phonemes.ToArray) & vbTab &
                                          String.Join(" ", InputWordGroup.MemberWords(word).BuildExtendedIpaArray(,,,, True)), "SyllablificationGuessingClusters")
                    End If

                    'Exporting unparsable cluster words
                    If ContainsUnparsableCluster = True Then
                        SendInfoToLog(String.Join(" ", InputWordGroup.MemberWords(word).Phonemes.ToArray) & vbTab &
                                      String.Join(" ", InputWordGroup.MemberWords(word).BuildExtendedIpaArray(,,,, True)), "SyllablificationUnparsedCluster")
                    End If
                End If
            Next

            'Closing the progress display
            myProgress.Close()

        End Sub

        Public Sub CollectWordInitialAndFinalClusters(ByRef InputWordGroup As WordGroup,
                                                          ByRef FrequencyWeightingUnit As WordGroup.WordFrequencyUnit,
                                                          ByRef IgnoreZeroFrequencyWords As Boolean,
                                                          ByRef InferLengthLessConsonants As Boolean,
                                                          ByRef InferSwedishSegmentalProcesses As Boolean)

            SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            WordInitialOnSetClusters = New Dictionary(Of String, Double)
            WordFinalCodaClusters = New Dictionary(Of String, Double)

            'Adding empty keys in WordInitialOnSetClusters and WordFinalCodaClusters
            WordInitialOnSetClusters.Add("", 0)
            WordFinalCodaClusters.Add("", 0)

            'Starting a progress window
            Dim myProgress As New ProgressDisplay
            myProgress.Initialize(InputWordGroup.MemberWords.Count - 1, 0, "Collecting unabmigous words initial and final clusters...", 100)
            myProgress.Show()

            For word = 0 To InputWordGroup.MemberWords.Count - 1

                'Updating progress
                myProgress.UpdateProgress(word)

                'Skipping if word frequency is 0 and IgnoreZeroFrequencyWords = True
                If IgnoreZeroFrequencyWords = True And InputWordGroup.MemberWords(word).RawWordTypeFrequency = 0 Then Continue For

                Dim WordFrequencyData As Double = 0
                Select Case FrequencyWeightingUnit
                    Case WordGroup.WordFrequencyUnit.WordType
                        WordFrequencyData = 1
                    Case WordGroup.WordFrequencyUnit.RawFrequency
                        WordFrequencyData = InputWordGroup.MemberWords(word).RawWordTypeFrequency
                    Case WordGroup.WordFrequencyUnit.ZipfValue
                        WordFrequencyData = InputWordGroup.MemberWords(word).ZipfValue_Word
                    Case Else
                        Throw New NotImplementedException
                End Select


                'Reading the Phonemes array from the start to detect initial clusters
                Dim CurrentInitialCluster As New List(Of String)
                For ph = 0 To InputWordGroup.MemberWords(word).Phonemes.Count - 1

                    If Not SwedishVowels_IPA.Contains(InputWordGroup.MemberWords(word).Phonemes(ph)) Then

                        CurrentInitialCluster.Add(InputWordGroup.MemberWords(word).Phonemes(ph))

                    Else
                        'Storing the current cluster
                        Dim CurrentClusterString As String = String.Join(" ", CurrentInitialCluster)
                        If Not WordInitialOnSetClusters.ContainsKey(CurrentClusterString) Then
                            WordInitialOnSetClusters.Add(CurrentClusterString, WordFrequencyData) 'Word frequency Weighting 
                        Else
                            WordInitialOnSetClusters(CurrentClusterString) += WordFrequencyData
                        End If
                        Exit For
                    End If
                Next


                'Reading the Phonemes array from the end to detect final clusters
                Dim CurrentFinalCluster As New List(Of String)
                For InversePhonemeIndex = 0 To InputWordGroup.MemberWords(word).Phonemes.Count - 1

                    Dim ph As Integer = InputWordGroup.MemberWords(word).Phonemes.Count - 1 - InversePhonemeIndex

                    If Not SwedishVowels_IPA.Contains(InputWordGroup.MemberWords(word).Phonemes(ph)) Then
                        CurrentFinalCluster.Insert(0, InputWordGroup.MemberWords(word).Phonemes(ph))
                    Else
                        'Storing the current cluster
                        Dim CurrentClusterString As String = String.Join(" ", CurrentFinalCluster)
                        If Not WordFinalCodaClusters.ContainsKey(CurrentClusterString) Then
                            WordFinalCodaClusters.Add(CurrentClusterString, WordFrequencyData) 'Word frequency Weighting 
                        Else
                            WordFinalCodaClusters(CurrentClusterString) += WordFrequencyData
                        End If
                        Exit For
                    End If
                Next
            Next

            'Closing the progress display
            myProgress.Close()

            'Using the collected clusters and add any non added lengthless version of the clusters
            If InferLengthLessConsonants = True Then
                'Onsets
                Dim OnsetClusterCopies As New Dictionary(Of String, Double)

                For Each CurrentCluster In WordInitialOnSetClusters
                    OnsetClusterCopies.Add(CurrentCluster.Key, CurrentCluster.Value)
                Next

                For Each CurrentCluster In OnsetClusterCopies
                    Dim CurrentClusterKey As String = CurrentCluster.Key
                    Dim LengthLessKey As String = CurrentClusterKey.Replace(PhoneticLength, "")
                    If Not WordInitialOnSetClusters.ContainsKey(LengthLessKey) Then WordInitialOnSetClusters.Add(LengthLessKey, -1)
                Next

                'Codas
                Dim CodaClusterCopies As New Dictionary(Of String, Double)

                For Each CurrentCluster In WordFinalCodaClusters
                    CodaClusterCopies.Add(CurrentCluster.Key, CurrentCluster.Value)
                Next

                For Each CurrentCluster In CodaClusterCopies
                    Dim CurrentClusterKey As String = CurrentCluster.Key
                    Dim LengthLessKey As String = CurrentClusterKey.Replace(PhoneticLength, "")
                    If Not WordFinalCodaClusters.ContainsKey(LengthLessKey) Then WordFinalCodaClusters.Add(LengthLessKey, -1)
                Next
            End If

            'Creating retroflexed and nasally assimilated clusters
            If InferSwedishSegmentalProcesses = True Then

                'Onsets
                Dim OnsetClusterCopies As New Dictionary(Of String, Double)

                For Each CurrentCluster In WordInitialOnSetClusters
                    OnsetClusterCopies.Add(CurrentCluster.Key, CurrentCluster.Value)
                Next

                For Each CurrentCluster In OnsetClusterCopies

                    'Performs segmental processes, and adds to the cluster collection
                    Dim CurrentClusterKey As String = CurrentCluster.Key
                    Dim RetroFlexedKey As String = Retroflexion(CurrentClusterKey)
                    If Not WordInitialOnSetClusters.ContainsKey(RetroFlexedKey) Then WordInitialOnSetClusters.Add(RetroFlexedKey, -1)

                    Dim NasalAssimilations = NasalAssimilationAndRetroflexionCombination(CurrentClusterKey)
                    For Each CurrentItem In NasalAssimilations
                        If Not WordInitialOnSetClusters.ContainsKey(CurrentItem) Then WordInitialOnSetClusters.Add(CurrentItem, -1)
                    Next
                Next

                'Codas
                Dim CodaClusterCopies As New Dictionary(Of String, Double)

                For Each CurrentCluster In WordFinalCodaClusters
                    CodaClusterCopies.Add(CurrentCluster.Key, CurrentCluster.Value)
                Next

                For Each CurrentCluster In CodaClusterCopies

                    'Performs segmental processes, and adds to the cluster collection
                    Dim CurrentClusterKey As String = CurrentCluster.Key
                    Dim RetroFlexedKey As String = Retroflexion(CurrentClusterKey)
                    If Not WordFinalCodaClusters.ContainsKey(RetroFlexedKey) Then WordFinalCodaClusters.Add(RetroFlexedKey, -1)

                    Dim NasalAssimilations = NasalAssimilationAndRetroflexionCombination(CurrentClusterKey)
                    For Each CurrentItem In NasalAssimilations
                        If Not WordFinalCodaClusters.ContainsKey(CurrentItem) Then WordFinalCodaClusters.Add(CurrentItem, -1)
                    Next
                Next
            End If


        End Sub

        ''' <summary>
        ''' Retroflexes all coronals that can be retroflexed
        ''' </summary>
        ''' <param name="InputString"></param>
        ''' <returns></returns>
        Private Function Retroflexion(ByVal InputString As String) As String

            Dim OutputString = InputString

            OutputString = OutputString.Replace("t", "ʈ")
            OutputString = OutputString.Replace("d", "ɖ")
            OutputString = OutputString.Replace("n", "ɳ")
            OutputString = OutputString.Replace("s", "ʂ")
            OutputString = OutputString.Replace("l", "ɭ")

            Return OutputString

        End Function

        ''' <summary>
        ''' Creates nasally assimilated versions of input strings
        ''' </summary>
        ''' <param name="InputString"></param>
        ''' <returns></returns>
        Private Function NasalAssimilationAndRetroflexionCombination(ByVal InputString As String) As List(Of String)

            Dim Output As New List(Of String)

            Dim LabioVelarAssimilation As String = InputString
            Output.Add(InputString.Replace("n", "m"))
            Output.Add(InputString.Replace("n", "ɱ"))
            Output.Add(InputString.Replace("n", "ŋ"))

            Output.Add(Retroflexion(InputString.Replace("n", "ŋ")))
            Output.Add(Retroflexion(InputString.Replace("n", "ɱ")))
            Output.Add(Retroflexion(InputString.Replace("n", "ŋ")))

            Return Output

        End Function

    Public Sub ExportClustersToFile(ByVal saveDirectory As String, Optional ByRef saveFileName As String = "SyllabificationClusters",
                                                Optional BoxTitle As String = "Choose location to store the syllabification cluster export file...")

        Try


            SendInfoToLog("Attempts to save syllabification cluster data to .txt file.")

            'Choosing file location
            Dim filepath As String = Path.Combine(saveDirectory, saveFileName & ".txt")
            If Not Directory.Exists(Path.GetDirectoryName(filepath)) Then Directory.CreateDirectory(Path.GetDirectoryName(filepath))

            'Save it to file
            Dim writer As New StreamWriter(filepath, False, Text.Encoding.UTF8)

            'Witing heading
            writer.WriteLine("Word initial onset clusters")
            writer.WriteLine("Cluster" & vbTab & "Frequency")
            writer.WriteLine()
            For Each CurrentCluster In WordInitialOnSetClusters
                writer.WriteLine(CurrentCluster.Key & vbTab & CurrentCluster.Value)
            Next
            writer.WriteLine()

            writer.WriteLine("Word final coda clusters")
            writer.WriteLine("Cluster" & vbTab & "Frequency")
            writer.WriteLine()
            For Each CurrentCluster In WordFinalCodaClusters
                writer.WriteLine(CurrentCluster.Key & vbTab & CurrentCluster.Value)
            Next
            writer.WriteLine()

            writer.WriteLine("Word medial clusters")
            writer.WriteLine("Cluster" & vbTab &
                                     "Coda" & vbTab & "Coda frequency" & vbTab & "Coda Type" & vbTab & "Coda StressType" & vbTab &
                                     "Onset" & vbTab & "Onset frequency" & vbTab & "Onset Type" & vbTab & "Onset StressType")
            writer.WriteLine()
            For Each CurrentCluster In WordMedialClusters
                writer.WriteLine(CurrentCluster.Key & vbTab &
                                         CurrentCluster.Value.Coda.PhonemeString & vbTab & CurrentCluster.Value.CodaFrequency & vbTab &
                                         CurrentCluster.Value.Coda.Type.ToString & vbTab & CurrentCluster.Value.Coda.StressType & vbTab &
                                         CurrentCluster.Value.Onset.PhonemeString & vbTab & CurrentCluster.Value.OnsetFrequency & vbTab &
                                         CurrentCluster.Value.Onset.Type.ToString & vbTab & CurrentCluster.Value.Onset.StressType)
            Next

            writer.Close()

            SendInfoToLog("   Syllabification cluster data were successfully saved to .txt file: " & filepath & vbCrLf &
                                  "   A total of " & WordInitialOnSetClusters.Keys.Count + WordFinalCodaClusters.Keys.Count +
                                  WordMedialClusters.Keys.Count & " clusters were saved.")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Load new syllabification cluster data from file.
    ''' </summary>
    ''' <param name="FilePath"></param>
    ''' <returns>Returns True if loading succeded, or false if loading failed.</returns>
    Public Function LoadClusterDataFromFile(Optional ByRef FilePath As String = "") As Boolean

        Dim NoExceptionsOccurred As Boolean = True

        Try

            Dim InputData() As String

            If FilePath = "" Then
                Dim dataString As String = My.Resources.SyllabificationClusters
                dataString = dataString.Replace(vbCrLf, vbLf)
                InputData = dataString.Split(vbLf)
            Else
                SendInfoToLog("Attempts to load syllabification cluster data from .txt file.")
                InputData = System.IO.File.ReadAllLines(FilePath, Text.Encoding.UTF8)
            End If

            Dim ReadDataType As Integer = 0

            'Creating/resetting data
            WordInitialOnSetClusters = New Dictionary(Of String, Double)
            WordFinalCodaClusters = New Dictionary(Of String, Double)
            WordMedialClusters = New Dictionary(Of String, CodaOnsetCombination)

            Dim TotalLoadedClusters As Integer = 0

            For line = 0 To InputData.Length - 1

                'Skipping lines
                If InputData(line).Trim = "" Then Continue For
                If InputData(line).StartsWith("Cluster") Then Continue For

                'Setting read step
                If InputData(line).StartsWith("Word initial onset clusters") Then
                    ReadDataType = 0
                ElseIf InputData(line).StartsWith("Word final coda clusters") Then
                    ReadDataType = 1
                ElseIf InputData(line).StartsWith("Word medial clusters") Then
                    ReadDataType = 2
                End If

                Dim InputDataSplit() As String = InputData(line).Split(vbTab)
                Select Case ReadDataType
                    Case 0
                        If InputDataSplit.Length < 2 Then Continue For 'Skipping if line length is too short
                        Dim CurrentCluster As String = InputDataSplit(0)
                        Dim CurrentFrequencyData As Double = InputDataSplit(1)
                        WordInitialOnSetClusters.Add(CurrentCluster, CurrentFrequencyData)

                        TotalLoadedClusters += 1

                    Case 1
                        If InputDataSplit.Length < 2 Then Continue For 'Skipping if line length is too short
                        Dim CurrentCluster As String = InputDataSplit(0)
                        Dim CurrentFrequencyData As Double = InputDataSplit(1)
                        WordFinalCodaClusters.Add(CurrentCluster, CurrentFrequencyData)

                        TotalLoadedClusters += 1

                    Case 2
                        If InputDataSplit.Length < 9 Then Continue For 'Skipping if line length is too short
                        Dim CurrentCluster As String = InputDataSplit(0)
                        Dim CodaString As String = InputDataSplit(1)
                        Dim CodaFrequencyData As Double = InputDataSplit(2)
                        Dim CodaType As New SyllableUnitType
                        CodaType = [Enum].Parse(GetType(SyllableUnitType), InputDataSplit(3))
                        Dim CodaStressType As String = InputDataSplit(4)

                        Dim OnsetString As String = InputDataSplit(5)
                        Dim OnsetFrequencyData As Double = InputDataSplit(6)
                        Dim OnsetType As New SyllableUnitType
                        OnsetType = [Enum].Parse(GetType(SyllableUnitType), InputDataSplit(7))
                        Dim OnsetStressType As String = InputDataSplit(8)

                        Dim NewCodaOnsetCombination As New CodaOnsetCombination
                        NewCodaOnsetCombination.Coda = New SyllableUnit With {.PhonemeString = CodaString, .StressType = CodaStressType, .Type = CodaType}
                        NewCodaOnsetCombination.CodaFrequency = CodaFrequencyData
                        NewCodaOnsetCombination.Onset = New SyllableUnit With {.PhonemeString = OnsetString, .StressType = OnsetStressType, .Type = OnsetType}
                        NewCodaOnsetCombination.OnsetFrequency = OnsetFrequencyData

                        WordMedialClusters.Add(CurrentCluster, NewCodaOnsetCombination)

                        TotalLoadedClusters += 1

                End Select

            Next

            SendInfoToLog("   Syllabification cluster data were successfully loaded from .txt file: " & FilePath)

        Catch ex As Exception
            NoExceptionsOccurred = False
        End Try

        Return NoExceptionsOccurred

    End Function


    Private Sub CreateMedialClusterCombinations()

            SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name)

            'Determining the most probable medial cluster syllable boundaries using statistics from word initial onsets And word final codas
            'The definition of most probable, is the highest minimum frequency of occurence for the combined onset+coda, of all possible combinations

            'Initializing the cluster combination dictionary
            WordMedialClusters = New Dictionary(Of String, CodaOnsetCombination)

            For Each coda In WordFinalCodaClusters
                For Each Onset In WordInitialOnSetClusters

                    Dim NewClusterCombination As New CodaOnsetCombination()
                    NewClusterCombination.Coda.PhonemeString = coda.Key
                    NewClusterCombination.CodaFrequency = coda.Value
                    NewClusterCombination.Onset.PhonemeString = Onset.Key
                    NewClusterCombination.OnsetFrequency = Onset.Value

                    'Determining the lowest frequency value
                    Dim LowestFrequency As Double = coda.Value
                    If Onset.Value < LowestFrequency Then LowestFrequency = Onset.Value

                    'Creating the current cluster combination string
                    Dim ClusterCombinationKey As String = NewClusterCombination.Coda.PhonemeString & " " & NewClusterCombination.Onset.PhonemeString
                    ClusterCombinationKey = ClusterCombinationKey.Trim 'This is done, since a cluster combination including an empty onset/coda string will end up having a blank space initially/finally, which need to be removed.

                    'Adding the cluster combination
                    If Not WordMedialClusters.ContainsKey(ClusterCombinationKey) Then
                        WordMedialClusters.Add(ClusterCombinationKey, NewClusterCombination)
                    Else
                        'Adding the new cluster combination only if its lowest frequency value is higher than the lowest frequency value of the existing word
                        If NewClusterCombination.GetLowestFrequency > WordMedialClusters(ClusterCombinationKey).GetLowestFrequency Then

                            'Removing the existing cluster combination
                            WordMedialClusters.Remove(ClusterCombinationKey)

                            'Adding the new cluster combination
                            WordMedialClusters.Add(ClusterCombinationKey, NewClusterCombination)
                        End If
                    End If
                Next
            Next


        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="InputPhonemes"></param>
        ''' <param name="ContainsSegmentationGuessing">If SegmentationGuessing holds a value of True upon return, CreateSyllabification has used only part of the medial clusters data, and from it made a "qualified guess" where to put the syllable boundary.</param>
        ''' <param name="ContainsUnparsableCluster">If ContainsUnparsableCluster holds a value of True upon return, CreateSyllabification failed to parse one or more oth the clusters in the current word.</param>
        ''' <returns></returns>
        Public Function CreateSyllabification(ByRef InputPhonemes As List(Of String),
                                                  Optional ByRef ContainsSegmentationGuessing As Boolean = False,
                                                  Optional ByRef ContainsUnparsableCluster As Boolean = False) As Word.ListOfSyllables

            'Making sure that CollectWordInitialAndFinalClusters is properly run
            If WordInitialOnSetClusters Is Nothing Or WordFinalCodaClusters Is Nothing Or WordMedialClusters Is Nothing Then
                Throw New Exception("The current instance of Syllabify must have been trained on a phonetically transscribed wordgroup before " &
                                        System.Reflection.MethodInfo.GetCurrentMethod.Name & " can be used.")
            End If

            ContainsSegmentationGuessing = False
            ContainsUnparsableCluster = False

            'Moving any stress markers that occur in the input transcription to be directly concatenated with the next detected vowel. 
            'This is later used to tell which type of stress the syllable has (only carried by the nucleus vowels!)
            Dim CurrentPhoneme As Integer = 0
            Do Until CurrentPhoneme > InputPhonemes.Count - 1

                'Looks for stress markings
                If SwedishStressList.Contains(InputPhonemes(CurrentPhoneme)) Then

                    'A stress marker is detected

                    'Moves the detected stress marker to the next detected vowel
                    For targetIndex = CurrentPhoneme To InputPhonemes.Count - 1

                        If SwedishVowels_IPA.Contains(InputPhonemes(targetIndex)) Then

                            'The next vowel index is found
                            'Adding the stress marker to it
                            InputPhonemes(targetIndex) = InputPhonemes(CurrentPhoneme) & InputPhonemes(targetIndex)

                            'Removing the current phoneme (which is a stress marker), and exits to the other loop, to look for more stress markers
                            InputPhonemes.RemoveAt(CurrentPhoneme)
                            Exit For
                        End If
                    Next
                Else
                    'The current phoneme is not a stress marker, increases the value of CurrentPhoneme and continues the loop
                    CurrentPhoneme += 1
                End If
            Loop


            'Getting the syllable nuclei indices, defined as vowels
            Dim VowelIndices As New List(Of Integer)
            For ph = 0 To InputPhonemes.Count - 1
                If SwedishVowels_IPA.Contains(InputPhonemes(ph).Replace(IpaMainStress, "").Replace(IpaMainSwedishAccent2, "").Replace(IpaSecondaryStress, "")) Then
                    VowelIndices.Add(ph)
                End If
            Next


            Dim CurrentSyllabification As New List(Of SyllableUnit)

            'A. Getting the word initial cluster
            'Only if there is any
            Dim OnsetLength As Integer = VowelIndices(0)
            If OnsetLength > 0 Then

                Dim localWordInitialCluster As New SyllableUnit
                localWordInitialCluster.PhonemeString = String.Join(" ", InputPhonemes.ToArray, 0, OnsetLength)

                'Adding only if the length is higher than 0
                If localWordInitialCluster.PhonemeString.Length > 0 Then
                    'Setting item type
                    localWordInitialCluster.Type = SyllableUnitType.Onset
                    'Adding the item
                    CurrentSyllabification.Add(localWordInitialCluster)
                End If
            End If


            'B Getting the first vowel
            Dim FirstWovel As New SyllableUnit
            Dim FirstVowelFullString As String = InputPhonemes(VowelIndices(0))
            'Stores a stress reduced form of the vowel
            FirstWovel.PhonemeString = FirstVowelFullString.Replace(IpaMainStress, "").Replace(IpaMainSwedishAccent2, "").Replace(IpaSecondaryStress, "")

            'Continues only if the length is higher than 0
            If FirstWovel.PhonemeString.Length > 0 Then

                'Determines the stress and tone accent type of the vowel
                If FirstVowelFullString.Contains(IpaMainStress) Then
                    FirstWovel.StressType = IpaMainStress
                ElseIf FirstVowelFullString.Contains(IpaMainSwedishAccent2) Then
                    FirstWovel.StressType = IpaMainSwedishAccent2
                ElseIf FirstVowelFullString.Contains(IpaSecondaryStress) Then
                    FirstWovel.StressType = IpaSecondaryStress
                Else
                    'Unstressed
                    'FirstWovel.StressType = "", which it already is from the initiation
                End If

                'Setting item type
                FirstWovel.Type = SyllableUnitType.Nucleus

                'Adding the item
                CurrentSyllabification.Add(FirstWovel)
            End If


            'C. Parsing word medial clusters into syllable coda + syllable onset
            For syllable = 0 To VowelIndices.Count - 2

                'Getting the phonemes between the last nucleus and the next nucleus (i.e. coda of the current, and onset of the next, syllable) 
                'Testing if there are any phonemes between theses vowels/nuclie. Skips to the next vowel if there are no.
                Dim Clusterlength As Integer = (VowelIndices(syllable + 1) - 1) - (VowelIndices(syllable))

                Select Case Clusterlength
                    Case 0
                            'Skips if there is no consonant between the vowels
                    Case 1
                        'If there is one consonant there, it is assigned to the right syllable, unless this is an illegal position (determined by a threshold value of only 1 % of the likelyhood of the left syllable, based on the initial and final cluster frequencies (that is, if it is short or DuplicateAmbiSyllabicLongConsonants = false)
                        'If it is long and DuplicateAmbiSyllabicLongConsonants = True, it is duplicated in both syllables, a long copy in the left syllable coda, and a short copy in the next syllable onset

                        Dim CurrentMedialConsonant As String = String.Join(" ", InputPhonemes.ToArray, VowelIndices(syllable) + 1, Clusterlength)

                        If CurrentMedialConsonant.Contains(PhoneticLength) And DuplicateAmbiSyllabicLongConsonants = True Then

                            Dim MedialSingleLongConsonant As New SyllableUnit
                            MedialSingleLongConsonant.PhonemeString = CurrentMedialConsonant
                            'Adding only if the length is higher than 0
                            If MedialSingleLongConsonant.PhonemeString.Length > 0 Then

                                'Setting item type
                                MedialSingleLongConsonant.Type = SyllableUnitType.Coda
                                'Adding the item
                                CurrentSyllabification.Add(MedialSingleLongConsonant)
                            End If

                            Dim MedialSingleShortConsonant As New SyllableUnit
                            MedialSingleShortConsonant.PhonemeString = CurrentMedialConsonant.Replace(PhoneticLength, "")
                            'Adding only if the length is higher than 0
                            If MedialSingleShortConsonant.PhonemeString.Length > 0 Then

                                'Setting item type
                                MedialSingleShortConsonant.Type = SyllableUnitType.Onset
                                'Adding the item
                                CurrentSyllabification.Add(MedialSingleShortConsonant)
                            End If

                        Else

                            'Determines if the consonant should be put in the left or the right syllable
                            Dim LeftSyllableProbability As Integer = 0
                            If WordFinalCodaClusters.ContainsKey(CurrentMedialConsonant) Then LeftSyllableProbability = WordFinalCodaClusters(CurrentMedialConsonant)

                            Dim RightSyllableProbability As Integer = 0
                            If WordInitialOnSetClusters.ContainsKey(CurrentMedialConsonant) Then RightSyllableProbability = WordInitialOnSetClusters(CurrentMedialConsonant)

                            'Determines which syllable it should be put in
                            Dim PutItInTheRightSyllable As Boolean = True 'The right syllable is used as default if probabilities are equal (Cf maximazation of onset principle)
                            'Overrides the default right syllable choice, if the probability of the left syllable, based on word initial and word final consonant occurences, is at least a 100 times larger.
                            If LeftSyllableProbability / 100 > RightSyllableProbability Then
                                PutItInTheRightSyllable = False
                            End If

                            'Changing it to the left syllable if, the syllable is long (and also DuplicateAmbiSyllabicLongConsonants = False (Other wise the code would not be here, but instead before the Else statement above...))
                            If CurrentMedialConsonant.Contains(PhoneticLength) Then PutItInTheRightSyllable = False

                            'Adding the consonant
                            Dim MedialSingleShortConsonant As New SyllableUnit
                            MedialSingleShortConsonant.PhonemeString = CurrentMedialConsonant
                            'Adding only if the length is higher than 0
                            If MedialSingleShortConsonant.PhonemeString.Length > 0 Then

                                'Setting item type
                                If PutItInTheRightSyllable = True Then
                                    'Putting it the default right syllable
                                    MedialSingleShortConsonant.Type = SyllableUnitType.Onset
                                Else
                                    'Putting it in the left syllable
                                    MedialSingleShortConsonant.Type = SyllableUnitType.Coda
                                End If

                                'Adding the item
                                CurrentSyllabification.Add(MedialSingleShortConsonant)
                            End If

                        End If

                    Case Else
                        'We've found a medial consonant cluster (more than one consonant)

                        Dim CurrentMedialCluster As String = String.Join(" ", InputPhonemes.ToArray, VowelIndices(syllable) + 1, Clusterlength)

                        If WordMedialClusters.ContainsKey(CurrentMedialCluster) Then

                            'Adding the medial clusters
                            'Adding the coda item
                            'Adding only if the length is higher than 0
                            If WordMedialClusters(CurrentMedialCluster).Coda.PhonemeString.Length > 0 Then
                                CurrentSyllabification.Add(WordMedialClusters(CurrentMedialCluster).Coda)
                            End If

                            'Adding the onset
                            'Adding only if the length is higher than 0
                            If WordMedialClusters(CurrentMedialCluster).Onset.PhonemeString.Length > 0 Then
                                CurrentSyllabification.Add(WordMedialClusters(CurrentMedialCluster).Onset)
                            End If

                        Else

                            'Testing to see how much is left of the cluster when the longest acceptable coda is selected
                            Dim LongestParseUsingInitialBit As Integer = 0
                            Dim BestFitCoda As String = ""
                            For Each item In WordFinalCodaClusters
                                If CurrentMedialCluster.StartsWith(item.Key) Then
                                    If item.Key.Length > LongestParseUsingInitialBit Then
                                        LongestParseUsingInitialBit = item.Key.Length
                                        BestFitCoda = item.Key
                                    End If
                                End If
                            Next

                            'Testing to see how much is left of the cluster when the longest acceptable onset is selected
                            Dim LongestParseUsingFinalBit As Integer = 0
                            Dim BestFitOnset As String = ""
                            For Each item In WordInitialOnSetClusters
                                If CurrentMedialCluster.EndsWith(item.Key) Then
                                    If item.Key.Length > LongestParseUsingFinalBit Then
                                        LongestParseUsingFinalBit = item.Key.Length
                                        BestFitOnset = item.Key
                                    End If
                                End If
                            Next

                            If LongestParseUsingInitialBit = 0 And LongestParseUsingFinalBit = 0 Then

                                'Adding an unparsable cluster
                                Dim UnparsableCluster As New SyllableUnit
                                UnparsableCluster.PhonemeString = "~"
                                UnparsableCluster.Type = SyllableUnitType.Unparsable
                                CurrentSyllabification.Add(UnparsableCluster)
                                ContainsUnparsableCluster = True

                            Else

                                'Splitting the current medial cluster into its phoneme parts
                                Dim CurrentMedialClusterSplit() As String = CurrentMedialCluster.Split(" ")

                                ContainsSegmentationGuessing = True

                                'Selects the best fit, or the rightmost bit if both parse types are equally good
                                If LongestParseUsingInitialBit > LongestParseUsingFinalBit Then
                                    'Choosing the initial bit

                                    Dim NewCodaCluster As New SyllableUnit
                                    NewCodaCluster.PhonemeString = BestFitCoda

                                    'Adding only if the length is higher than 0
                                    If NewCodaCluster.PhonemeString.Length > 0 Then
                                        'Setting item type
                                        NewCodaCluster.Type = SyllableUnitType.Coda
                                        'Adding the item
                                        CurrentSyllabification.Add(NewCodaCluster)
                                    End If

                                    Dim NewOnsetCluster As New SyllableUnit
                                    NewOnsetCluster.PhonemeString = String.Join(" ", CurrentMedialClusterSplit, (BestFitCoda.Split(" ").Count), CurrentMedialClusterSplit.Count - (BestFitCoda.Split(" ").Count))
                                    'Adding only if the length is higher than 0
                                    If NewOnsetCluster.PhonemeString.Length > 0 Then
                                        'Setting item type
                                        NewOnsetCluster.Type = SyllableUnitType.Onset
                                        'Adding the item
                                        CurrentSyllabification.Add(NewOnsetCluster)
                                    End If


                                Else
                                    'Choosing the final bit
                                    Dim NewCodaCluster As New SyllableUnit
                                    NewCodaCluster.PhonemeString = String.Join(" ", CurrentMedialClusterSplit, 0, CurrentMedialClusterSplit.Count - (BestFitOnset.Split(" ").Count))

                                    'Adding only if the length is higher than 0
                                    If NewCodaCluster.PhonemeString.Length > 0 Then
                                        'Setting item type
                                        NewCodaCluster.Type = SyllableUnitType.Coda
                                        'Adding the item
                                        CurrentSyllabification.Add(NewCodaCluster)
                                    End If

                                    Dim NewOnsetCluster As New SyllableUnit
                                    NewOnsetCluster.PhonemeString = BestFitOnset

                                    'Adding only if the length is higher than 0
                                    If NewOnsetCluster.PhonemeString.Length > 0 Then
                                        'Setting item type
                                        NewOnsetCluster.Type = SyllableUnitType.Onset
                                        'Adding the item
                                        CurrentSyllabification.Add(NewOnsetCluster)
                                    End If
                                End If
                            End If

                        End If

                End Select


                'Getting the vowel after the medial cluster
                Dim PostMedialClusterVowel As New SyllableUnit
                Dim PostMedialClusterVowelFullString As String = InputPhonemes(VowelIndices(syllable + 1))
                'Stores a stress reduced form of the vowel
                PostMedialClusterVowel.PhonemeString = PostMedialClusterVowelFullString.Replace(IpaMainStress, "").Replace(IpaMainSwedishAccent2, "").Replace(IpaSecondaryStress, "")

                'Continues only if the length is higher than 0
                If PostMedialClusterVowel.PhonemeString.Length > 0 Then

                    'Determines the stress and tone accent type of the vowel
                    If PostMedialClusterVowelFullString.Contains(IpaMainStress) Then
                        PostMedialClusterVowel.StressType = IpaMainStress
                    ElseIf PostMedialClusterVowelFullString.Contains(IpaMainSwedishAccent2) Then
                        PostMedialClusterVowel.StressType = IpaMainSwedishAccent2
                    ElseIf PostMedialClusterVowelFullString.Contains(IpaSecondaryStress) Then
                        PostMedialClusterVowel.StressType = IpaSecondaryStress
                    Else
                        'Unstressed
                        'FirstWovel.StressType = "", which it already is from the initiation
                    End If

                    'Setting item type
                    PostMedialClusterVowel.Type = SyllableUnitType.Nucleus
                    'Adding the item
                    CurrentSyllabification.Add(PostMedialClusterVowel)
                End If
            Next

            'D. Getting the word final syllable coda, if there is one
            Dim CodaLength As Integer = InputPhonemes.Count - 1 - VowelIndices(VowelIndices.Count - 1)
            If CodaLength > 0 Then

                Dim localWordFinalCluster As New SyllableUnit
                localWordFinalCluster.PhonemeString = String.Join(" ", InputPhonemes.ToArray,
                                                    (VowelIndices(VowelIndices.Count - 1) + 1),
                                                   CodaLength)

                'Adding only if the length is higher than 0
                If localWordFinalCluster.PhonemeString.Length > 0 Then
                    'Setting item type
                    localWordFinalCluster.Type = SyllableUnitType.Coda

                    'Adding the item
                    CurrentSyllabification.Add(localWordFinalCluster)
                End If

            End If



            'Creates a syllables instance 
            Dim NewSyllables As New Word.ListOfSyllables
            Dim NewSyllable As New Word.Syllable
            For item = 0 To CurrentSyllabification.Count - 1

                Select Case CurrentSyllabification(item).Type
                    Case SyllableUnitType.Onset
                        'Creates a new syllable, and adds the current phonemes
                        NewSyllable = New Word.Syllable

                        'Getting a phoneme array
                        Dim CurrentPhonemes() As String = CurrentSyllabification(item).PhonemeString.Split(" ")
                        For ph = 0 To CurrentPhonemes.Count - 1
                            NewSyllable.Phonemes.Add(CurrentPhonemes(ph))
                        Next
                        NewSyllable.LengthOfOnset = CurrentPhonemes.Count

                    Case SyllableUnitType.Nucleus

                        'Creating a new syllable if there is no active syllable
                        If NewSyllable Is Nothing Then NewSyllable = New Word.Syllable

                        'Adding the nucleus to the last detected syllable coda
                        Dim CurrentPhonemes() As String = CurrentSyllabification(item).PhonemeString.Split(" ")
                        'Tests that there is only one phoneme in the nucleus string
                        If CurrentPhonemes.Count > 1 Then MsgBox("Syllable nucleus with more than one phoneme detected in word: " & String.Join(" ", InputPhonemes.ToArray))
                        For ph = 0 To CurrentPhonemes.Count - 1
                            NewSyllable.Phonemes.Add(CurrentPhonemes(ph))
                        Next

                        'Sets the stress type and tone of the current syllable/word
                        If CurrentSyllabification(item).StressType = IpaMainStress Then
                            NewSyllables.MainStressSyllableIndex = NewSyllables.Count + 2
                            NewSyllables.Tone = 1
                            NewSyllable.IsStressed = True
                        ElseIf CurrentSyllabification(item).StressType = IpaMainSwedishAccent2 Then
                            NewSyllables.MainStressSyllableIndex = NewSyllables.Count + 2
                            NewSyllables.Tone = 2
                            NewSyllable.IsStressed = True
                        ElseIf CurrentSyllabification(item).StressType = IpaSecondaryStress Then
                            NewSyllables.SecondaryStressSyllableIndex = NewSyllables.Count + 2
                            NewSyllable.IsStressed = True
                            NewSyllable.CarriesSecondaryStress = True
                        Else
                            NewSyllable.IsStressed = False
                        End If

                        'Sets the index of nucleus in the current syllable
                        NewSyllable.IndexOfNuclues = NewSyllable.Phonemes.Count '1-based

                        'Looks ahead to see whether a coda is coming. If not, the syllable is stored
                        If (item + 1) < CurrentSyllabification.Count Then
                            If Not CurrentSyllabification(item + 1).Type = SyllableUnitType.Coda Then
                                NewSyllables.Add(NewSyllable)
                                NewSyllable = New Word.Syllable
                                NewSyllable = Nothing
                            End If
                        Else
                            'Storing the last syllable
                            NewSyllables.Add(NewSyllable)
                            NewSyllable = New Word.Syllable
                            NewSyllable = Nothing
                        End If

                    Case SyllableUnitType.Coda

                        'Creating a new syllable if there is no active syllable
                        If NewSyllable Is Nothing Then
                            MsgBox("A coda without preceding nucleus has been detected in word: " & String.Join(" ", InputPhonemes.ToArray))
                            NewSyllable = New Word.Syllable
                        End If

                        'Getting a phoneme array
                        Dim CurrentPhonemes() As String = CurrentSyllabification(item).PhonemeString.Split(" ")
                        For ph = 0 To CurrentPhonemes.Count - 1
                            NewSyllable.Phonemes.Add(CurrentPhonemes(ph))
                        Next
                        NewSyllable.LengthOfCoda = CurrentPhonemes.Count

                        'Stores the completed syllable
                        NewSyllables.Add(NewSyllable)
                        NewSyllable = New Word.Syllable
                        NewSyllable = Nothing

                    Case SyllableUnitType.Unparsable 'This is considered to be a coda

                        'Creating a new syllable if there is no active syllable
                        If NewSyllable Is Nothing Then NewSyllable = New Word.Syllable

                        'Getting a phoneme array
                        Dim CurrentPhonemes() As String = CurrentSyllabification(item).PhonemeString.Split(" ")
                        For ph = 0 To CurrentPhonemes.Count - 1
                            NewSyllable.Phonemes.Add(CurrentPhonemes(ph))
                        Next
                        NewSyllable.LengthOfCoda = CurrentPhonemes.Count

                        'Stores the completed syllable
                        NewSyllables.Add(NewSyllable)
                        NewSyllable = New Word.Syllable
                        NewSyllable = Nothing

                End Select

            Next

            Return NewSyllables

        End Function




        Public Class CodaOnsetCombination
            Property Coda As New SyllableUnit
            Property Onset As New SyllableUnit
            Property CodaFrequency As Double
            Property OnsetFrequency As Double

            Public Sub New()
                Coda.Type = SyllableUnitType.Coda
                Onset.Type = SyllableUnitType.Onset
            End Sub

            'Returns the lowest frequency value of CodaFrequency and OnsetFrequency
            Public Function GetLowestFrequency() As Double

                Dim LowestFrequency As Double = CodaFrequency
                If OnsetFrequency < CodaFrequency Then LowestFrequency = OnsetFrequency
                Return LowestFrequency

            End Function
        End Class

        Public Class SyllableUnit
            Public Property PhonemeString As String
            Public Property Type As SyllableUnitType
            Public Property StressType As String = "" 'Can be used to store a stress marker that should go with the PhonemeString
        End Class

        Enum SyllableUnitType
            Onset
            Nucleus
            Coda
            Unparsable
        End Enum

    End Class

