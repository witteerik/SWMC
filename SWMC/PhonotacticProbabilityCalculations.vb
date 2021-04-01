'This software is available under the following license:
'MIT/X11 License
'
'Copyright (c) 2021 Erik Witte
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the ''Software''), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED ''AS IS'', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Imports System.IO
Imports System.Windows.Forms


Public Enum PhonotacticTransitionDataDirections
    CreateTransitionMatrix
    RetrievePhonotacticProbabilities
End Enum

''' <summary>
''' This class calculates the Stress and syllable structure based normalized phonotactic probability (SSPP) described by Witte and Köbler 2019.
''' </summary>
Public Class PhonoTactics

        Public ReadOnly Property WordStartMarker As String
        Public ReadOnly Property WordEndMarker As String
        Public ReadOnly Property UseAlternateSyllabification As Boolean
        Public ReadOnly Property ReduceSecondaryStress As Boolean
        Public ReadOnly Property UseSurfacePhones As Boolean
        Public ReadOnly Property CalculationType As PhonoTacticCalculationTypes
        Public ReadOnly Property FrequencyUnit As FrequencyUnits
        Public ReadOnly Property PhonoTacticGroupingType As PhonoTacticGroupingTypes
        Public ReadOnly Property ExcludeForeignWordsFromMatrixCreation As Boolean
        Public ReadOnly Property Do_z_ScoreTransformation As Boolean

        Private Property TransitionalProbabilities As PhonoTacticProbabilities

        Enum PhonoTacticCalculationTypes
            PhonotacticProbability
            PhonotacticPredictability
        End Enum

        Enum FrequencyUnits
            WordCount
            PhonemeCount
        End Enum

        Enum PhonoTacticGroupingTypes
            Standard 'Distinguishes between Stressed, Ante_Stress, Inter_Stress and Post_Stress syllables (There were previosly also other types... but now the enumerator has only one member and could thus be removed.)
        End Enum

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="SetCalculationType"></param>
        ''' <param name="SetWordStartMarker"></param>
        ''' <param name="SetWordEndMarker"></param>
        ''' <param name="SetUseAlternateSyllabification"></param>
        ''' <param name="SetUseSurfacePhones">If set to true, the quality of all length-reduced underlyingly long vowel phonemes will be reduced to that of their short allophones.</param>
        ''' <param name="SetReduceSecondaryStress">Only has effect if SetUseSurfacePhonesis = True. If set to true, secondarily stressed syllables will be treated as non stressed syllables, with their phonetic length ignored, and vowel phonemes reduced to their short allophones.</param>
        Public Sub New(ByRef SetCalculationType As PhonoTacticCalculationTypes,
                          Optional ByRef SetWordStartMarker As String = "*", Optional ByRef SetWordEndMarker As String = "_",
                           Optional ByRef SetUseAlternateSyllabification As Boolean = True,
                           Optional ByRef SetUseSurfacePhones As Boolean = False,
                           Optional ByRef SetReduceSecondaryStress As Boolean = False,
                           Optional ByVal SetFrequencyUnit As FrequencyUnits = FrequencyUnits.PhonemeCount,
                           Optional ByVal SetPhonoTacticGroupingType As PhonoTacticGroupingTypes = PhonoTacticGroupingTypes.Standard,
                           Optional ByVal SetExcludeForeignWordsFromMatrixCreation As Boolean = True,
                           Optional ByVal SetDo_z_ScoreTransformation As Boolean = False)

            UseAlternateSyllabification = SetUseAlternateSyllabification
            ReduceSecondaryStress = SetReduceSecondaryStress
            UseSurfacePhones = SetUseSurfacePhones
            WordStartMarker = SetWordStartMarker
            WordEndMarker = SetWordEndMarker
            FrequencyUnit = SetFrequencyUnit
            PhonoTacticGroupingType = SetPhonoTacticGroupingType
            ExcludeForeignWordsFromMatrixCreation = SetExcludeForeignWordsFromMatrixCreation
            Do_z_ScoreTransformation = SetDo_z_ScoreTransformation

            If UseSurfacePhones = False Then ReduceSecondaryStress = False

            CalculationType = SetCalculationType

            'Initializes the data set
            TransitionalProbabilities = New PhonoTacticProbabilities(SetPhonoTacticGroupingType)

        End Sub


        ''' <summary>
        ''' Adds data for a phoneme transition.
        ''' </summary>
        ''' <param name="CurrentPhoneme">The current phoneme.</param>
        ''' <param name="NextPhoneme">The following phoneme</param>
        ''' <param name="WordFrequency">The word frequency data for the current word.</param>
        ''' <param name="SyllabicPosition">Indicate whether the current phoneme resides in a stressed syllable or, in an unstressed syllable and its position in relation to other stressed syllables.</param>
        ''' <param name="SyllablePart">Indicate whether the current phoneme resides either in a syllable onset, or in a nuclues or coda.</param>
        Private Sub AddProbabilityData(ByVal CurrentPhoneme As String, ByVal NextPhoneme As String, ByVal WordFrequency As Double,
                                          ByVal SyllabicPosition As SyllabicPositions, ByVal SyllablePart As SyllableParts,
                                           ByVal DoNotAddDataWithZeroFrequency As Boolean)

            'Not adding if DoNotAddDataWithZeroFrequency = true and the Frequency value is zero
            If DoNotAddDataWithZeroFrequency = True And WordFrequency = 0 Then Exit Sub

            'Adding Current phoneme if it doesn't exists, and accumulates its frequency value
            If Not TransitionalProbabilities(SyllabicPosition)(SyllablePart).ContainsKey(CurrentPhoneme) Then
                TransitionalProbabilities(SyllabicPosition)(SyllablePart).Add(CurrentPhoneme, New TransitionPhonemes(CalculationType))
            End If
            'TransitionalProbabilities(Stress)(SyllablePosition)(CurrentPhoneme).FrequencyData += WordFrequency ' Phoneme FrequencyData is summed later instead

            'Adding Next phoneme to the transition phonemes, if it doesn't exists, and accumulates its frequency value
            If Not TransitionalProbabilities(SyllabicPosition)(SyllablePart)(CurrentPhoneme).ContainsKey(NextPhoneme) Then
                Dim NewProbabilityData As New TransitionPhonemes.ProbabilityData
                NewProbabilityData.FrequencyData = WordFrequency
                NewProbabilityData.OccurenceCount = 1
                TransitionalProbabilities(SyllabicPosition)(SyllablePart)(CurrentPhoneme).Add(NextPhoneme, NewProbabilityData)

            Else
                TransitionalProbabilities(SyllabicPosition)(SyllablePart)(CurrentPhoneme)(NextPhoneme).FrequencyData += WordFrequency
                TransitionalProbabilities(SyllabicPosition)(SyllablePart)(CurrentPhoneme)(NextPhoneme).OccurenceCount += 1

            End If

            'Also counting CurrentPhoneme occurences
            TransitionalProbabilities(SyllabicPosition)(SyllablePart)(CurrentPhoneme).OccurenceCount += 1

        End Sub

        ''' <summary>
        ''' Adds data for a phoneme transition.
        ''' </summary>
        ''' <param name="CurrentPhoneme">The current phoneme.</param>
        ''' <param name="NextPhoneme">The following phoneme</param>
        ''' <param name="SyllabicPosition">Indicate whether the current phoneme resides in a stressed syllable or, in an unstressed syllable and its position in relation to other stressed syllables.</param>
        ''' <param name="SyllablePart">Indicate whether the current phoneme resides either in a syllable onset, or in a nuclues or coda.</param>
        Private Function GetProbabilityData(ByVal CurrentPhoneme As String, ByVal NextPhoneme As String,
                                          ByVal SyllabicPosition As SyllabicPositions, ByVal SyllablePart As SyllableParts) As Double

            'Checking that the current data exists in the probability collection, returns 0 otherwise
            If Not TransitionalProbabilities.ContainsKey(SyllabicPosition) Then
                Return 0
            End If
            If Not TransitionalProbabilities(SyllabicPosition).ContainsKey(SyllablePart) Then
                Return 0
            End If
            If Not TransitionalProbabilities(SyllabicPosition)(SyllablePart).ContainsKey(CurrentPhoneme) Then
                Return 0
            End If
            If Not TransitionalProbabilities(SyllabicPosition)(SyllablePart)(CurrentPhoneme).ContainsKey(NextPhoneme) Then
                Return 0
            End If

            Select Case CalculationType
                Case PhonoTacticCalculationTypes.PhonotacticProbability
                    Return TransitionalProbabilities(SyllabicPosition)(SyllablePart)(CurrentPhoneme)(NextPhoneme).TransitionalProbability

                Case PhonoTacticCalculationTypes.PhonotacticPredictability
                    Return TransitionalProbabilities(SyllabicPosition)(SyllablePart)(CurrentPhoneme)(NextPhoneme).TransitionPredictability

                Case Else
                    Throw New NotImplementedException

            End Select

        End Function

    ''' <summary>
    ''' This sub either creates a phonotactic transition matrix from an input word group or retrieves the phonotactic probability data for each word using a pre-existing transition matrix.
    ''' The type of use is selected using the TransitionDataDirection argument.
    ''' </summary>
    ''' <param name="TransitionDataDirection"></param>
    ''' <param name="InputWordGroup"></param>
    ''' <param name="OutputFolder"></param>
    ''' <param name="OutputFileName"></param>
    ''' <param name="ExportDataToFile"></param>
    Public Sub TransitionData(ByVal TransitionDataDirection As PhonotacticTransitionDataDirections, ByVal InputWordGroup As WordGroup,
                                      Optional ByRef OutputFolder As String = "", Optional ByRef OutputFileName As String = "", Optional ExportDataToFile As Boolean = True)

        SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " TransitionDataDirection:" & TransitionDataDirection.ToString)

        'Going through each phoneme in each word in the input word group and stores its frequency data

        'Starting a progress window
        Dim myProgress As New ProgressDisplay
        If TransitionDataDirection = PhonotacticTransitionDataDirections.RetrievePhonotacticProbabilities Then
            myProgress.Initialize(InputWordGroup.MemberWords.Count - 1, 0, "Calculating phonotactic probability...", 100)
        Else
            myProgress.Initialize(InputWordGroup.MemberWords.Count - 1, 0, "Collecting phoneme transition data...", 100)
        End If

        myProgress.Show()
        Dim ProcessedWordsCount As Integer = 0

        For Each CurrentWord In InputWordGroup.MemberWords

            'Updating progress
            myProgress.UpdateProgress(ProcessedWordsCount)
            ProcessedWordsCount += 1

            'Creating lists to hold the phoneme equivalents to the added probability data, and their probability data
            If TransitionDataDirection = PhonotacticTransitionDataDirections.RetrievePhonotacticProbabilities Then
                CurrentWord.PP = New List(Of Double)
                CurrentWord.PP_Phonemes = New List(Of String)
            End If

            'Does not add probabilities based on foreign words
            If TransitionDataDirection = PhonotacticTransitionDataDirections.CreateTransitionMatrix Then
                If ExcludeForeignWordsFromMatrixCreation = True And CurrentWord.ForeignWord = True Then Continue For
            End If

            'Sets which syllables that should be read
            Dim SyllablesToRead As Word.ListOfSyllables
            If UseAlternateSyllabification = True Then
                SyllablesToRead = CurrentWord.Syllables_AlternateSyllabification
            Else
                SyllablesToRead = CurrentWord.Syllables
            End If


            'Determines the first and last stressed syllable (one based) indices
            Dim FirstStressedSyllableOneBasedIndex As Integer = SyllablesToRead.MainStressSyllableIndex
            If Not SyllablesToRead.SecondaryStressSyllableIndex = 0 Then
                If SyllablesToRead.SecondaryStressSyllableIndex < SyllablesToRead.MainStressSyllableIndex Then FirstStressedSyllableOneBasedIndex = SyllablesToRead.SecondaryStressSyllableIndex
            End If

            Dim LastStressedSyllableOneBasedIndex As Integer = SyllablesToRead.MainStressSyllableIndex
            If Not SyllablesToRead.SecondaryStressSyllableIndex = 0 Then
                If SyllablesToRead.SecondaryStressSyllableIndex > SyllablesToRead.MainStressSyllableIndex Then LastStressedSyllableOneBasedIndex = SyllablesToRead.SecondaryStressSyllableIndex
            End If
            Try

                For syllable = 0 To SyllablesToRead.Count - 1

                    'Determines the syllabic position of the current syllable
                    Dim CurrentSyllabicPosition As SyllabicPositions
                    Select Case PhonoTacticGroupingType

                        Case PhonoTacticGroupingTypes.Standard

                            'Determines if the syllable is stressed
                            Dim SyllableIsStressed As Boolean = False
                            If SyllablesToRead(syllable).IsStressed = True Then SyllableIsStressed = True
                            'However(!), if it is secondarily stressed and ReduceSecondaryStress = True, CurrentSyllableStress is changed back to Unstressed
                            If ReduceSecondaryStress = True Then
                                If SyllablesToRead(syllable).CarriesSecondaryStress = True Then SyllableIsStressed = False
                            End If

                            If SyllableIsStressed = True Then
                                CurrentSyllabicPosition = SyllabicPositions.Stressed
                            Else

                                'Determines the position of the unstressed syllable in relation to the stressed syllable/s, and also to the word end

                                'Determines where in realtion to the stressed syllable indices the syllable reside
                                If syllable + 1 < FirstStressedSyllableOneBasedIndex Then
                                    'We're before the first stressed syllable

                                    CurrentSyllabicPosition = SyllabicPositions.Ante_Stress

                                ElseIf syllable + 1 > LastStressedSyllableOneBasedIndex Then
                                    'We're after the last stressed syllable
                                    CurrentSyllabicPosition = SyllabicPositions.Post_Stress

                                Else
                                    'We're in between two stressed syllables
                                    CurrentSyllabicPosition = SyllabicPositions.Inter_Stress

                                End If
                            End If

                        Case Else
                            Throw New NotImplementedException

                    End Select


                    Dim phonemeIndex As Integer = -1
                    For i = 0 To SyllablesToRead(syllable).Phonemes.Count

                        'Determines if the phoneme is in the onset or nucleus+coda part
                        Dim CurrentSyllablePart As SyllableParts = SyllableParts.Onset

                        'Changes it to nucleus, if it is a nucleus
                        If phonemeIndex = SyllablesToRead(syllable).IndexOfNuclues - 1 Then CurrentSyllablePart = SyllableParts.Nucleus

                        'Changes it to coda, if it is a coda
                        If phonemeIndex > SyllablesToRead(syllable).IndexOfNuclues - 1 Then CurrentSyllablePart = SyllableParts.Coda


                        'Gets the current sound
                        Dim CurrentSound As String = ""
                        If phonemeIndex = -1 Then

                            'We're (either) before the word (or on a syllable boundary before a syllable)
                            If syllable = 0 Then
                                CurrentSound = WordStartMarker
                            Else
                                phonemeIndex += 1
                                Continue For 'Skipping to position 0 if its not the word start (but instead the start of a non word-initial syllable)
                                '    CurrentSound = IpaSyllableBoundary
                            End If

                        Else
                            If UseSurfacePhones = True Then
                                CurrentSound = SyllablesToRead(syllable).SurfacePhones(phonemeIndex, ReduceSecondaryStress)
                            Else
                                CurrentSound = SyllablesToRead(syllable).Phonemes(phonemeIndex)
                            End If
                        End If

                        'Getting the next sound, and changing it to the word end marker if it is returned empty (as GetNextSound does for the sound after the last sound)
                        Dim NextSound As String = ""
                        If phonemeIndex = SyllablesToRead(syllable).Phonemes.Count - 1 Then
                            'We're on the last phoneme in the syllable. The next will either be a syllable boundary or a word end marker
                            If syllable = SyllablesToRead.Count - 1 Then
                                NextSound = WordEndMarker
                            Else

                                'If we're on a syllable boundary.
                                'The next phoneme is read from the following syllable, and a syllable boundary marker is inserted

                                If UseSurfacePhones = True Then
                                    NextSound = IpaSyllableBoundary & SyllablesToRead(syllable + 1).SurfacePhones(0, ReduceSecondaryStress)
                                Else
                                    NextSound = IpaSyllableBoundary & SyllablesToRead(syllable + 1).Phonemes(0)
                                End If

                            End If

                        Else
                            'Gets the next sound in the syllable (which will also be surface form / lengthreduced if this is true for the current sound)
                            If UseSurfacePhones = True Then
                                NextSound = SyllablesToRead(syllable).SurfacePhones(phonemeIndex + 1, ReduceSecondaryStress)
                            Else
                                NextSound = SyllablesToRead(syllable).Phonemes(phonemeIndex + 1)
                            End If
                        End If

                        'Stores or sets probability data
                        If TransitionDataDirection = PhonotacticTransitionDataDirections.CreateTransitionMatrix Then
                            'Adding frequency info to the probability data
                            Dim FrequencyValue As Double = 0

                            Select Case FrequencyUnit
                                Case FrequencyUnits.WordCount
                                    'Logs are taken straigth away (Log10(token count) => word frequency weighted phoneme count)
                                    FrequencyValue = Math.Log10(CurrentWord.RawWordTypeFrequency + 2) '+2 is used to avoid frequency of 0, which would occur if raw frequency=1 (Log10(1)=0). Thus all freqquency data is slightly elevated, by 2

                                Case FrequencyUnits.PhonemeCount
                                    'Phoneme frequencies are summed first (1*token count => total phoneme count in the corpus), Logs are calculated later
                                    FrequencyValue = CurrentWord.RawWordTypeFrequency

                                Case Else
                                    Throw New NotImplementedException

                            End Select


                            Me.AddProbabilityData(CurrentSound, NextSound, FrequencyValue,
                                              CurrentSyllabicPosition, CurrentSyllablePart, True)

                        Else
                            'Adding probability to the word probability array
                            CurrentWord.PP.Add(Me.GetProbabilityData(CurrentSound, NextSound,
                                          CurrentSyllabicPosition, CurrentSyllablePart))

                            'Adding the phonemes in the used-phonemes array
                            'If the current sound is a stressed vowel, getting a string that respressents stress/tone to add to the vowel
                            Dim CurrentStressMarker As String = ""
                            If CurrentSyllabicPosition = SyllabicPositions.Stressed And SwedishVowels_IPA.Contains(CurrentSound) Then
                                If CurrentWord.MainStressSyllableIndex = syllable + 1 Then
                                    'Getting the tone accent
                                    If CurrentWord.Tone = 1 Then
                                        CurrentStressMarker = IpaMainStress
                                    Else
                                        CurrentStressMarker = IpaMainSwedishAccent2
                                    End If
                                End If

                                If ReduceSecondaryStress = False Then
                                    If CurrentWord.SecondaryStressSyllableIndex = syllable + 1 Then
                                        CurrentStressMarker = IpaSecondaryStress
                                    End If
                                End If
                            End If
                            'Adding the current state
                            CurrentWord.PP_Phonemes.Add(CurrentStressMarker & CurrentSound)
                        End If
                        phonemeIndex += 1
                    Next
                Next

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try


            If TransitionDataDirection = PhonotacticTransitionDataDirections.RetrievePhonotacticProbabilities Then
                'Calculating average and minimum PP
                If Not CurrentWord.PP.Count = 0 Then
                    CurrentWord.SSPP_Average = CurrentWord.PP.Average
                    CurrentWord.SSPP_Min = CurrentWord.PP.Min

                Else
                    CurrentWord.SSPP_Average = 0
                    CurrentWord.SSPP_Min = 0

                    'Reporting lacking probability data to log file
                    SendInfoToLog(String.Join(" ", CurrentWord.BuildExtendedIpaArray), "WordsLackingPhonotacticProbabilityData")

                End If
            End If

            'Adding the final word end marker to PhonoTacticProbabilityPhonemes (which is not added above, since only current phonemes are added)
            If TransitionDataDirection = PhonotacticTransitionDataDirections.RetrievePhonotacticProbabilities Then CurrentWord.PP_Phonemes.Add(WordEndMarker)

        Next

        If TransitionDataDirection = PhonotacticTransitionDataDirections.CreateTransitionMatrix Then
            'Calculates transitional probabilities after all frequency data is set
            CalculateTransitionalProbabilities()

            'Sorting transitional probabilities data
            SortProbabilityData()

            If ExportDataToFile = True Then

                'Exporting transitional probabilities to txt file, without rounding
                ExportTransitionalProbabilityData(OutputFolder, OutputFileName, False,,, True)
                ExportTransitionalProbabilityData(OutputFolder, OutputFileName & "_FullLines", True, ,, True)

                'Exporting transitional probabilities to txt file, with rounded values
                ExportTransitionalProbabilityData(OutputFolder, OutputFileName & "_Rounded", False,,, False)
                ExportTransitionalProbabilityData(OutputFolder, OutputFileName & "_Rounded" & "_FullLines", True,,, False)

            End If
        End If

        'Closing the progress display
        myProgress.Close()

    End Sub

    Private Sub CalculateTransitionalProbabilities(Optional ByVal AcceptedRoundingErrorSize As Double = 0.0001)


            Dim DoFilterring As Boolean = True
            If DoFilterring = True Then

                Dim PhonemeThresholdValue As Decimal = 2
                Dim TransitionPhonemeThresholdValue As Integer = 2

                Dim PhonemesRemoved As Integer = 0
                Dim TransitionPhonemesRemoved As Integer = 0

                SendInfoToLog("Filterring out unusual phoneme transition.")

                'Filterring out very uncommon phoneme transitions
                For Each SyllabicPosition In TransitionalProbabilities
                    For Each SyllablePart In SyllabicPosition.Value
                        For Each SourcePhoneme In SyllablePart.Value

                            'Collecting phonemes to remove (which have a occurence count below PhonemeThresholdValue)
                            Dim TransitionPhonemesToRemove As New List(Of String)
                            For Each TransitionPhoneme In SourcePhoneme.Value
                                If TransitionPhoneme.Value.OccurenceCount < TransitionPhonemeThresholdValue Then TransitionPhonemesToRemove.Add(TransitionPhoneme.Key)
                            Next

                            'Removing transition phonemes
                            For Each CurrentPhoneme In TransitionPhonemesToRemove
                                SourcePhoneme.Value.Remove(CurrentPhoneme)
                                TransitionPhonemesRemoved += 1
                            Next
                        Next
                    Next
                Next

                'Filterring out very uncommon phoneme transitions
                For Each SyllabicPosition In TransitionalProbabilities
                    For Each SyllablePart In SyllabicPosition.Value


                        'Collecting phonemes to remove (which have a occurence count below PhonemeThresholdValue)
                        Dim PhonemesToRemove As New List(Of String)
                        For Each SourcePhoneme In SyllablePart.Value
                            If SourcePhoneme.Value.OccurenceCount < PhonemeThresholdValue Then PhonemesToRemove.Add(SourcePhoneme.Key)

                            'Checking if there are any transition phonemes left after removal of transition phonemes above. If not, the phoneme is removed.
                            If SourcePhoneme.Value.Count = 0 Then If Not PhonemesToRemove.Contains(SourcePhoneme.Key) Then PhonemesToRemove.Add(SourcePhoneme.Key)
                        Next

                        'Removing phonemes (and counting the transition phonemes also removed at the same time)
                        For Each CurrentPhoneme In PhonemesToRemove
                            TransitionPhonemesRemoved += SyllablePart.Value(CurrentPhoneme).OccurenceCount

                            SyllablePart.Value.Remove(CurrentPhoneme)
                            PhonemesRemoved += 1
                        Next
                    Next
                Next


                SendInfoToLog("Finished filterring out unusual phoneme transition. Results: " & PhonemesRemoved & " target phonemes and " & TransitionPhonemesRemoved & " transition phonemes were removed.")

            End If


            SendInfoToLog("Calculating transitional probabilities.")

            'Summing phoneme frequency data
            For Each SyllabicPosition In TransitionalProbabilities
                For Each SyllablePart In SyllabicPosition.Value
                    For Each SourcePhoneme In SyllablePart.Value
                        SourcePhoneme.Value.FrequencyDataSum = 0
                        For Each TargetPhoneme In SourcePhoneme.Value


                            Select Case FrequencyUnit
                                Case FrequencyUnits.WordCount
                                        'Logs were taken in the first step, now only summing
                                        'Nothing need to be done

                                Case FrequencyUnits.PhonemeCount
                                    'Raw phoneme frequencies were collected in the first step, now log 10 of these are summed
                                    TargetPhoneme.Value.FrequencyData = Math.Log10(TargetPhoneme.Value.FrequencyData + 1) ' +1 is used to avoid that a raw frequency of 1 becomes a zero probability: log10(1) = 0. This means that the phoneme count is very slightly elevated, by 1

                            End Select


                            'Summing the data
                            SourcePhoneme.Value.FrequencyDataSum += TargetPhoneme.Value.FrequencyData

                        Next
                    Next
                Next
            Next


            'Calculating transitional probabilies
            Dim ZeroProbabilityTransitionCount As Integer = 0
            For Each SyllabicPosition In TransitionalProbabilities
                For Each SyllablePart In SyllabicPosition.Value
                    For Each SourcePhoneme In SyllablePart.Value
                        For Each TargetPhoneme In SourcePhoneme.Value

                            If SourcePhoneme.Value.FrequencyDataSum = 0 Then
                                TargetPhoneme.Value.TransitionalProbability = 0 'TransitionalProbability is set to zero if the frequency in the word list is 0 (Log10(RawFrequency of 0+1))
                                ZeroProbabilityTransitionCount += 1
                            Else
                                TargetPhoneme.Value.TransitionalProbability = TargetPhoneme.Value.FrequencyData / SourcePhoneme.Value.FrequencyDataSum

                            End If

                        Next
                    Next
                Next
            Next
            SendInfoToLog("Finished calculating transitional probabilities. " & ZeroProbabilityTransitionCount & " probabilities of 0 detected.")


            'Checking that the transitional probabilies for all transitional possibilities for each source phoneme adds up to 1

            SendInfoToLog("Checking for errors in calculations of transitional probabilities.")

            Dim ErrorsInTransitionalProbabilitiyCalculations As Integer = 0
            For Each SyllabicPosition In TransitionalProbabilities
                For Each SyllablePart In SyllabicPosition.Value
                    For Each SourcePhoneme In SyllablePart.Value
                        SourcePhoneme.Value.TransitionProbabilitySum = 0
                        For Each TargetPhoneme In SourcePhoneme.Value
                            Try
                                SourcePhoneme.Value.TransitionProbabilitySum += TargetPhoneme.Value.TransitionalProbability
                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                        Next
                        If Not SourcePhoneme.Value.TransitionProbabilitySum > (1 - AcceptedRoundingErrorSize) And
                                SourcePhoneme.Value.TransitionProbabilitySum < (1 + AcceptedRoundingErrorSize) Then
                            ErrorsInTransitionalProbabilitiyCalculations += 1
                        End If
                    Next
                Next
            Next

            SendInfoToLog("   Checked for errors in calculations of transitional probabilities completed successfully. Results: " & ErrorsInTransitionalProbabilitiyCalculations & " probability values do not add up to 1 +/-" & AcceptedRoundingErrorSize)

            If Do_z_ScoreTransformation = True Then
                SendInfoToLog("   Doing z-score transformation.")

                'Determining the highest phonotactic probability for each source phoneme
                For Each SyllabicPosition In TransitionalProbabilities
                    For Each SyllablePart In SyllabicPosition.Value

                        'Collecting probability data
                        Dim TransformationData As New List(Of Double)
                        For Each SourcePhoneme In SyllablePart.Value
                            For Each TargetPhoneme In SourcePhoneme.Value
                                TransformationData.Add(TargetPhoneme.Value.TransitionalProbability)
                            Next
                        Next

                        'Transforming the data
                        MathMethods.Standardization(TransformationData)

                        'Putting the transformed data back
                        Dim CurrentDataIndex As Integer = 0
                        For Each SourcePhoneme In SyllablePart.Value
                            For Each TargetPhoneme In SourcePhoneme.Value
                                TargetPhoneme.Value.TransitionalProbability = TransformationData(CurrentDataIndex)
                                CurrentDataIndex += 1
                            Next
                        Next
                    Next
                Next
            End If


            If CalculationType = PhonoTacticCalculationTypes.PhonotacticPredictability Then

                SendInfoToLog("Calculating transitional predictabilities...")

                'Determining the highest phonotactic probability for each source phoneme
                For Each item_level_A In TransitionalProbabilities
                    For Each item_level_B In item_level_A.Value
                        For Each SourcePhoneme In item_level_B.Value
                            SourcePhoneme.Value.HighestTransitionalProbability = 0
                            For Each TargetPhoneme In SourcePhoneme.Value
                                If TargetPhoneme.Value.TransitionalProbability > SourcePhoneme.Value.HighestTransitionalProbability Then
                                    SourcePhoneme.Value.HighestTransitionalProbability = TargetPhoneme.Value.TransitionalProbability
                                End If
                            Next
                        Next
                    Next
                Next

                'Calculating transitional predictabilities
                For Each item_level_A In TransitionalProbabilities
                    For Each item_level_B In item_level_A.Value
                        For Each SourcePhoneme In item_level_B.Value
                            For Each TargetPhoneme In SourcePhoneme.Value
                                If Not SourcePhoneme.Value.HighestTransitionalProbability = 0 Then
                                    TargetPhoneme.Value.TransitionPredictability = TargetPhoneme.Value.TransitionalProbability / SourcePhoneme.Value.HighestTransitionalProbability
                                Else
                                    TargetPhoneme.Value.TransitionPredictability = 0
                                End If
                            Next
                        Next
                    Next
                Next

            End If

        End Sub


        Public Sub ExportTransitionalProbabilityData(Optional ByRef saveDirectory As String = "", Optional ByRef saveFileName As String = "SyllableBasedPhonemeTransitionalProbabilities",
                                                Optional ByVal OutputFullLines As Boolean = False, Optional BoxTitle As String = "Choose location to store the transitional probability export file...",
                                                         Optional SkipZeroProbabilityTransitions As Boolean = True, Optional SkipRounding As Boolean = False)

            Try

                SendInfoToLog("Attempts to save transitional probability data to .txt file.")

                'Choosing file location
                Dim filepath As String = ""
                'Ask the user for file path if not incomplete file path is given
                If saveDirectory = "" Or saveFileName = "" Then
                    filepath = GetSaveFilePath(saveDirectory, saveFileName, {"txt"}, BoxTitle)
                Else
                    filepath = Path.Combine(saveDirectory, saveFileName & ".txt")
                    If Not Directory.Exists(Path.GetDirectoryName(filepath)) Then Directory.CreateDirectory(Path.GetDirectoryName(filepath))
                End If

                'Save it to file
                Dim writer As New StreamWriter(filepath, False, Text.Encoding.UTF8)

                If OutputFullLines = False Then


                    For Each item_level_A In TransitionalProbabilities
                        If item_level_A.Value.Count = 0 Then Continue For
                        writer.WriteLine(item_level_A.Key.ToString)

                        For Each item_level_B In item_level_A.Value
                            If item_level_B.Value.Count = 0 Then Continue For
                            writer.WriteLine(item_level_B.Key.ToString)

                            Select Case CalculationType
                                Case PhonoTacticCalculationTypes.PhonotacticProbability


                                    'Witing heading
                                    writer.WriteLine(vbTab & "Phoneme" & vbTab & "TransitionTo" & vbTab & "SummedFrequencyData" & vbTab & "TransitionalPropbability" & vbTab & "Occurences")

                                    For Each item_level_C In item_level_B.Value
                                        If item_level_C.Value.Count = 0 Then Continue For
                                        writer.WriteLine(vbTab & item_level_C.Key.ToString & vbTab & vbTab & Rounding(item_level_C.Value.FrequencyDataSum,, 4, SkipRounding) & vbTab &
                                                             Rounding(item_level_C.Value.TransitionProbabilitySum,, 6, SkipRounding) & vbTab & item_level_C.Value.OccurenceCount)

                                        For Each TransitionPhoneme In item_level_C.Value
                                            If SkipZeroProbabilityTransitions = True And TransitionPhoneme.Value.TransitionalProbability = 0 Then Continue For

                                            writer.WriteLine(vbTab & vbTab & TransitionPhoneme.Key.ToString & vbTab & Rounding(TransitionPhoneme.Value.FrequencyData, , 4, SkipRounding) & vbTab &
                                                                 Rounding(TransitionPhoneme.Value.TransitionalProbability, , 4, SkipRounding) & vbTab & TransitionPhoneme.Value.OccurenceCount)

                                        Next
                                        writer.WriteLine() 'Inserting an empty line before the next section
                                    Next
                                    writer.WriteLine() 'Inserting an empty line before the next section

                                Case PhonoTacticCalculationTypes.PhonotacticPredictability

                                    'Witing heading
                                    writer.WriteLine(vbTab & "Phoneme" & vbTab & "TransitionTo" & vbTab & "SummedFrequencyData" & vbTab & "TransitionalPropbability" & vbTab & "TransitionalPredictability" & vbTab & "Occurences")

                                    For Each item_level_C In item_level_B.Value
                                        If item_level_C.Value.Count = 0 Then Continue For
                                        writer.WriteLine(vbTab & item_level_C.Key.ToString & vbTab & vbTab & Rounding(item_level_C.Value.FrequencyDataSum,, 4, SkipRounding) & vbTab &
                                                             Rounding(item_level_C.Value.TransitionProbabilitySum,, 4, SkipRounding) & vbTab & item_level_C.Value.OccurenceCount)

                                        For Each TransitionPhoneme In item_level_C.Value
                                            If SkipZeroProbabilityTransitions = True And TransitionPhoneme.Value.TransitionPredictability = 0 Then Continue For

                                            writer.WriteLine(vbTab & vbTab & TransitionPhoneme.Key.ToString & vbTab & Rounding(TransitionPhoneme.Value.FrequencyData,, 4, SkipRounding) & vbTab &
                                                                 Rounding(TransitionPhoneme.Value.TransitionalProbability, , 4, SkipRounding) & vbTab & Rounding(TransitionPhoneme.Value.TransitionPredictability, , 4, SkipRounding) & vbTab &
                                                                  TransitionPhoneme.Value.OccurenceCount)

                                        Next
                                        writer.WriteLine() 'Inserting an empty line before the next section
                                    Next
                                    writer.WriteLine() 'Inserting an empty line before the next section


                            End Select
                        Next
                    Next

                Else

                    'Writing heading
                    Select Case CalculationType
                        Case PhonoTacticCalculationTypes.PhonotacticProbability
                            writer.WriteLine("Stress" & vbTab & "Position" & vbTab & "Phoneme" & vbTab & "TransitionTo" & vbTab & "TransitionalPropbability" & vbTab & "JointPosition" & vbTab & "JointTransition" & vbTab & "PhonemeOccurences" & vbTab & "TransitionToOccurences")
                        Case PhonoTacticCalculationTypes.PhonotacticPredictability
                            writer.WriteLine("Stress" & vbTab & "Position" & vbTab & "Phoneme" & vbTab & "TransitionTo" & vbTab & "TransitionalPropbability" & vbTab & "TransitionalPredictability" & vbTab & "JointPosition" & vbTab & "JointTransition" & vbTab & "PhonemeOccurences" & vbTab & "TransitionToOccurences")
                    End Select

                    For Each item_level_A In TransitionalProbabilities
                        If item_level_A.Value.Count = 0 Then Continue For
                        For Each item_level_B In item_level_A.Value
                            If item_level_B.Value.Count = 0 Then Continue For
                            Select Case CalculationType
                                Case PhonoTacticCalculationTypes.PhonotacticProbability
                                    For Each item_level_C In item_level_B.Value
                                        If item_level_C.Value.Count = 0 Then Continue For
                                        For Each TransitionPhoneme In item_level_C.Value
                                            If SkipZeroProbabilityTransitions = True And TransitionPhoneme.Value.TransitionalProbability = 0 Then Continue For
                                            writer.WriteLine(item_level_A.Key.ToString & vbTab & item_level_B.Key.ToString & vbTab &
                                                                 item_level_C.Key.ToString & vbTab & TransitionPhoneme.Key.ToString & vbTab &
                                                                 Rounding(TransitionPhoneme.Value.TransitionalProbability, , 4, SkipRounding) & vbTab & vbTab &
                                                                 item_level_A.Key.ToString & "-" & item_level_B.Key.ToString & vbTab &
                                                                 item_level_C.Key.ToString & "-" & TransitionPhoneme.Key.ToString & vbTab & item_level_C.Value.OccurenceCount & vbTab & TransitionPhoneme.Value.OccurenceCount)
                                        Next
                                    Next
                                Case PhonoTacticCalculationTypes.PhonotacticPredictability

                                    For Each item_level_C In item_level_B.Value
                                        If item_level_C.Value.Count = 0 Then Continue For
                                        For Each TransitionPhoneme In item_level_C.Value
                                            If SkipZeroProbabilityTransitions = True And TransitionPhoneme.Value.TransitionPredictability = 0 Then Continue For
                                            writer.WriteLine(item_level_A.Key.ToString & vbTab & item_level_B.Key.ToString & vbTab &
                                                                 item_level_C.Key.ToString & vbTab & TransitionPhoneme.Key.ToString & vbTab &
                                                                 Rounding(TransitionPhoneme.Value.TransitionalProbability, , 4, SkipRounding) & vbTab &
                                                                  Rounding(TransitionPhoneme.Value.TransitionPredictability, , 4, SkipRounding) & vbTab &
                                                                 item_level_A.Key.ToString & "-" & item_level_B.Key.ToString & vbTab &
                                                                 item_level_C.Key.ToString & "-" & TransitionPhoneme.Key.ToString & vbTab & item_level_C.Value.OccurenceCount & vbTab & TransitionPhoneme.Value.OccurenceCount)
                                        Next
                                    Next
                            End Select
                        Next
                    Next

                End If


                writer.Close()

                SendInfoToLog("   Transitional probability data were successfully saved to .txt file: " & filepath)

            Catch ex As Exception

            End Try


        End Sub



    ''' <summary>
    ''' Loads the probability data from a "Full-Lines" type exported SSPP type phonotactic probability data file.
    ''' </summary>
    ''' <param name="FilePath"></param>
    Public Sub LoadProbabilityDataFromFile(Optional ByRef FilePath As String = "")

        'This should load probability data from file, to enable phonotactic probability calculations of words without running the word list analysis part
        TransitionalProbabilities = New PhonoTacticProbabilities(PhonoTacticGroupingType)

        Dim InputLines() As String

        If FilePath = "" Then
            Dim dataString As String = My.Resources.SSPP_Matrix_FullLines
            dataString = dataString.Replace(vbCrLf, vbLf)
            InputLines = dataString.Split(vbLf)
        Else
            InputLines = System.IO.File.ReadAllLines(FilePath, Text.Encoding.UTF8)
        End If


        For LineIndex = 1 To InputLines.Length - 1 'Skipping heading line

            If InputLines(LineIndex).Trim = "" Then Continue For

            'Reading all data in the current line
            Dim LineSplit() As String = InputLines(LineIndex).Split(vbTab)

            'Check line split length
            If LineSplit.Length < 10 Then MsgBox("Error reading line: " & InputLines(LineIndex))

            Dim SyllabicPosition As SyllabicPositions = [Enum].Parse(GetType(SyllabicPositions), LineSplit(0))
            Dim SyllablePart As SyllableParts = [Enum].Parse(GetType(SyllableParts), LineSplit(1))
            Dim SourcePhonemeString As String = LineSplit(2)
            Dim TargetPhonemeString As String = LineSplit(3)
            Dim TransitionalProbability As Double = LineSplit(4)
            Dim TransitionalPredictability As Double = LineSplit(5)
            Dim PhonemeOccurences As Double = LineSplit(8)
            Dim TransitionToOccurences As Double = LineSplit(9)

            'Checking that the appropriate locvation in the structure exists, or adds it otherwise
            If Not TransitionalProbabilities.ContainsKey(SyllabicPosition) Then TransitionalProbabilities.Add(SyllabicPosition, New PartOfSyllable)
            If Not TransitionalProbabilities(SyllabicPosition)(SyllablePart).ContainsKey(SourcePhonemeString) Then TransitionalProbabilities(SyllabicPosition)(SyllablePart).Add(SourcePhonemeString, New TransitionPhonemes(CalculationType))
            If Not TransitionalProbabilities(SyllabicPosition)(SyllablePart)(SourcePhonemeString).ContainsKey(TargetPhonemeString) Then TransitionalProbabilities(SyllabicPosition)(SyllablePart)(SourcePhonemeString).Add(TargetPhonemeString, New TransitionPhonemes.ProbabilityData)

            'Adding the probabililty data
            TransitionalProbabilities(SyllabicPosition)(SyllablePart)(SourcePhonemeString).OccurenceCount = PhonemeOccurences 'This is read for every transition phoneme, even if it is only necessary to read it once (a bit lacy programming here)
            TransitionalProbabilities(SyllabicPosition)(SyllablePart)(SourcePhonemeString)(TargetPhonemeString).OccurenceCount = TransitionToOccurences
            TransitionalProbabilities(SyllabicPosition)(SyllablePart)(SourcePhonemeString)(TargetPhonemeString).TransitionalProbability = TransitionalProbability
            TransitionalProbabilities(SyllabicPosition)(SyllablePart)(SourcePhonemeString)(TargetPhonemeString).TransitionPredictability = TransitionalPredictability

        Next

    End Sub

    ''' <summary>
    ''' Sorts the output data in descending order determined by the values of TransitionalProbability or TransitionPredictability depending on the selected calculation type.
    ''' </summary>
    Private Sub SortProbabilityData()

            'Sorting transition phonemes according to their transitional probability (source phonemes are already sorted as they are in a SortedList)
            Dim MySortOrder As SortOrder = SortOrder.Descending
            For Each item_level_A In TransitionalProbabilities
                For Each item_level_B In item_level_A.Value
                    For Each item_level_C In item_level_B.Value
                        item_level_C.Value.SortOutputData()
                    Next
                Next
            Next

        End Sub


        Public Enum SyllableParts
            Onset
            Nucleus
            Coda
        End Enum

        Public Enum SyllabicPositions
            Stressed
            Ante_Stress
            Inter_Stress
            Post_Stress
        End Enum

        Private Class PhonoTacticProbabilities
            Inherits Dictionary(Of SyllabicPositions, PartOfSyllable)

            Public Sub New(ByVal PhonoTacticGroupingType As PhonoTacticGroupingTypes)

                Select Case PhonoTacticGroupingType

                    Case PhonoTacticGroupingTypes.Standard
                        Add(SyllabicPositions.Ante_Stress, New PartOfSyllable)
                        Add(SyllabicPositions.Stressed, New PartOfSyllable)
                        Add(SyllabicPositions.Inter_Stress, New PartOfSyllable)
                        Add(SyllabicPositions.Post_Stress, New PartOfSyllable)
                    Case Else
                        Throw New NotImplementedException
                End Select

            End Sub

        End Class

        Private Class PartOfSyllable
            Inherits Dictionary(Of SyllableParts, SourcePhonemes)

            Public Sub New()
                Add(SyllableParts.Onset, New SourcePhonemes)
                Add(SyllableParts.Nucleus, New SourcePhonemes)
                Add(SyllableParts.Coda, New SourcePhonemes)
            End Sub

        End Class

        Private Class SourcePhonemes
            Inherits SortedList(Of String, TransitionPhonemes)
        End Class

        Public Class TransitionPhonemes
            Inherits Dictionary(Of String, ProbabilityData)

            Public Property FrequencyDataSum As Double
            Public Property TransitionProbabilitySum As Double
            Public Property HighestTransitionalProbability As Double
            Public Property OccurenceCount As Double

            ReadOnly Property Calculationtype As PhonoTactics.PhonoTacticCalculationTypes

            Public Sub New(ByVal SetCalculationtype As PhonoTactics.PhonoTacticCalculationTypes)
                Calculationtype = SetCalculationtype
            End Sub

            Public Class ProbabilityData

                Public Property FrequencyData As Double
                Public Property TransitionalProbability As Double
                Public Property TransitionPredictability As Double
                Public Property OccurenceCount As Double

            End Class

            Public Sub SortOutputData()

                'Putting all data to sort in a list
                Dim newSortList As New List(Of SortData)
                For Each CurrentItem In Me
                    Dim newSortData As New SortData(Calculationtype)
                    newSortData.MyKey = CurrentItem.Key
                    newSortData.MyValue = CurrentItem.Value
                    newSortList.Add(newSortData)
                Next

                'Sorting the list
                newSortList.Sort()

                'Reversing the order
                newSortList.Reverse()

                'Clearing Me
                Me.Clear()

                'Putting the data back into Me
                For Each CurrentItem In newSortList
                    Me.Add(CurrentItem.MyKey, CurrentItem.MyValue)
                Next

            End Sub

            Private Class SortData
                Implements IComparable

                Property MyKey As String
                Property MyValue As ProbabilityData
                ReadOnly Property Calculationtype As PhonoTactics.PhonoTacticCalculationTypes

                Public Sub New(ByVal SetCalculationtype As PhonoTactics.PhonoTacticCalculationTypes)
                    Calculationtype = SetCalculationtype
                End Sub

                Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo

                    If Not TypeOf (obj) Is SortData Then
                        Throw New ArgumentException()
                    Else

                        Dim tempData As SortData = DirectCast(obj, SortData)

                        If Not Me.Calculationtype = tempData.Calculationtype Then
                            Throw New ArgumentException()
                        Else

                            Select Case Calculationtype
                                Case PhonoTacticCalculationTypes.PhonotacticProbability

                                    If Me.MyValue.TransitionalProbability < tempData.MyValue.TransitionalProbability Then
                                        Return -1
                                    ElseIf Me.MyValue.TransitionalProbability = tempData.MyValue.TransitionalProbability Then
                                        Return 0
                                    Else
                                        Return 1
                                    End If

                                Case PhonoTacticCalculationTypes.PhonotacticPredictability

                                    If Me.MyValue.TransitionPredictability < tempData.MyValue.TransitionPredictability Then
                                        Return -1
                                    ElseIf Me.MyValue.TransitionPredictability = tempData.MyValue.TransitionPredictability Then
                                        Return 0
                                    Else
                                        Return 1
                                    End If
                                Case Else

                            End Select
                            Throw New ArgumentException()
                        End If
                    End If
                End Function
            End Class



        End Class


    End Class

    ''' <summary>
    ''' This class calculates the types of phonotactic probability described by Vitevitch and Luce, and Storkel et al. based on the Swedish language. (Cf. Witte and Köbler 2019)
    ''' </summary>
    Public Class Positional_PhonoTactics

        Public ReadOnly Property WordStartMarker As String
        Public ReadOnly Property WordEndMarker As String
        Public ReadOnly Property UseAlternateSyllabification As Boolean
        Public ReadOnly Property ReduceSecondaryStress As Boolean
        Public ReadOnly Property UseSurfacePhones As Boolean
        Public ReadOnly Property PhonemeCombinationLength As PhonemeCombinationLengths
        Public ReadOnly Property FrequencyUnit As FrequencyUnits
        Public ReadOnly Property ExcludeForeignWordsFromMatrixCreation As Boolean
        Public Property Do_z_ScoreTransformation As Boolean = False


        Private Property TransitionalProbabilities As New PhonoTacticProbabilities

        Public Enum PhonemeCombinationLengths
            MonoGramCalculation
            BiGramCalculation
        End Enum


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="SetWordStartMarker"></param>
        ''' <param name="SetWordEndMarker"></param>
        ''' <param name="SetUseAlternateSyllabification"></param>
        ''' <param name="SetUseSurfacePhones">If set to true, the quality of all length-reduced underlyingly long vowel phonemes will be reduced to that of their short allophones.</param>
        ''' <param name="SetReduceSecondaryStress">Only has effect if SetUseSurfacePhonesis = True. If set to true, secondarily stressed syllables will be treated as non stressed syllables, with their phonetic length ignored, and vowel phonemes reduced to their short allophones.</param>
        Public Sub New(Optional ByRef SetWordStartMarker As String = "*", Optional ByRef SetWordEndMarker As String = "_",
                           Optional ByRef SetUseAlternateSyllabification As Boolean = False,
                           Optional ByRef SetUseSurfacePhones As Boolean = False,
                           Optional ByRef SetReduceSecondaryStress As Boolean = False,
                           Optional ByRef SetPhonemeCombinationLength As PhonemeCombinationLengths = PhonemeCombinationLengths.BiGramCalculation,
                           Optional ByVal SetFrequencyUnit As FrequencyUnits = FrequencyUnits.WordCount,
                           Optional ByVal SetExcludeForeignWordsFromMatrixCreation As Boolean = True)

            SendInfoToLog("Creating an instance: " & Me.ToString & ", SetPhonemeCombinationLength: " & SetPhonemeCombinationLength.ToString)

            UseAlternateSyllabification = SetUseAlternateSyllabification
            ReduceSecondaryStress = SetReduceSecondaryStress
            UseSurfacePhones = SetUseSurfacePhones
            WordStartMarker = SetWordStartMarker
            WordEndMarker = SetWordEndMarker
            PhonemeCombinationLength = SetPhonemeCombinationLength
            FrequencyUnit = SetFrequencyUnit
            ExcludeForeignWordsFromMatrixCreation = SetExcludeForeignWordsFromMatrixCreation

            If UseSurfacePhones = False Then ReduceSecondaryStress = False

        End Sub

        Public Enum FrequencyUnits
            WordCount
            PhonemeCount
        End Enum

        ''' <summary>
        ''' Adds data for a phoneme transition.
        ''' </summary>
        ''' <param name="StartIndex">The position starting from word start</param>
        ''' <param name="CurrentPhoneme">The current phoneme.</param>
        ''' <param name="WordFrequency">The word frequency data for the current word.</param>
        Private Sub AddProbabilityData(ByVal StartIndex As Integer, ByVal CurrentPhoneme As String, ByVal WordFrequency As Double, ByVal DoNotAddDataWithZeroFrequency As Boolean)

            'Not adding if DoNotAddDataWithZeroFrequency = true and the Frequency value is zero
            If DoNotAddDataWithZeroFrequency = True And WordFrequency = 0 Then Exit Sub

            'Adding Current startindex key if it doesn't exists
            If Not TransitionalProbabilities.ContainsKey(StartIndex) Then
                TransitionalProbabilities.Add(StartIndex, New PhonemeData)
            End If

            'Adding Current phoneme/combination if it doesn't exists, and accumulates its frequency value
            If Not TransitionalProbabilities(StartIndex).ContainsKey(CurrentPhoneme) Then
                Dim NewProbabilityData As New PhonemeData.ProbabilityData
                NewProbabilityData.FrequencyData = WordFrequency
                TransitionalProbabilities(StartIndex).Add(CurrentPhoneme, NewProbabilityData)
            Else
                TransitionalProbabilities(StartIndex)(CurrentPhoneme).FrequencyData += WordFrequency
            End If

            'Counting occurences
            TransitionalProbabilities(StartIndex)(CurrentPhoneme).OccurenceCount += 1

        End Sub

        ''' <summary>
        ''' Adds data for a phoneme transition.
        ''' </summary>
        ''' <param name="StartIndex">The position starting from word start</param>
        ''' <param name="CurrentPhoneme">The current phoneme.</param>
        Private Function GetProbabilityData(ByVal StartIndex As Integer, ByVal CurrentPhoneme As String) As Double

            'Checking that the current data exists in the probability collection, returns 0 otherwise
            If Not TransitionalProbabilities.ContainsKey(StartIndex) Then
                Return 0
            End If
            If Not TransitionalProbabilities(StartIndex).ContainsKey(CurrentPhoneme) Then
                Return 0
            End If

            If Do_z_ScoreTransformation = True Then
                Return TransitionalProbabilities(StartIndex)(CurrentPhoneme).Z_Transformed_TransitionalProbability
            Else
                Return TransitionalProbabilities(StartIndex)(CurrentPhoneme).TransitionalProbability
            End If


        End Function


        ''' <summary>
        ''' This sub either creates a phonotactic transition matrix from an input word group or retrieves the phonotactic probability data for each word using a pre-existing transition matrix.
        ''' The type of use is selected using the TransitionDataDirection argument.
        ''' </summary>
        ''' <param name="TransitionDataDirection"></param>
        ''' <param name="InputWordGroup"></param>
        ''' <param name="IncludeWordStartMarker"></param>
        ''' <param name="IncludeWordEndMarker"></param>
        ''' <param name="IncludeSyllableBoundaries"></param>
        ''' <param name="OutputFolder"></param>
        ''' <param name="OutputFileName"></param>
        ''' <param name="SetDo_z_ScoreTransformation">Only used on transition direction RetrievePhonotacticProbabilities. If set to False, untransformed Vitevitch type data will be retrieved. If set to true, Storkel type z-transformed data will will retrieved.</param>
        Public Sub TransitionData(ByVal TransitionDataDirection As PhonotacticTransitionDataDirections, ByVal InputWordGroup As WordGroup,
                Optional ByVal IncludeWordStartMarker As Boolean = False,
                Optional ByVal IncludeWordEndMarker As Boolean = False,
                                      Optional ByVal IncludeSyllableBoundaries As Boolean = False,
                                      Optional ByRef OutputFolder As String = "", Optional ByRef OutputFileName As String = "",
                                      Optional ByVal SetDo_z_ScoreTransformation As Boolean = True)

            SendInfoToLog("Initializing method: " & System.Reflection.MethodInfo.GetCurrentMethod.Name & " TransitionDataDirection:" & TransitionDataDirection.ToString)

            'Setting Z-score active or not
            Do_z_ScoreTransformation = SetDo_z_ScoreTransformation

            'Going through each phoneme in each word in the input word group and stores its frequency data

            'Starting a progress window
            Dim myProgress As New ProgressDisplay
            If TransitionDataDirection = PhonotacticTransitionDataDirections.RetrievePhonotacticProbabilities Then
                myProgress.Initialize(InputWordGroup.MemberWords.Count - 1, 0, "Calculating phonotactic probability...", 100)
            Else
                myProgress.Initialize(InputWordGroup.MemberWords.Count - 1, 0, "Collecting phoneme transition data...", 100)
            End If

            myProgress.Show()
            Dim ProcessedWordsCount As Integer = 0

            'Creating an array of phonemes as calculation base
            'If TransitionDataDirection = PhonotacticTransitionDataDirections.CreateTransitionMatrix Then '(As with the added possibility of loading probability data, TransitionData may be run without first PhonotacticTransitionDataDirections.CreateTransitionMatrix, where by this step is needed also on PhonotacticTransitionDataDirections.RetrievePhonotacticProbabilities.)
            For Each CurrentWord In InputWordGroup.MemberWords
                CurrentWord.Phonemes = CurrentWord.BuildExtendedIpaArray(, True, ReduceSecondaryStress, UseSurfacePhones, UseAlternateSyllabification, IncludeSyllableBoundaries, False)
                If IncludeWordStartMarker = True Then CurrentWord.Phonemes.Insert(0, WordStartMarker)
                If IncludeWordEndMarker = True Then CurrentWord.Phonemes.Add(WordEndMarker)
            Next
            'End If



            For Each CurrentWord In InputWordGroup.MemberWords

                'Updating progress
                myProgress.UpdateProgress(ProcessedWordsCount)
                ProcessedWordsCount += 1

                'Creating output arrays
                Dim PP_Data As List(Of Double) = Nothing 'Creating a list to hold PP data
                Dim PP_Phonemes As List(Of String) = Nothing 'Creating a list to hold the phoneme equivalents to the added probability data

                If TransitionDataDirection = PhonotacticTransitionDataDirections.RetrievePhonotacticProbabilities Then
                    'Initiating the lists on RetrievePhonotacticProbabilities
                    PP_Data = New List(Of Double)
                    PP_Phonemes = New List(Of String)
                End If

                'Does not add probabilities based on foreign words
                If TransitionDataDirection = PhonotacticTransitionDataDirections.CreateTransitionMatrix Then
                    If ExcludeForeignWordsFromMatrixCreation = True And CurrentWord.ForeignWord = True Then Continue For
                End If

                For phonemeIndex = 0 To CurrentWord.Phonemes.Count - 1

                    'Gets the current sound
                    Dim CurrentSound As String = CurrentWord.Phonemes(phonemeIndex)

                    If PhonemeCombinationLength = PhonemeCombinationLengths.BiGramCalculation Then

                        'Skipping the last step, since this does not contain any biphone
                        If phonemeIndex = CurrentWord.Phonemes.Count - 1 Then Continue For

                        'Setting creating the biphone string
                        CurrentSound &= " " & CurrentWord.Phonemes(phonemeIndex + 1)

                    Else
                        'Skips the first index if this i a word start marker
                        If IncludeWordStartMarker = True And phonemeIndex = 0 Then
                            Continue For
                        End If

                        'Skips the last index if this is a word end marker
                        If IncludeWordEndMarker = True And phonemeIndex = CurrentWord.Phonemes.Count - 1 Then
                            Continue For
                        End If

                    End If

                    'Stores or sets probability data
                    If TransitionDataDirection = PhonotacticTransitionDataDirections.CreateTransitionMatrix Then
                        'Adding frequency info to the probability data
                        Dim FrequencyValue As Double = 0

                        Select Case FrequencyUnit
                            Case FrequencyUnits.WordCount
                                'Logs are taken straigth away (Log10(token count) => word frequency weighted phoneme count)
                                FrequencyValue = Math.Log10(CurrentWord.RawWordTypeFrequency + 2) ' +2 is used to avoid frequency of 0, which would occur if raw frequency=1 (Log10(1)=0). Thus all freqquency data is slightly elevated, by 2

                            Case FrequencyUnits.PhonemeCount
                                'Phoneme frequencies are summed first (1*token count => total phoneme count in the corpus), Logs are calculated later
                                FrequencyValue = CurrentWord.RawWordTypeFrequency

                            Case Else
                                Throw New NotImplementedException

                        End Select
                        Me.AddProbabilityData(phonemeIndex, CurrentSound, FrequencyValue, True)

                    Else
                        'Adding probability to the word probability array
                        PP_Data.Add(Me.GetProbabilityData(phonemeIndex, CurrentSound))

                        'Adding the phonemes in the used-phonemes array
                        PP_Phonemes.Add(CurrentSound)
                    End If
                Next

                If TransitionDataDirection = PhonotacticTransitionDataDirections.RetrievePhonotacticProbabilities Then

                    'Storing retrieves probability data

                    Dim ZString As String = ""
                    If Do_z_ScoreTransformation = False Then

                        'No z-standardization
                        Select Case PhonemeCombinationLength
                            Case PhonemeCombinationLengths.MonoGramCalculation

                                CurrentWord.PSP = PP_Data
                                CurrentWord.PSP_Phonemes = PP_Phonemes

                                If PP_Data.Count > 0 Then
                                    CurrentWord.PSP_Sum = PP_Data.Sum
                                Else
                                    CurrentWord.PSP_Sum = 0
                                End If

                            Case PhonemeCombinationLengths.BiGramCalculation

                                CurrentWord.PSBP = PP_Data
                                CurrentWord.PSBP_Phonemes = PP_Phonemes

                                If PP_Data.Count > 0 Then
                                    CurrentWord.PSBP_Sum = PP_Data.Sum
                                Else
                                    CurrentWord.PSBP_Sum = 0
                                End If

                        End Select

                    Else

                        'z-Standardization
                        ZString = "_Z_Trans"

                        Select Case PhonemeCombinationLength
                            Case PhonemeCombinationLengths.MonoGramCalculation

                                CurrentWord.S_PSP = PP_Data
                                CurrentWord.S_PSP_Phonemes = PP_Phonemes

                                If PP_Data.Count > 0 Then
                                    CurrentWord.S_PSP_Average = PP_Data.Average
                                Else
                                    CurrentWord.S_PSP_Average = 0
                                End If

                            Case PhonemeCombinationLengths.BiGramCalculation

                                CurrentWord.S_PSBP = PP_Data
                                CurrentWord.S_PSBP_Phonemes = PP_Phonemes

                                If PP_Data.Count > 0 Then
                                    CurrentWord.S_PSBP_Average = PP_Data.Average
                                Else
                                    CurrentWord.S_PSBP_Average = 0
                                End If

                        End Select


                    End If

                    'Reporting lacking probability data to log file
                    If PP_Data.Count = 0 Then
                        SendInfoToLog(String.Join(" ", CurrentWord.BuildExtendedIpaArray), "WordsLackingPP_" & PhonemeCombinationLength.ToString & "_" & ZString)
                    End If

                End If

            Next

            If TransitionDataDirection = PhonotacticTransitionDataDirections.CreateTransitionMatrix Then
                'Calculates transitional probabilities after all frequency data is set
                CalculateTransitionalProbabilities()

                'Sorting transitional probabilities data
                SortProbabilityData()

                'Exporting transitional probabilities to txt file, without rounding
                ExportTransitionalProbabilityData(OutputFolder, OutputFileName,,,, True)
                ExportTransitionalProbabilityData(OutputFolder, OutputFileName & "_FullLines",,, True, True)

                'Exporting transitional probabilities to txt file, with rounded values
                ExportTransitionalProbabilityData(OutputFolder, OutputFileName & "_Rounded",,,, False)
                ExportTransitionalProbabilityData(OutputFolder, OutputFileName & "_Rounded" & "_FullLines",,, True, False)

            End If

            'Closing the progress display
            myProgress.Close()

        End Sub

        Private Sub CalculateTransitionalProbabilities(Optional ByVal AcceptedRoundingErrorSize As Double = 0.0001)

            SendInfoToLog("Initializing method " & Reflection.MethodInfo.GetCurrentMethod.Name)

            Dim DoFilterring As Boolean = True
            If DoFilterring = True Then

                Dim PhonemeThresholdValue As Decimal = 2
                Dim TransitionPhonemeThresholdValue As Integer = 2

                Dim PhonemesRemoved As Integer = 0

                SendInfoToLog("Filterring out unusual phoneme transition.")

                'Filterring out very uncommon phoneme transitions
                For Each StartPosition In TransitionalProbabilities

                    'Collecting phonemes to remove (which have a occurence count below PhonemeThresholdValue)
                    Dim PhonemesToRemove As New List(Of String)

                    For Each PhoneData In StartPosition.Value
                        If PhoneData.Value.OccurenceCount < TransitionPhonemeThresholdValue Then PhonemesToRemove.Add(PhoneData.Key)
                    Next

                    'Removing transition phonemes
                    For Each CurrentPhoneme In PhonemesToRemove
                        StartPosition.Value.Remove(CurrentPhoneme)
                        PhonemesRemoved += 1
                    Next
                Next

                SendInfoToLog("Finished filterring out unusual phoneme transition. Results: " & PhonemesRemoved & " phonemes/biphones were removed.")

            End If


            For Each StartPosition In TransitionalProbabilities
                StartPosition.Value.FrequencyDataSum = 0
                For Each PhoneData In StartPosition.Value

                    Select Case FrequencyUnit
                        Case FrequencyUnits.WordCount
                                'Logs were taken in the first step, no modification is needed

                        Case FrequencyUnits.PhonemeCount
                            'Raw phoneme frequencies were collected in the first step, now log 10 of these are summed
                            PhoneData.Value.FrequencyData = Math.Log10(PhoneData.Value.FrequencyData + 1) ' +1 is used to avoid that raw word type frequency of 1 gets a zero probability: (log10(1)= 0). All phoneme counts are slightly elevated, by 1.

                    End Select

                    'Summing
                    StartPosition.Value.FrequencyDataSum += PhoneData.Value.FrequencyData
                Next
            Next


            'Calculating transitional probabilies
            Dim ZeroProbabilityTransitionCount As Integer = 0
            For Each StartPosition In TransitionalProbabilities
                For Each PhoneData In StartPosition.Value

                    If StartPosition.Value.FrequencyDataSum = 0 Then
                        PhoneData.Value.TransitionalProbability = 0 'TransitionalProbability is set to zero if the frequency in the word list is 0 (Log10(RawFrequency of 0+1))
                        ZeroProbabilityTransitionCount += 1
                    Else
                        PhoneData.Value.TransitionalProbability = PhoneData.Value.FrequencyData / StartPosition.Value.FrequencyDataSum
                    End If
                Next
            Next

            'NB. Any zero probability transitions caused by log10(raw frequency=0 +1) could be deleted here, as is done in the Witte PP type!

            SendInfoToLog("Finished calculating transitional probabilities. " & ZeroProbabilityTransitionCount & " probabilities of 0 detected.")

            'Checking that the transitional probabilies for all transitional possibilities for each source phoneme adds up to 1

            SendInfoToLog("Checking for errors in calculations of transitional probabilities.")

            Dim ErrorsInTransitionalProbabilitiyCalculaltions As Integer = 0
            For Each StartPosition In TransitionalProbabilities
                StartPosition.Value.TransitionProbabilitySum = 0
                For Each PhoneData In StartPosition.Value
                    StartPosition.Value.TransitionProbabilitySum += PhoneData.Value.TransitionalProbability
                Next

                If Not StartPosition.Value.TransitionProbabilitySum > (1 - AcceptedRoundingErrorSize) And
                                StartPosition.Value.TransitionProbabilitySum < (1 + AcceptedRoundingErrorSize) Then
                    ErrorsInTransitionalProbabilitiyCalculaltions += 1
                End If
            Next

            SendInfoToLog("   Checked for errors in calculations of transitional probabilities completed successfully. Results: " & ErrorsInTransitionalProbabilitiyCalculaltions & " probability values do not add up to 1 +/-" & AcceptedRoundingErrorSize)

            SendInfoToLog("   Doing z-score transformation.")

            'Determining the highest phonotactic probability for each source phoneme
            For Each StartPosition In TransitionalProbabilities

                'Collecting probability data
                Dim TransformationData As New List(Of Double)

                For Each PhoneData In StartPosition.Value
                    TransformationData.Add(PhoneData.Value.TransitionalProbability)
                Next

                'Transforming the data
                MathMethods.Standardization(TransformationData)

                'Putting the transformed data back into the PhoneData-instances, but now into the Z_Transformed_TransitionalProbability variable
                Dim CurrentDataIndex As Integer = 0
                For Each PhoneData In StartPosition.Value
                    PhoneData.Value.Z_Transformed_TransitionalProbability = TransformationData(CurrentDataIndex)
                    CurrentDataIndex += 1
                Next
            Next

            SendInfoToLog("Method " & Reflection.MethodInfo.GetCurrentMethod.Name & " completed successfully.")

        End Sub

    ''' <summary>
    ''' Loads the probability data from a "Full-Lines" type exported Vitevitch/Luce/Storkel type phonotactic probability data file.
    ''' </summary>
    ''' <param name="FilePath"></param>
    Public Sub LoadProbabilityDataFromFile(Optional ByRef FilePath As String = "")

        'This sub loads probability data from file, to enable phonotactic probability calculations of words without running the word list analysis part
        TransitionalProbabilities = New PhonoTacticProbabilities()

        Dim InputLines() As String

        If FilePath = "" Then
            Dim dataString As String = ""

            Select Case PhonemeCombinationLength
                Case PhonemeCombinationLengths.MonoGramCalculation
                    dataString = My.Resources.PSP_Matrix_FullLines
                Case PhonemeCombinationLengths.BiGramCalculation
                    dataString = My.Resources.PSBP_Matrix_FullLines
            End Select

            dataString = dataString.Replace(vbCrLf, vbLf)
            InputLines = dataString.Split(vbLf)

        Else
            InputLines = System.IO.File.ReadAllLines(FilePath, Text.Encoding.UTF8)
        End If

        For LineIndex = 1 To InputLines.Length - 1 'Skipping heading line

            If InputLines(LineIndex).Trim = "" Then Continue For

            'Reading all data in the current line
            Dim LineSplit() As String = InputLines(LineIndex).Split(vbTab)

            'Check line split length
            If LineSplit.Length < 6 Then MsgBox("Error reading line: " & InputLines(LineIndex))

            Dim PhonemePosition As Integer = LineSplit(0)
            Dim PhonemeString As String = LineSplit(1)
            Dim FrequencyData As String = LineSplit(2)
            Dim PhontacticProbability As Double = LineSplit(3)
            Dim StandardizedPhontacticProbability As Double = LineSplit(4)
            Dim Occurences As Double = LineSplit(5)

            'Checking that the appropriate locvation in the structure exists, or adds it otherwise
            If Not TransitionalProbabilities.ContainsKey(PhonemePosition) Then TransitionalProbabilities.Add(PhonemePosition, New PhonemeData)
            If Not TransitionalProbabilities(PhonemePosition).ContainsKey(PhonemeString) Then TransitionalProbabilities(PhonemePosition).Add(PhonemeString, New PhonemeData.ProbabilityData)

            'Adding the probabililty data
            TransitionalProbabilities(PhonemePosition)(PhonemeString).FrequencyData = FrequencyData
            TransitionalProbabilities(PhonemePosition)(PhonemeString).TransitionalProbability = PhontacticProbability
            TransitionalProbabilities(PhonemePosition)(PhonemeString).Z_Transformed_TransitionalProbability = StandardizedPhontacticProbability
            TransitionalProbabilities(PhonemePosition)(PhonemeString).OccurenceCount = Occurences

        Next

    End Sub


    Public Sub ExportTransitionalProbabilityData(Optional ByRef saveDirectory As String = "", Optional ByRef saveFileName As String = "SyllableBasedPhonemeTransitionalProbabilities",
                                                Optional BoxTitle As String = "Choose location to store the transitional probability export file...",
                                                         Optional SkipZeroProbabilityTransitions As Boolean = True,
                                                          Optional ByVal OutputFullLines As Boolean = False, Optional ByVal SkipRounding As Boolean = False)

            Try


                SendInfoToLog("Attempts to save transitional probability data to .txt file.")

                'Choosing file location
                Dim filepath As String = ""
                'Ask the user for file path if not incomplete file path is given
                If saveDirectory = "" Or saveFileName = "" Then
                    filepath = GetSaveFilePath(saveDirectory, saveFileName, {"txt"}, BoxTitle)
                Else
                    filepath = Path.Combine(saveDirectory, saveFileName & ".txt")
                    If Not Directory.Exists(Path.GetDirectoryName(filepath)) Then Directory.CreateDirectory(Path.GetDirectoryName(filepath))
                End If

                'Save it to file
                Dim writer As New StreamWriter(filepath, False, Text.Encoding.UTF8)

                'Witing heading
                writer.WriteLine("Position" & vbTab & "Phoneme/s" & vbTab & "FrequencyData" & vbTab & "TransitionalProbability" & vbTab & "Z-TransformedTransitionalProbability" & vbTab & "Occurences" & vbCrLf)

                If OutputFullLines = False Then

                    For Each StartPosition In TransitionalProbabilities
                        writer.WriteLine(StartPosition.Key & vbTab & Rounding(StartPosition.Value.FrequencyDataSum, , 4, SkipRounding) & vbTab & Rounding(StartPosition.Value.TransitionProbabilitySum,, 6, SkipRounding))
                        For Each PhoneData In StartPosition.Value
                            If SkipZeroProbabilityTransitions = True And PhoneData.Value.TransitionalProbability = 0 Then Continue For
                            Select Case PhonemeCombinationLength
                                Case PhonemeCombinationLengths.MonoGramCalculation
                                    writer.WriteLine(vbTab & PhoneData.Key.ToString & vbTab & Rounding(PhoneData.Value.FrequencyData,, 4, SkipRounding) & vbTab &
                                                         Rounding(PhoneData.Value.TransitionalProbability, , 4, SkipRounding) & vbTab & Rounding(PhoneData.Value.Z_Transformed_TransitionalProbability, , 4, SkipRounding) & vbTab & PhoneData.Value.OccurenceCount)
                                Case PhonemeCombinationLengths.BiGramCalculation
                                    writer.WriteLine(vbTab & PhoneData.Key.ToString & vbTab & Rounding(PhoneData.Value.FrequencyData,, 4, SkipRounding) & vbTab &
                                                         Rounding(PhoneData.Value.TransitionalProbability,, 6, SkipRounding) & vbTab & Rounding(PhoneData.Value.Z_Transformed_TransitionalProbability,, 6, SkipRounding) & vbTab & PhoneData.Value.OccurenceCount)
                            End Select

                        Next
                        writer.WriteLine() 'Inserting an empty line before the next section

                    Next


                Else

                    For Each StartPosition In TransitionalProbabilities
                        For Each PhoneData In StartPosition.Value
                            If SkipZeroProbabilityTransitions = True And PhoneData.Value.TransitionalProbability = 0 Then Continue For
                            Select Case PhonemeCombinationLength
                                Case PhonemeCombinationLengths.MonoGramCalculation
                                    writer.WriteLine(StartPosition.Key & vbTab & PhoneData.Key.ToString & vbTab & Rounding(PhoneData.Value.FrequencyData,, 4, SkipRounding) &
                                                         vbTab & Rounding(PhoneData.Value.TransitionalProbability, , 4, SkipRounding) & vbTab & Rounding(PhoneData.Value.Z_Transformed_TransitionalProbability, , 4, SkipRounding) & vbTab & PhoneData.Value.OccurenceCount)
                                Case PhonemeCombinationLengths.BiGramCalculation
                                    writer.WriteLine(StartPosition.Key & vbTab & PhoneData.Key.ToString & vbTab & Rounding(PhoneData.Value.FrequencyData,, 4, SkipRounding) &
                                                         vbTab & Rounding(PhoneData.Value.TransitionalProbability,, 6, SkipRounding) & vbTab & Rounding(PhoneData.Value.Z_Transformed_TransitionalProbability,, 6, SkipRounding) & vbTab & PhoneData.Value.OccurenceCount)
                            End Select
                        Next

                    Next

                End If

                writer.Close()

                SendInfoToLog("   Transitional probability data were successfully saved to .txt file: " & filepath)

            Catch ex As Exception

            End Try


        End Sub

        ''' <summary>
        ''' Sorts the output data in descending order determined by the values of TransitionalProbability or TransitionPredictability depending on the selected calculation type.
        ''' </summary>
        Private Sub SortProbabilityData()

            'Sorting transition phonemes according to their transitional probability (source phonemes are already sorted as they are in a SortedList)
            Dim MySortOrder As SortOrder = SortOrder.Descending
            For Each StartPosition In TransitionalProbabilities
                StartPosition.Value.SortOutputData()
            Next

        End Sub

        Private Class PhonoTacticProbabilities
            Inherits Dictionary(Of Integer, PhonemeData) 'Holding the phoneme combination start index, and a list of phonemes/phoneme combinations starting on that index
        End Class


        Public Class PhonemeData
            Inherits SortedList(Of String, ProbabilityData)

            'Public Property FrequencyData As Double
            Public Property FrequencyDataSum As Double
            Public Property TransitionProbabilitySum As Double
            Public Property HighestTransitionalProbability As Double

            Public Class ProbabilityData
                Public Property FrequencyData As Double
                Public Property TransitionalProbability As Double
                Public Property Z_Transformed_TransitionalProbability As Double
                Public Property OccurenceCount As Integer
            End Class

            Public Sub SortOutputData()

                'Putting all data to sort in a list
                Dim newSortList As New List(Of SortData)
                For Each CurrentItem In Me
                    Dim newSortData As New SortData()
                    newSortData.MyKey = CurrentItem.Key
                    newSortData.MyValue = CurrentItem.Value
                    newSortList.Add(newSortData)
                Next

                'Sorting the list
                newSortList.Sort()

                'Reversing the order
                newSortList.Reverse()

                'Clearing Me
                Me.Clear()

                'Putting the data back into Me
                For Each CurrentItem In newSortList
                    Me.Add(CurrentItem.MyKey, CurrentItem.MyValue)
                Next

            End Sub

            Private Class SortData
                Implements IComparable

                Property MyKey As String
                Property MyValue As ProbabilityData

                Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo

                    If Not TypeOf (obj) Is SortData Then
                        Throw New ArgumentException()
                    Else

                        Dim tempData As SortData = DirectCast(obj, SortData)

                        If Me.MyValue.TransitionalProbability < tempData.MyValue.TransitionalProbability Then
                            Return -1
                        ElseIf Me.MyValue.TransitionalProbability = tempData.MyValue.TransitionalProbability Then
                            Return 0
                        Else
                            Return 1
                        End If

                    End If
                End Function
            End Class
        End Class

    End Class


